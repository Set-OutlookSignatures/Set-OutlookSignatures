function Format-PostalAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][hashtable]$Components,

        # Optional individual options (will be merged with -Options if provided)
        [string]$Country,
        [switch]$Abbreviate,
        [string]$AddressTemplate,
        [Nullable[bool]]$OnlyAddress,

        # Perl-like options hashtable { country, abbreviate, address_template, only_address }
        [hashtable]$Options
    )

    # All logic is delegated to the internal, original function name for simplicity
    return (Format-PostalAddressInternal -Components $Components -Country $Country -Abbreviate:$Abbreviate -AddressTemplate $AddressTemplate -OnlyAddress $OnlyAddress -Options $Options)
}


# ---------------------------------------------------------------------------------
# Internal functions and variables (Not exported, only available within the module)
# ---------------------------------------------------------------------------------

# The original function is renamed to keep the module entry point clean
function Format-PostalAddressInternal {
    # Original Format-PostalAddress logic goes here...
    param(
        [Parameter(Mandatory, Position = 0)][hashtable]$Components,

        # Optional individual options (will be merged with -Options if provided)
        [string]$Country,
        [switch]$Abbreviate,
        [string]$AddressTemplate,
        [Nullable[bool]]$OnlyAddress,

        # Perl-like options hashtable { country, abbreviate, address_template, only_address }
        [hashtable]$Options
    )

    if ($null -eq $Components) { return $null }
    # Shallow clone is sufficient — components are flat (scalar) values
    $rh_components = $Components.Clone()

    # Reset final components at start
    $Script:AddrFmt.FinalComponents = $null

    # Merge options (shallow clone is sufficient — options are flat scalar values)
    $opt = @{}
    if ($Options) { $opt = $Options.Clone() }
    if ($PSBoundParameters.ContainsKey('Country')) { $opt['country'] = $Country }
    if ($PSBoundParameters.ContainsKey('Abbreviate')) { $opt['abbreviate'] = [int][bool]$Abbreviate }
    if ($PSBoundParameters.ContainsKey('AddressTemplate')) { $opt['address_template'] = $AddressTemplate }
    if ($PSBoundParameters.ContainsKey('OnlyAddress')) { $opt['only_address'] = $OnlyAddress }

    # Determine country code
    $cc = $null
    if ($opt.ContainsKey('country') -and $opt['country']) {
        $cc = $opt['country']
        $rh_components['country_code'] = $cc
        Set-DistrictAlias -CountryCode $cc
    } else {
        $cc = Determine-CountryCode -Components $rh_components
        if ($cc) {
            $rh_components['country_code'] = $cc
            Set-DistrictAlias -CountryCode $cc
        }
    }

    # Abbreviate flag
    $abbrv = 0
    if ($opt.ContainsKey('abbreviate')) { $abbrv = [int]$opt['abbreviate'] }

    # OnlyAddress (call-level overrides object-level)
    $oa = $Script:AddrFmt.OnlyAddress
    if ($opt.ContainsKey('only_address')) { $oa = [bool]$opt['only_address'] }

    # 1) Sanity-cleaning
    Invoke-SanityCleaning -Components $rh_components

    # 2) Apply component aliases into primary types
    Resolve-PrimaryAliases -Components $rh_components

    # 3) Determine template (robust, with layered fallback; last-resort built-in)
    $templateText = $null
    $rh_config = $null
    $defaultConfig = $null
    if ($Script:AddrFmt.Templates.ContainsKey('default')) {
        $defaultConfig = $Script:AddrFmt.Templates['default']
    }

    if ($cc) {
        $rh_config = $Script:AddrFmt.Templates[$cc]
    }
    if (-not $rh_config) { $rh_config = $defaultConfig }

    if ($opt.ContainsKey('address_template') -and $opt['address_template']) {
        $templateText = [string]$opt['address_template']
    }
    if ([string]::IsNullOrWhiteSpace($templateText) -and $rh_config -and $rh_config.address_template) {
        $templateText = [string]$rh_config.address_template
    }
    # If still empty, try country fallback_template
    if ([string]::IsNullOrWhiteSpace($templateText) -and $rh_config -and $rh_config.fallback_template) {
        $templateText = [string]$rh_config.fallback_template
    }
    # If still empty, try default address_template
    if ([string]::IsNullOrWhiteSpace($templateText) -and $defaultConfig -and $defaultConfig.address_template) {
        $templateText = [string]$defaultConfig.address_template
    }
    # If still empty and minimal components are missing, try default fallback_template
    if ([string]::IsNullOrWhiteSpace($templateText) -and $defaultConfig -and $defaultConfig.fallback_template) {
        $templateText = [string]$defaultConfig.fallback_template
    }

    # Last-resort: keep the formatter functional even if the config is pathological
    if ([string]::IsNullOrWhiteSpace($templateText)) {
        $templateText = @'
{{#first}}
{{{attention}}}
{{/first}}
{{#first}}
{{{house}}}{{{house_number}}} {{{road}}}
{{/first}}
{{#first}}
{{{postcode}}} {{{city}}}
{{/first}}
{{#first}}
{{{state}}}
{{/first}}
{{#first}}
{{{country}}}
{{/first}}
'@
        Warn-If 'Using built-in last-resort template because config provided no address/fallback template.'
    }

    # Prefer configured fallback when minimal components are missing
    $haveMinimal = Test-MinimalComponents -Components $rh_components
    if (-not $haveMinimal -and $rh_config -and $rh_config.fallback_template) {
        $templateText = [string]$rh_config.fallback_template
    } elseif (-not $haveMinimal -and $defaultConfig -and $defaultConfig.fallback_template) {
        $templateText = [string]$defaultConfig.fallback_template
    }

    # 4) Fix country hacks
    Fix-Country -Components $rh_components

    # 5) Apply replacements (pre-render). Prefer pre-compiled rules.
    if ($rh_config) {
        if ($rh_config.ContainsKey('_compiled_replace')) {
            Apply-Replacements -Components $rh_components -Rules $rh_config['_compiled_replace']
        } elseif ($rh_config.replace) {
            Apply-Replacements -Components $rh_components -Rules $rh_config.replace
        }
    }

    # 6) Add state/county codes
    Add-StateCode -Components $rh_components | Out-Null
    Add-CountyCode -Components $rh_components | Out-Null

    # 7) Unknown components -> attention (unless only_address)
    if (-not $oa) {
        $unknown = Find-UnknownComponents -Components $rh_components
        if ($unknown.Count -gt 0) {
            $sorted = $unknown | Sort-Object
            $vals = foreach ($k in $sorted) { $rh_components[$k] }
            $rh_components['attention'] = ($vals -join ', ')
        }
    }

    # 8) Abbreviate if requested
    if ($abbrv) {
        $tmp = Invoke-Abbreviate -Components $rh_components
        if ($tmp) { $rh_components = $tmp }
    }

    # 9) Prepare template (replace lambda) and compile to AST (cached)
    $templateText = Replace-TemplateLambdas -TemplateText $templateText
    $compiledTpl = Compile-MustacheTemplate -Template $templateText

    # 10) Render
    $rendered = Render-CompiledMustache -Nodes $compiledTpl -Context $rh_components
    $rendered = Evaluate-TemplateLambdas -Rendered $rendered

    # 11) Postformat replacements and duplicate removal (use pre-compiled rules if available)
    $pfRules = $null
    if ($rh_config) {
        if ($rh_config.ContainsKey('_compiled_postformat')) {
            $pfRules = $rh_config['_compiled_postformat']
        } else {
            $pfRules = $rh_config.postformat_replace
        }
    }
    $rendered = Invoke-Postformat -Text $rendered -Rules $pfRules

    # 12) Line-by-line clean (using \s, no multiline \h)
    $rendered = Invoke-Clean -Text $rendered

    # 13) If empty and only one component exists, use that value (Perl behavior)
    if ($rendered.Length -eq 0) {
        $keys = $rh_components.Keys
        if ($keys.Count -eq 1) {
            $k = $keys | Select-Object -First 1
            $rendered = [string]$rh_components[$k]
        }
    }

    # 14) Set final components
    $Script:AddrFmt.FinalComponents = $rh_components

    return $Script:RxNewline.Replace($rendered, $Script:NewlineEnv)
}


$Script:AddrFmt = @{
    Templates             = @{}   # country templates + default
    ComponentAliases      = @{}   # primary -> [aliases]
    Component2Type        = @{}   # alias -> primary
    OrderedComponents     = @()   # [ primary1, alias1a, alias1b, primary2, ... ]
    HKnown                = @{}   # set of known component keys
    State_Codes           = @{}   # country_code -> map
    County_Codes          = @{}   # country_code -> map
    Country2Lang          = @{}   # country_code -> "en,de"
    Abbreviations         = @{}   # lang -> component -> (long -> short)
    SetDistrictAlias      = @{}   # cache to avoid rework
    FinalComponents       = $null # set after successful Format-PostalAddress

    # Derived caches built once at init (mirror Geo::Address::Formatter):
    State_Codes_Reverse   = @{}   # cc -> { UPPERCASE_NAME -> code }
    State_Codes_Name      = @{}   # cc -> { code -> default_name }
    County_Codes_Reverse  = @{}   # cc -> { UPPERCASE_NAME -> code }
    County_Codes_Name     = @{}   # cc -> { code -> default_name }
    CompiledAbbreviations = @{}   # lang -> compName -> [ {Re; Short} ]

    ShowWarnings          = $true
    OnlyAddress           = $false

    ConfPath              = $null
}

# Countries where plain "district" should be treated as "neighbourhood" (small district)
$Script:SmallDistrict = @{
    'BR' = 1; 'CR' = 1; 'ES' = 1; 'NI' = 1; 'PY' = 1; 'RO' = 1; 'TG' = 1; 'TM' = 1; 'XK' = 1
}

# ------------------------------
# Pre-compiled, cached regex objects (created once at module load time)
# ------------------------------
$Script:RxOpts = [System.Text.RegularExpressions.RegexOptions]::Compiled

# Mustache renderer
$Script:RxSection = [regex]::new('(?s){{\#\s*([\w\.\-]+)\s*}}(.*?){{/\s*\1\s*}}', $Script:RxOpts)
$Script:RxInvertedSection = [regex]::new('(?s){{\^\s*([\w\.\-]+)\s*}}(.*?){{/\s*\1\s*}}', $Script:RxOpts)
$Script:RxTriple = [regex]::new('(?s){{{\s*([\w\.\-]+)\s*}}}', $Script:RxOpts)
$Script:RxAmp = [regex]::new('(?s){{&\s*([\w\.\-]+)\s*}}', $Script:RxOpts)
$Script:RxSimple = [regex]::new('(?s){{\s*([\w\.\-]+)\s*}}', $Script:RxOpts)

# Template lambdas
$Script:RxFirst = [regex]::new('(?s){{\#first}}(.+?){{\/first}}', $Script:RxOpts)
$Script:RxFirstSentinel = [regex]::new('(?s)FIRSTSTART\s*(.+?)\s*FIRSTEND', $Script:RxOpts)
$Script:RxSelectFirst = [regex]::new('\s*(?:\|\||\r?\n)\s*', $Script:RxOpts)

# Generic
$Script:RxNewline = [regex]::new('\r?\n', $Script:RxOpts)
$Script:RxLeadingWS = [regex]::new('^\s+', $Script:RxOpts)
$Script:RxTrailingWS = [regex]::new('\s+$', $Script:RxOpts)
$Script:NewlineEnv = [System.Environment]::NewLine

# Invoke-Postformat
$Script:RxPF_Dash = [regex]::new('^- ', $Script:RxOpts)
$Script:RxPF_CommaCom = [regex]::new(',\s*,', $Script:RxOpts)
$Script:RxPF_SpcCommaSpc = [regex]::new('\s+,\s+', $Script:RxOpts)
$Script:RxPF_MultiSpace = [regex]::new('\s\s+', $Script:RxOpts)
$Script:RxPF_LeadComma = [regex]::new('^,', $Script:RxOpts)
$Script:RxPF_MultiComma = [regex]::new(',,+', $Script:RxOpts)
$Script:RxPF_TrailComma = [regex]::new(',$', $Script:RxOpts)

# Invoke-Clean
$Script:RxCl_LBracket = [regex]::new('^[\[\{,\s]+', $Script:RxOpts)
$Script:RxCl_RBracket = [regex]::new('[\}\],\s]+$', $Script:RxOpts)
$Script:RxCl_LComma = [regex]::new('^\s*,+\s*', $Script:RxOpts)
$Script:RxCl_RComma = [regex]::new('\s*,+\s*$', $Script:RxOpts)
$Script:RxCl_CommaCommas = [regex]::new(',\s*,+', $Script:RxOpts)
$Script:RxCl_CommaSpace = [regex]::new('\s*,\s*', $Script:RxOpts)
$Script:RxCl_MultiSpace = [regex]::new('\s{2,}', $Script:RxOpts)
$Script:RxCl_HtmlApos = [regex]::new('&#39;', $Script:RxOpts)

# Sanity cleaning
$Script:RxSC_PostcodeSemi = [regex]::new('^\d+;\d+$', $Script:RxOpts)
$Script:RxSC_PostcodeFive = [regex]::new('^(\d{5}),\d{5}', $Script:RxOpts)
$Script:RxSC_HasWord = [regex]::new('\w', $Script:RxOpts)
$Script:RxSC_Url = [regex]::new('(?s)https?://', $Script:RxOpts)

# Apply-Replacements compiled regex cache (keyed by pattern string)
$Script:ApplyRxCache = @{}

# Compiled mustache template cache (keyed by template_text after lambda replacement)
$Script:CompiledTemplateCache = @{}

# Token regex used to compile a mustache template into an AST.
# Order matters: longer/more-specific alternatives come first.
$Script:RxMustacheTokens = [regex]::new(
    '(?s)\{\{\{\s*[\w\.\-]+\s*\}\}\}|\{\{[#^/&]\s*[\w\.\-]+\s*\}\}|\{\{\s*[\w\.\-]+\s*\}\}',
    $Script:RxOpts)

# ------------------------------
# Utility helpers
# ------------------------------

function Warn-If {
    param([string]$Message)
    if ($Script:AddrFmt.ShowWarnings) { Write-Warning $Message }
}

function Throw-IfNullOrMissingPath {
    param([string]$Path, [string]$What)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        throw "Missing $What at '$Path'."
    }
}

function Clone-Deep {
    param([Parameter(ValueFromPipeline)][object]$InputObject)
    process {
        if ($null -eq $InputObject) { return $null }
        # Scalars (most common): strings, numbers, bools — return as-is (immutable)
        if ($InputObject -is [string] -or $InputObject -is [System.ValueType]) {
            return $InputObject
        }
        if ($InputObject -is [System.Collections.IDictionary]) {
            $clone = @{}
            foreach ($k in $InputObject.Keys) {
                $v = $InputObject[$k]
                if ($null -eq $v -or $v -is [string] -or $v -is [System.ValueType]) {
                    $clone[$k] = $v
                } else {
                    $clone[$k] = Clone-Deep $v
                }
            }
            return $clone
        } elseif ($InputObject -is [System.Collections.IEnumerable]) {
            $list = New-Object System.Collections.Generic.List[object]
            foreach ($i in $InputObject) {
                if ($null -eq $i -or $i -is [string] -or $i -is [System.ValueType]) {
                    $list.Add($i)
                } else {
                    $list.Add((Clone-Deep $i))
                }
            }
            return , ($list.ToArray())
        } else {
            return $InputObject
        }
    }
}

# HTML escape for {{var}} (not used for {{{var}}} / {{& var}})
function ConvertTo-HtmlEscapedText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $t = $Text
    $t = $t -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    $t = $t -replace '"', '&quot;'
    $t = $t -replace "'", '&#39;'
    return $t
}

# --- YAML helpers: merging + PSCustomObject -> hashtable conversion ---

function ConvertTo-Hashtable {
    param([Parameter(Mandatory)][object]$InputObject)
    if ($null -eq $InputObject) { return $null }
    # Fast scalar path
    if ($InputObject -is [string] -or $InputObject -is [System.ValueType]) {
        return $InputObject
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $InputObject.Keys) {
            $v = $InputObject[$k]
            if ($null -eq $v -or $v -is [string] -or $v -is [System.ValueType]) {
                $h[$k] = $v
            } else {
                $h[$k] = ConvertTo-Hashtable -InputObject $v
            }
        }
        return $h
    } elseif ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $v = $p.Value
            if ($null -eq $v -or $v -is [string] -or $v -is [System.ValueType]) {
                $h[$p.Name] = $v
            } else {
                $h[$p.Name] = ConvertTo-Hashtable -InputObject $v
            }
        }
        return $h
    } elseif ($InputObject -is [System.Collections.IEnumerable]) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($i in $InputObject) {
            if ($null -eq $i -or $i -is [string] -or $i -is [System.ValueType]) {
                $list.Add($i)
            } else {
                $list.Add((ConvertTo-Hashtable -InputObject $i))
            }
        }
        return , ($list.ToArray())
    } else {
        return $InputObject
    }
}

function Merge-Hashtables {
    param(
        [Parameter(Mandatory)][hashtable]$Base,
        [Parameter(Mandatory)][hashtable]$Overlay
    )
    $result = @{}
    foreach ($k in $Base.Keys) { $result[$k] = Clone-Deep $Base[$k] }
    foreach ($k in $Overlay.Keys) {
        if ($result.ContainsKey($k) -and ($result[$k] -is [hashtable]) -and ($Overlay[$k] -is [hashtable])) {
            $result[$k] = Merge-Hashtables -Base $result[$k] -Overlay $Overlay[$k]
        } else {
            $result[$k] = Clone-Deep $Overlay[$k]
        }
    }
    return $result
}


# YAML loader.
#
# powershell-yaml is correct but slow on first load: every YAML node goes
# through several PowerShell function calls (Convert-YamlDocumentToPSObject ->
# Convert-YamlMappingToHashtable -> Convert-ValueToProperType, switch on
# GetType().FullName). For ~10k nodes (worldwide.yaml) that overhead dominates
# init time. We bypass it by calling YamlDotNet directly (the .NET assembly
# powershell-yaml already loaded into the AppDomain on its own import) and
# walking the resulting YamlStream with one tight recursive function.
#
# Assumptions verified against the OpenCageData address-formatting conf:
#   * No typed scalars (no booleans, integers, floats, custom tags).
#   * All mapping keys are string scalars.
# Untyped null-like scalars ("", "~", "null") are coerced to $null to match
# powershell-yaml's behaviour. Everything else is returned as a string.
$Script:_YamlScalarT = $null
$Script:_YamlMapT = $null
$Script:_YamlSeqT = $null
$Script:_YamlParserT = $null
$Script:_YamlMergingParserT = $null
$Script:_YamlStreamT = $null

function Initialize-YamlDotNetTypes {
    # Re-enter only when *all* cached type vars are set, so a partially-failed
    # init (e.g. one type missing) doesn't leave a stale half-initialised state.
    if ($null -ne $Script:_YamlScalarT -and `
            $null -ne $Script:_YamlMapT -and `
            $null -ne $Script:_YamlSeqT -and `
            $null -ne $Script:_YamlParserT -and `
            $null -ne $Script:_YamlMergingParserT -and `
            $null -ne $Script:_YamlStreamT) { return }

    $asm = $null
    foreach ($a in [System.AppDomain]::CurrentDomain.GetAssemblies()) {
        if ($a.GetName().Name -eq 'YamlDotNet') { $asm = $a; break }
    }
    if (-not $asm) {
        # Force powershell-yaml to load YamlDotNet into the AppDomain.
        $null = ConvertFrom-Yaml -Yaml ''
        foreach ($a in [System.AppDomain]::CurrentDomain.GetAssemblies()) {
            if ($a.GetName().Name -eq 'YamlDotNet') { $asm = $a; break }
        }
    }
    if (-not $asm) { throw 'YamlDotNet assembly not loaded; powershell-yaml import failed.' }

    # Resolve *all* required types into locals first; only commit to Script
    # scope once every lookup succeeds. This way a missing type throws
    # cleanly without leaving partial state behind.
    $needed = [ordered]@{
        '_YamlScalarT'        = 'YamlDotNet.RepresentationModel.YamlScalarNode'
        '_YamlMapT'           = 'YamlDotNet.RepresentationModel.YamlMappingNode'
        '_YamlSeqT'           = 'YamlDotNet.RepresentationModel.YamlSequenceNode'
        '_YamlParserT'        = 'YamlDotNet.Core.Parser'
        '_YamlMergingParserT' = 'YamlDotNet.Core.MergingParser'
        '_YamlStreamT'        = 'YamlDotNet.RepresentationModel.YamlStream'
    }
    $resolved = @{}
    foreach ($var in $needed.Keys) {
        $tn = $needed[$var]
        $tt = $asm.GetType($tn)
        if (-not $tt) {
            throw "YamlDotNet type '$tn' not found in assembly $($asm.GetName().Name) v$($asm.GetName().Version) (this YamlDotNet build is incompatible with the fast loader)."
        }
        $resolved[$var] = $tt
    }
    foreach ($var in $resolved.Keys) {
        Set-Variable -Scope Script -Name $var -Value $resolved[$var]
    }
}

function Convert-YamlNodeFast {
    param([object]$Node)

    $t = $Node.GetType()

    if ($t -eq $Script:_YamlScalarT) {
        # Tag handling intentionally mirrors powershell-yaml: coerce to string
        # via [string] (works whether YamlDotNet exposes Tag as a string or as
        # a TagName struct -- ToString() yields the textual tag in both cases).
        $tag = [string]$Node.Tag
        if ($tag -eq 'tag:yaml.org,2002:null') { return $null }
        $v = $Node.Value
        # Plain (unquoted) null-like scalars become $null; quoted strings
        # ("", "''", '""') keep their literal value because Style != Plain.
        if ($null -ne $v -and $Node.Style -eq 'Plain') {
            if ($v.Length -eq 0 -or $v -eq '~' -or $v -eq 'null' -or $v -eq 'Null' -or $v -eq 'NULL') {
                return $null
            }
        }
        return $v
    }

    if ($t -eq $Script:_YamlMapT) {
        $h = @{}
        foreach ($kv in $Node.Children) {
            $kNode = $kv.Key
            if ($kNode.GetType() -eq $Script:_YamlScalarT) {
                $key = $kNode.Value
            } else {
                $key = Convert-YamlNodeFast $kNode
            }
            $h[$key] = Convert-YamlNodeFast $kv.Value
        }
        return $h
    }

    if ($t -eq $Script:_YamlSeqT) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $Node.Children) {
            $list.Add((Convert-YamlNodeFast $item))
        }
        # Comma prevents PowerShell from unrolling the list into the pipeline.
        return , $list
    }

    return $Node
}

function Import-Yaml {
    param([Parameter(Mandatory)][string]$Path)

    Throw-IfNullOrMissingPath -Path $Path -What 'YAML file'
    Initialize-YamlDotNetTypes

    # Stream the file straight into YamlDotNet rather than allocating a
    # ~500KB string + StringReader copy. Parser/MergingParser/YamlStream
    # consume the StreamReader incrementally.
    $reader = [System.IO.StreamReader]::new($Path, [System.Text.Encoding]::UTF8)
    $stream = $Script:_YamlStreamT::new()
    try {
        $parser = $Script:_YamlParserT::new($reader)
        $parser = $Script:_YamlMergingParserT::new($parser)
        $stream.Load($parser)
    } finally {
        $reader.Dispose()
    }

    $docCount = $stream.Documents.Count
    if ($docCount -eq 0) { return $null }
    if ($docCount -eq 1) {
        return Convert-YamlNodeFast $stream.Documents[0].RootNode
    }
    # Multiple-document YAML: match powershell-yaml -AllDocuments shape, which
    # is an [object[]] of roots emitted to the pipeline (no unary comma, so
    # downstream pipeline consumers see one item per doc, exactly as before).
    $results = New-Object 'object[]' $docCount
    for ($i = 0; $i -lt $docCount; $i++) {
        $results[$i] = Convert-YamlNodeFast $stream.Documents[$i].RootNode
    }
    return $results
}

# Resolve "a.b.c" in nested hashtables
function Get-ContextValue {
    param([hashtable]$Context, [string]$KeyPath)
    if ($null -eq $Context) { return $null }
    if (-not $KeyPath) { return $null }
    # Fast path: simple (non-dotted) key — by far the most common case
    if ($KeyPath.IndexOf('.') -lt 0) {
        if ($Context.ContainsKey($KeyPath)) { return $Context[$KeyPath] }
        return $null
    }
    $node = $Context
    foreach ($part in $KeyPath.Split('.')) {
        if ($node -is [System.Collections.IDictionary]) {
            if ($node.Contains($part)) {
                $node = $node[$part]
            } else { return $null }
        } else { return $null }
    }
    return $node
}

function Test-Truthy {
    param([object]$Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [string]) { return -not [string]::IsNullOrWhiteSpace($Value) }
    if ($Value -is [System.Collections.IEnumerable]) {
        foreach ($x in $Value) { return $true }
        return $false
    }
    return $true
}

# ------------------------------
# Mustache-like renderer (subset used by templates)
# ------------------------------
function Render-Mustache {
    param(
        [Parameter(Mandatory)][string]$Template,
        [Parameter(Mandatory)][hashtable]$Context
    )

    $output = $Template

    # Quick exit: nothing to do if no mustache tag present
    if ($output.IndexOf('{{') -lt 0) { return $output }

    # Expand regular sections (nested same-name not supported; nested different names ok)
    $m = $Script:RxSection.Match($output)
    while ($m.Success) {
        $key = $m.Groups[1].Value
        $inner = $m.Groups[2].Value
        $val = Get-ContextValue -Context $Context -KeyPath $key

        $replacement = ''
        if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
            $acc = ''
            foreach ($item in $val) {
                if ($item -is [System.Collections.IDictionary]) {
                    $child = $Context.Clone()
                    foreach ($k in $item.Keys) { $child[$k] = $item[$k] }
                    $acc += Render-Mustache -Template $inner -Context $child
                } else {
                    $child = $Context.Clone()
                    $child['.'] = $item
                    $acc += Render-Mustache -Template $inner -Context $child
                }
            }
            $replacement = $acc
        } elseif (Test-Truthy $val) {
            $replacement = Render-Mustache -Template $inner -Context $Context
        }

        $output = $output.Substring(0, $m.Index) + $replacement + $output.Substring($m.Index + $m.Length)
        $m = $Script:RxSection.Match($output)
    }

    # Expand inverted sections
    $m = $Script:RxInvertedSection.Match($output)
    while ($m.Success) {
        $key = $m.Groups[1].Value
        $inner = $m.Groups[2].Value
        $val = Get-ContextValue -Context $Context -KeyPath $key
        $replacement = ''
        if (-not (Test-Truthy $val)) {
            $replacement = Render-Mustache -Template $inner -Context $Context
        }
        $output = $output.Substring(0, $m.Index) + $replacement + $output.Substring($m.Index + $m.Length)
        $m = $Script:RxInvertedSection.Match($output)
    }

    # Variables
    $output = $Script:RxTriple.Replace($output, {
            param($match)
            $val = Get-ContextValue -Context $Context -KeyPath $match.Groups[1].Value
            if ($null -eq $val) { '' } else { [string]$val }
        })

    $output = $Script:RxAmp.Replace($output, {
            param($match)
            $val = Get-ContextValue -Context $Context -KeyPath $match.Groups[1].Value
            if ($null -eq $val) { '' } else { [string]$val }
        })

    $output = $Script:RxSimple.Replace($output, {
            param($match)
            $val = Get-ContextValue -Context $Context -KeyPath $match.Groups[1].Value
            if ($null -eq $val) { '' } else { ConvertTo-HtmlEscapedText ([string]$val) }
        })

    return $output
}

# Compile a mustache template into an AST of nodes. The result is cached by
# template text. Each node is a 2- or 3-element array:
#   ('L', literal_text)         -- literal text
#   ('E', key)                  -- escaped variable        {{ key }}
#   ('U', key)                  -- unescaped variable      {{{ key }}} or {{& key }}
#   ('S', key, body_nodes)      -- section                 {{# key }}...{{/ key }}
#   ('I', key, body_nodes)      -- inverted section        {{^ key }}...{{/ key }}
function Compile-MustacheTemplate {
    param([string]$Template)
    if ([string]::IsNullOrEmpty($Template)) {
        return New-Object System.Collections.Generic.List[object]
    }
    $cached = $Script:CompiledTemplateCache[$Template]
    if ($null -ne $cached) { return $cached }

    # Tokenize
    $tokenMatches = $Script:RxMustacheTokens.Matches($Template)
    $tokens = New-Object System.Collections.Generic.List[object]
    $pos = 0
    foreach ($m in $tokenMatches) {
        if ($m.Index -gt $pos) {
            $tokens.Add(@('L', $Template.Substring($pos, $m.Index - $pos)))
        }
        $v = $m.Value
        if ($v[2] -eq '{') {
            # {{{ key }}}
            $key = $v.Substring(3, $v.Length - 6).Trim()
            $tokens.Add(@('U', $key))
        } else {
            $sigil = $v[2]
            switch ($sigil) {
                '#' { $tokens.Add(@('SO', $v.Substring(3, $v.Length - 5).Trim())) }
                '^' { $tokens.Add(@('IO', $v.Substring(3, $v.Length - 5).Trim())) }
                '/' { $tokens.Add(@('SC', $v.Substring(3, $v.Length - 5).Trim())) }
                '&' { $tokens.Add(@('U', $v.Substring(3, $v.Length - 5).Trim())) }
                default {
                    $tokens.Add(@('E', $v.Substring(2, $v.Length - 4).Trim()))
                }
            }
        }
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $Template.Length) {
        $tokens.Add(@('L', $Template.Substring($pos)))
    }

    # Build AST using a stack
    $root = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[object]
    $current = $root
    foreach ($t in $tokens) {
        $type = $t[0]
        if ($type -eq 'SO') {
            $body = New-Object System.Collections.Generic.List[object]
            $node = @('S', $t[1], $body)
            $current.Add($node)
            $stack.Push($current)
            $current = $body
        } elseif ($type -eq 'IO') {
            $body = New-Object System.Collections.Generic.List[object]
            $node = @('I', $t[1], $body)
            $current.Add($node)
            $stack.Push($current)
            $current = $body
        } elseif ($type -eq 'SC') {
            if ($stack.Count -gt 0) { $current = $stack.Pop() }
        } else {
            $current.Add($t)
        }
    }

    $Script:CompiledTemplateCache[$Template] = $root
    return $root
}

# Render a pre-compiled mustache AST into a StringBuilder.
# Behaves like Render-Mustache (lazy section semantics, dotted keys, etc.)
# but without re-parsing the template on every call.
function Render-CompiledMustache {
    param(
        [Parameter(Mandatory)][object]$Nodes,
        [Parameter(Mandatory)][hashtable]$Context,
        [System.Text.StringBuilder]$Sb = $null
    )
    $top = $false
    if ($null -eq $Sb) {
        $Sb = [System.Text.StringBuilder]::new()
        $top = $true
    }

    foreach ($n in $Nodes) {
        $type = $n[0]
        if ($type -eq 'L') {
            [void]$Sb.Append([string]$n[1])
        } elseif ($type -eq 'E') {
            $key = $n[1]
            if ($key.IndexOf('.') -lt 0) {
                $val = $Context[$key]
            } else {
                $val = Get-ContextValue -Context $Context -KeyPath $key
            }
            if ($null -ne $val) {
                [void]$Sb.Append((ConvertTo-HtmlEscapedText ([string]$val)))
            }
        } elseif ($type -eq 'U') {
            $key = $n[1]
            if ($key.IndexOf('.') -lt 0) {
                $val = $Context[$key]
            } else {
                $val = Get-ContextValue -Context $Context -KeyPath $key
            }
            if ($null -ne $val) { [void]$Sb.Append([string]$val) }
        } elseif ($type -eq 'S') {
            $key = $n[1]
            $body = $n[2]
            if ($key.IndexOf('.') -lt 0) {
                $val = $Context[$key]
            } else {
                $val = Get-ContextValue -Context $Context -KeyPath $key
            }
            if ($null -eq $val) { continue }
            if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                foreach ($item in $val) {
                    if ($item -is [System.Collections.IDictionary]) {
                        $child = $Context.Clone()
                        foreach ($k in $item.Keys) { $child[$k] = $item[$k] }
                        Render-CompiledMustache -Nodes $body -Context $child -Sb $Sb
                    } else {
                        $child = $Context.Clone()
                        $child['.'] = $item
                        Render-CompiledMustache -Nodes $body -Context $child -Sb $Sb
                    }
                }
            } elseif (Test-Truthy $val) {
                Render-CompiledMustache -Nodes $body -Context $Context -Sb $Sb
            }
        } elseif ($type -eq 'I') {
            $key = $n[1]
            if ($key.IndexOf('.') -lt 0) {
                $val = $Context[$key]
            } else {
                $val = Get-ContextValue -Context $Context -KeyPath $key
            }
            if (-not (Test-Truthy $val)) {
                Render-CompiledMustache -Nodes $n[2] -Context $Context -Sb $Sb
            }
        }
    }

    if ($top) { return $Sb.ToString() }
}

# Replace {{#first}}...{{/first}} with sentinel tags that we evaluate after render
function Replace-TemplateLambdas {
    param([string]$TemplateText)
    if ($null -eq $TemplateText) { return $TemplateText }
    return $Script:RxFirst.Replace($TemplateText, { param($m) "FIRSTSTART$($m.Groups[1].Value)FIRSTEND" })
}

# Evaluate FIRSTSTART...FIRSTEND by selecting the first non-empty chunk (split on newline)
function Evaluate-TemplateLambdas {
    param([string]$Rendered)
    if ($null -eq $Rendered) { return '' }
    $m = $Script:RxFirstSentinel.Match($Rendered)
    while ($m.Success) {
        $replacement = Select-First -Text $m.Groups[1].Value
        $Rendered = $Rendered.Substring(0, $m.Index) + $replacement + $Rendered.Substring($m.Index + $m.Length)
        $m = $Script:RxFirstSentinel.Match($Rendered)
    }
    return $Script:RxNewline.Replace($Rendered, $Script:NewlineEnv)
}

function Select-First {
    param([string]$Text)
    if ($null -eq $Text) { return '' }

    # Split on '||' or newlines (with optional surrounding spaces)
    $parts = $Script:RxSelectFirst.Split($Text)

    foreach ($p in $parts) {
        $candidate = $p.Trim()
        if ($candidate.Length -gt 0 -and -not [string]::IsNullOrWhiteSpace($candidate)) {
            return $candidate
        }
    }
    return ''
}

# ------------------------------
# Configuration loader
# ------------------------------
function New-AddressFormatter {
    param(
        [Parameter(Mandatory)][string]$ConfPath,
        [switch]$NoWarnings,
        [switch]$OnlyAddress
    )

    $Script:AddrFmt.ConfPath = $ConfPath
    $Script:AddrFmt.ShowWarnings = -not [bool]$NoWarnings
    $Script:AddrFmt.OnlyAddress = [bool]$OnlyAddress
    $Script:AddrFmt.FinalComponents = $null
    $Script:AddrFmt.SetDistrictAlias = @{}

    # Reset maps
    $Script:AddrFmt.Templates = @{}
    $Script:AddrFmt.ComponentAliases = @{}
    $Script:AddrFmt.Component2Type = @{}
    $Script:AddrFmt.OrderedComponents = @()
    $Script:AddrFmt.HKnown = @{}
    $Script:AddrFmt.State_Codes = @{}
    $Script:AddrFmt.County_Codes = @{}
    $Script:AddrFmt.Country2Lang = @{}
    $Script:AddrFmt.Abbreviations = @{}

    # Reset derived caches (built by Build-CodeLookups / Compile-AllAbbreviations)
    $Script:AddrFmt.State_Codes_Reverse = @{}
    $Script:AddrFmt.State_Codes_Name = @{}
    $Script:AddrFmt.County_Codes_Reverse = @{}
    $Script:AddrFmt.County_Codes_Name = @{}
    $Script:AddrFmt.CompiledAbbreviations = @{}

    # components.yaml
    $componentsPath = Join-Path $ConfPath 'components.yaml'
    Throw-IfNullOrMissingPath $componentsPath 'components.yaml'
    $components = Import-Yaml -Path $componentsPath

    # worldwide.yaml (try countries/worldwide.yaml, then worldwide.yaml)
    $wwFile = Join-Path $ConfPath 'countries/worldwide.yaml'
    if (-not (Test-Path -LiteralPath $wwFile)) {
        $wwFile = Join-Path $ConfPath 'worldwide.yaml'
    }
    Throw-IfNullOrMissingPath $wwFile 'worldwide.yaml'
    $templatesMap = Import-Yaml -Path $wwFile

    # ConvertFrom-Yaml (without -Ordered) already returns plain hashtables for
    # mappings and List[object] for sequences, so the previous deep walk via
    # ConvertTo-Hashtable was just producing a structurally identical copy of
    # the entire ~10k-node tree. Skip it. Fall back to a defensive walk only
    # if some unexpected loader returned a PSCustomObject root.
    if (-not ($templatesMap -is [System.Collections.IDictionary])) {
        $templatesMap = ConvertTo-Hashtable -InputObject $templatesMap
    }

    # Some forks wrap under 'countries:' or similar. Normalize to a flat map of countryCode -> config.
    if ($templatesMap.ContainsKey('countries') -and ($templatesMap['countries'] -is [hashtable])) {
        $templatesMap = $templatesMap['countries']
    } elseif ($templatesMap.ContainsKey('worldwide') -and ($templatesMap['worldwide'] -is [hashtable])) {
        # Just in case a wrapper key is 'worldwide' instead of being the file name.
        $templatesMap = $templatesMap['worldwide']
    }


    # Adopt the parsed map directly (avoids copying ~250 entries one by one).
    # We don't add new top-level keys to Templates anywhere else, so sharing
    # the underlying hashtable is safe.
    $Script:AddrFmt.Templates = $templatesMap

    # Build alias maps and ordered components
    foreach ($c in $components) {
        if ($c.name) {
            if ($c.aliases) {
                $Script:AddrFmt.ComponentAliases[$c.name] = @($c.aliases)
            } else {
                $Script:AddrFmt.ComponentAliases[$c.name] = @()
            }
        }
    }
    $orderedList = New-Object System.Collections.Generic.List[object]
    $hknown = @{}
    foreach ($c in $components) {
        $name = $c.name
        $orderedList.Add($name)
        $Script:AddrFmt.Component2Type[$name] = $name
        $hknown[$name] = 1
        if ($c.aliases) {
            foreach ($a in $c.aliases) {
                $orderedList.Add($a)
                $Script:AddrFmt.Component2Type[$a] = $name
                $hknown[$a] = 1
            }
        }
    }
    $Script:AddrFmt.OrderedComponents = $orderedList.ToArray()
    $Script:AddrFmt.HKnown = $hknown

    # Load conf files: state_codes.yaml, county_codes.yaml, country2lang.yaml
    foreach ($fileBase in @('state_codes', 'county_codes', 'country2lang')) {
        $yf = Join-Path $ConfPath "$fileBase.yaml"
        if (Test-Path -LiteralPath $yf) {
            try {
                $y = Import-Yaml -Path $yf
                switch ($fileBase) {
                    'state_codes' { $Script:AddrFmt.State_Codes = $y }
                    'county_codes' { $Script:AddrFmt.County_Codes = $y }
                    'country2lang' { $Script:AddrFmt.Country2Lang = $y }
                }
            } catch {
                Warn-If "Error parsing $fileBase configuration: $($_.Exception.Message)"
            }
        }
    }

    # Abbreviations directory
    $abbrDir = Join-Path $ConfPath 'abbreviations'
    if (Test-Path -LiteralPath $abbrDir -PathType Container) {
        $abbrFiles = [System.IO.Directory]::GetFiles($abbrDir, '*.yaml')
        foreach ($f in $abbrFiles) {
            $name = [System.IO.Path]::GetFileName($f)
            if ($name -match '^(\w\w)\.yaml$') {
                $lang = $Matches[1]
                try {
                    $Script:AddrFmt.Abbreviations[$lang] = Import-Yaml -Path $f
                } catch {
                    Warn-If "Error parsing abbreviations in '$f': $($_.Exception.Message)"
                }
            }
        }
    }

    # Reset compiled-template cache (per-instance, in case ConfPath changed)
    $Script:CompiledTemplateCache = @{}

    # One-time pre-compilation, mirroring Geo::Address::Formatter::_read_configuration:
    #  - compile every replace / postformat_replace regex once per country
    #  - build O(1) reverse lookups for state_codes / county_codes
    #  - compile every abbreviation pattern once per (lang, component)
    Compile-AllReplacements
    Build-CodeLookups
    Compile-AllAbbreviations
}

# Pre-compile replace and postformat_replace regex patterns for each country
# template. Stores results back on the template hashtable as
# `_compiled_replace` and `_compiled_postformat`.
function Compile-AllReplacements {
    $rxIcase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor $Script:RxOpts
    $rxPlain = $Script:RxOpts

    # Pattern-string -> compiled regex caches. Many country templates inherit
    # identical replace/postformat rules via YAML merge keys, so the same
    # pattern string is encountered many times. Caching avoids re-compiling.
    $rxCacheIcase = @{}
    $rxCachePlain = @{}

    foreach ($cc in @($Script:AddrFmt.Templates.Keys)) {
        $tpl = $Script:AddrFmt.Templates[$cc]
        if (-not ($tpl -is [hashtable])) { continue }

        # replace rules
        if ($tpl.ContainsKey('replace') -and $tpl['replace']) {
            $compiled = New-Object System.Collections.Generic.List[object]
            foreach ($ra in $tpl['replace']) {
                if ($null -eq $ra -or $ra.Count -lt 2) { continue }
                $pattern = [string]$ra[0]
                $replacement = [string]$ra[1]

                $compName = $null
                $exactRest = $null
                $reKey = $pattern

                if ($pattern -match '^(\w+)=(.+)$') {
                    $compName = $Matches[1]
                    $exactRest = $Matches[2]
                    $reKey = $exactRest
                }

                $re = $rxCacheIcase[$reKey]
                if ($null -eq $re) {
                    try {
                        $re = [regex]::new($reKey, $rxIcase)
                        $rxCacheIcase[$reKey] = $re
                    } catch {
                        Warn-If "invalid replacement regex '$reKey' for $cc, skipping"
                        continue
                    }
                }

                $compiled.Add(@{
                        Component   = $compName
                        ExactMatch  = $exactRest
                        Re          = $re
                        Replacement = $replacement
                    })
            }
            $tpl['_compiled_replace'] = $compiled
        }

        # postformat_replace rules
        if ($tpl.ContainsKey('postformat_replace') -and $tpl['postformat_replace']) {
            $compiled = New-Object System.Collections.Generic.List[object]
            foreach ($ra in $tpl['postformat_replace']) {
                if ($null -eq $ra -or $ra.Count -lt 2) { continue }
                $pattern = [string]$ra[0]
                $replacement = [string]$ra[1]

                $re = $rxCachePlain[$pattern]
                if ($null -eq $re) {
                    try {
                        $re = [regex]::new($pattern, $rxPlain)
                        $rxCachePlain[$pattern] = $re
                    } catch {
                        Warn-If "invalid postformat regex '$pattern' for $cc, skipping"
                        continue
                    }
                }
                $compiled.Add(@{ Re = $re; Replacement = $replacement })
            }
            $tpl['_compiled_postformat'] = $compiled
        }
    }
}

# Build reverse lookup tables for state_codes and county_codes so that
# Add-Code is O(1) per call instead of iterating all entries.
#   $Script:AddrFmt.State_Codes_Reverse[$cc][UPPERCASE_NAME]  = code
#   $Script:AddrFmt.State_Codes_Name[$cc][code]               = default name
function Build-CodeLookups {
    foreach ($type in @('State_Codes', 'County_Codes')) {
        $reverseKey = "${type}_Reverse"
        $nameKey = "${type}_Name"
        $Script:AddrFmt[$reverseKey] = @{}
        $Script:AddrFmt[$nameKey] = @{}

        $data = $Script:AddrFmt[$type]
        if ($null -eq $data) { continue }

        foreach ($cc in $data.Keys) {
            $mapping = $data[$cc]
            if ($null -eq $mapping) { continue }
            $rev = @{}
            $name = @{}

            # mapping might be a hashtable or a PSObject (depending on YAML loader)
            $mappingIsDict = $mapping -is [System.Collections.IDictionary]
            $codes = $null
            if ($mappingIsDict) {
                $codes = $mapping.Keys
            } elseif ($mapping.PSObject -and $mapping.PSObject.Properties) {
                $codes = $mapping.PSObject.Properties.Name
            } else {
                continue
            }

            foreach ($code in $codes) {
                $val = if ($mappingIsDict) { $mapping[$code] } else { $mapping.$code }
                if ($val -is [System.Collections.IDictionary]) {
                    if ($val.Contains('default')) {
                        $name[$code] = [string]$val['default']
                    }
                    foreach ($v in $val.Values) {
                        if ($null -ne $v) {
                            $rev[([string]$v).ToUpperInvariant()] = $code
                        }
                    }
                } elseif ($val.PSObject -and $val.PSObject.Properties -and -not ($val -is [string])) {
                    if ($val.PSObject.Properties.Name -contains 'default') {
                        $name[$code] = [string]$val.default
                    }
                    foreach ($p in $val.PSObject.Properties) {
                        if ($null -ne $p.Value) {
                            $rev[([string]$p.Value).ToUpperInvariant()] = $code
                        }
                    }
                } else {
                    if ($null -ne $val) {
                        $name[$code] = [string]$val
                        $rev[([string]$val).ToUpperInvariant()] = $code
                    }
                }
                # code-to-code identity lookup (Perl: $rev->{$code} = $code)
                $rev[([string]$code).ToUpperInvariant()] = $code
            }
            $Script:AddrFmt[$reverseKey][$cc] = $rev
            $Script:AddrFmt[$nameKey][$cc] = $name
        }
    }
}

# Pre-compile abbreviation regexes per (lang, component). Stored as
# $Script:AddrFmt.CompiledAbbreviations[$lang][$compName] = list of
#   { Re = <regex>; Short = <short string> }
function Compile-AllAbbreviations {
    $Script:AddrFmt.CompiledAbbreviations = @{}
    if ($null -eq $Script:AddrFmt.Abbreviations) { return }

    # Cache compiled regex by pattern string (same long-form key in multiple
    # components/languages is common).
    $rxCache = @{}
    $rxOpts = $Script:RxOpts

    foreach ($lang in $Script:AddrFmt.Abbreviations.Keys) {
        $abbr = $Script:AddrFmt.Abbreviations[$lang]
        if ($null -eq $abbr) { continue }
        $perLang = @{}
        foreach ($compName in $abbr.Keys) {
            $rh_pairs = $abbr[$compName]
            if ($null -eq $rh_pairs) { continue }
            $list = New-Object System.Collections.Generic.List[object]
            $isDict = $rh_pairs -is [System.Collections.IDictionary]
            # rh_pairs may be a hashtable or PSObject
            $longs = $null
            if ($isDict) {
                $longs = $rh_pairs.Keys
            } elseif ($rh_pairs.PSObject -and $rh_pairs.PSObject.Properties) {
                $longs = $rh_pairs.PSObject.Properties.Name
            } else {
                continue
            }
            foreach ($long in $longs) {
                $short = if ($isDict) { $rh_pairs[$long] } else { $rh_pairs.$long }
                if ($null -eq $short) { continue }
                $patt = '(^|\s)' + [regex]::Escape([string]$long) + '\b'
                $re = $rxCache[$patt]
                if ($null -eq $re) {
                    try {
                        $re = [regex]::new($patt, $rxOpts)  # Perl version is case-sensitive here
                        $rxCache[$patt] = $re
                    } catch {
                        Warn-If "invalid abbreviation pattern: $patt"
                        continue
                    }
                }
                $list.Add(@{ Re = $re; Short = [string]$short })
            }
            $perLang[$compName] = $list
        }
        $Script:AddrFmt.CompiledAbbreviations[$lang] = $perLang
    }
}

# ------------------------------
# Public: return final components after last Format-PostalAddress
# ------------------------------
function Get-FinalComponents {
    if ($null -ne $Script:AddrFmt.FinalComponents) { return $Script:AddrFmt.FinalComponents }
    Warn-If 'final_components not yet set'
    return $null
}

# ------------------------------
# Core formatting
# ------------------------------
# NOTE: The public function is Format-PostalAddress, which calls Format-PostalAddressInternal


# ------------------------------
# Internal helpers (behavioral parity)
# ------------------------------

function Invoke-Postformat {
    param([string]$Text, [object]$Rules)

    # Split, filter empty lines, normalize each, and keep non-empty lines
    $rawLines = $Script:RxNewline.Split($Text)
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($line in $rawLines) {
        if (-not $line) { continue }
        $l = $line.Trim()
        if ($l.Length -eq 0) { continue }
        $l = $Script:RxPF_Dash.Replace($l, '')
        $l = $Script:RxPF_CommaCom.Replace($l, ', ')
        $l = $Script:RxPF_SpcCommaSpc.Replace($l, ', ')
        $l = $Script:RxPF_MultiSpace.Replace($l, ' ')
        $l = $Script:RxPF_LeadComma.Replace($l, '')
        $l = $Script:RxPF_MultiComma.Replace($l, ',')
        $l = $Script:RxPF_TrailComma.Replace($l, '')
        $l = $l.Trim()
        if ($l.Length -gt 0) { $kept.Add($l) }
    }
    $Text = [string]::Join("`n", $kept)

    # Remove duplicates across comma-separated pieces (keep first; except "new york")
    $before = $Text.Split(@(', '), [System.StringSplitOptions]::None)
    $seen = @{}
    $after = New-Object System.Collections.Generic.List[string]
    foreach ($p in $before) {
        $piece = $Script:RxLeadingWS.Replace($p, '')
        if ($piece -ine 'new york') {
            if ($seen.ContainsKey($piece)) { continue }
            $seen[$piece] = 1
        }
        $after.Add($piece)
    }
    $Text = [string]::Join(', ', $after)


    # Country-specific regex replacements with $1/$2/$3 backrefs.
    # Accepts either pre-compiled rules (list of @{Re=..; Replacement=..}) or
    # raw array-of-arrays. Pre-compiled is the fast path used at runtime.
    if ($Rules) {
        $isCompiled = $false
        if ($Rules -is [System.Collections.IList] -and $Rules.Count -gt 0) {
            $first = $Rules[0]
            if ($first -is [System.Collections.IDictionary] -and $first.Contains('Re')) {
                $isCompiled = $true
            }
        }
        if ($isCompiled) {
            foreach ($rule in $Rules) {
                $Text = $rule['Re'].Replace($Text, [string]$rule['Replacement'])
            }
        } else {
            foreach ($rule in $Rules) {
                try {
                    $from = [string]$rule[0]
                    $to = [string]$rule[1]
                    $rx = $Script:ApplyRxCache[$from]
                    if ($null -eq $rx) {
                        $rx = [regex]::new($from, $Script:RxOpts)
                        $Script:ApplyRxCache[$from] = $rx
                    }
                    $Text = $rx.Replace($Text, $to)
                } catch {
                    Warn-If ('invalid replacement: ' + ($rule -join ', '))
                }
            }
        }
    }

    return $Text
}

function Invoke-SanityCleaning {
    param([hashtable]$Components)

    # Postcode sanity
    if ($Components.ContainsKey('postcode')) {
        $pc = [string]$Components['postcode']
        if ($pc.Length -gt 20) {
            $Components.Remove('postcode') | Out-Null
        } elseif ($Script:RxSC_PostcodeSemi.IsMatch($pc)) {
            $Components.Remove('postcode') | Out-Null
        } else {
            $mm = $Script:RxSC_PostcodeFive.Match($pc)
            if ($mm.Success) {
                $Components['postcode'] = $mm.Groups[1].Value
            }
        }
    }

    # Remove null/empty/no-word/URL values
    $keys = @($Components.Keys)
    foreach ($c in $keys) {
        $v = $Components[$c]
        if ($null -eq $v) { $Components.Remove($c) | Out-Null; continue }
        $sv = [string]$v
        if (-not $Script:RxSC_HasWord.IsMatch($sv)) { $Components.Remove($c) | Out-Null; continue }
        if ($Script:RxSC_Url.IsMatch($sv)) { $Components.Remove($c) | Out-Null; continue }
    }
}

function Test-MinimalComponents {
    param([hashtable]$Components)
    # Perl: required (road, postcode), threshold=2 => if both missing -> false
    $missing = 0
    foreach ($c in @('road', 'postcode')) {
        if (-not ($Components.ContainsKey($c))) { $missing++ }
        if ($missing -eq 2) { return $false }
    }
    return $true
}

# Build primary types from aliases according to ordering
function Resolve-PrimaryAliases {
    param([hashtable]$Components)

    # Collect primary types whose alias(es) exist but primary not set
    $p2aliases = @{}  # primary -> set of alias keys present in Components
    foreach ($c in @($Components.Keys)) {
        if ($Script:AddrFmt.ComponentAliases.ContainsKey($c)) { continue } # it's a primary type
        if ($Script:AddrFmt.Component2Type.ContainsKey($c)) {
            $ptype = $Script:AddrFmt.Component2Type[$c]
            if (-not $Components.ContainsKey($ptype)) {
                if (-not $p2aliases.ContainsKey($ptype)) { $p2aliases[$ptype] = @{} }
                $p2aliases[$ptype][$c] = 1
            }
        }
    }

    foreach ($ptype in $p2aliases.Keys) {
        $aliases = @($p2aliases[$ptype].Keys)
        if ($aliases.Count -eq 1) {
            $Components[$ptype] = $Components[$aliases[0]]
            continue
        }
        # multiple aliases => follow configured alias order for the primary
        foreach ($a in $Script:AddrFmt.ComponentAliases[$ptype]) {
            if ($Components.ContainsKey($a)) {
                $Components[$ptype] = $Components[$a]
                break
            }
        }
    }
}

# Country code determination + dependent-territory adjustments
function Determine-CountryCode {
    param([hashtable]$Components)

    if (-not $Components.ContainsKey('country_code')) { return $null }
    $cc = [string]$Components['country_code']
    if ([string]::IsNullOrWhiteSpace($cc)) { return $null }
    if ($cc.Length -ne 2) { return $null }
    if ($cc -ieq 'uk') { return 'GB' }

    # Dependent territory: use another country's configuration
    if ($Script:AddrFmt.Templates.ContainsKey($cc) -and
        $Script:AddrFmt.Templates[$cc].use_country) {

        $old_cc = $cc
        $cc = [string]$Script:AddrFmt.Templates[$old_cc].use_country

        # change_country string with $component substitution
        if ($Script:AddrFmt.Templates[$old_cc].change_country) {
            $newCountry = [string]$Script:AddrFmt.Templates[$old_cc].change_country
            $m = [regex]::Match($newCountry, '\$(\w*)')
            if ($m.Success) {
                $component = $m.Groups[1].Value
                if ($Components.ContainsKey($component)) {
                    $newCountry = $newCountry -replace "\`$$([regex]::escape($component))", [string]$Components[$component]
                } else {
                    $newCountry = $newCountry -replace "\`$$([regex]::escape($component))", ''
                }
            }
            $Components['country'] = $newCountry
        }
        if ($Script:AddrFmt.Templates[$old_cc].add_component) {
            $tmp = [string]$Script:AddrFmt.Templates[$old_cc].add_component
            $kv = $tmp.Split('=', 2)
            if ($kv.Count -eq 2) {
                $k, $v = $kv[0], $kv[1]
                if ($k -ieq 'state') { $Components['state'] = $v }
            }
        }
    }

    # NL special handling -> CW/SX/AW
    if ($cc -ieq 'NL') {
        if ($Components.ContainsKey('state')) {
            switch -regex ($Components['state']) {
                '^Cura[cç]ao' { $cc = 'CW'; $Components['country'] = 'Curaçao' }
                '^sint maarten' { $cc = 'SX'; $Components['country'] = 'Sint Maarten' }
                '^Aruba' { $cc = 'AW'; $Components['country'] = 'Aruba' }
            }
        }
    }

    return $cc
}

function Fix-Country {
    param([hashtable]$Components)
    if ($Components.ContainsKey('country') -and $Components.ContainsKey('state')) {
        $country = $Components['country']
        $state = $Components['state']
        $isNumber = $false
        try { [void][double]::Parse([string]$country); $isNumber = $true } catch { $isNumber = $false }
        if ($isNumber) {
            $Components['country'] = $state
            $Components.Remove('state') | Out-Null
        }
    }
}

function Add-StateCode {
    param([hashtable]$Components)
    if ($Components.ContainsKey('state')) { return Add-Code -KeyName 'state' -Components $Components }
    return $null
}

function Add-CountyCode {
    param([hashtable]$Components)
    if ($Components.ContainsKey('county')) { return Add-Code -KeyName 'county' -Components $Components }
    return $null
}

function Add-Code {
    param([string]$KeyName, [hashtable]$Components)

    if (-not $Components.ContainsKey('country_code')) { return $null }
    if (-not $Components.ContainsKey($KeyName)) { return $null }

    $codeKey = "${KeyName}_code"
    if ($Components.ContainsKey($codeKey)) {
        if ($Components[$codeKey] -ine $Components[$KeyName]) { return $Components[$codeKey] }
    }

    $cc = $Components['country_code'].ToString().ToUpperInvariant()

    # Pick the right reverse-lookup table (built once at init)
    $reverseAll = if ($KeyName -ieq 'state') { $Script:AddrFmt.State_Codes_Reverse } else { $Script:AddrFmt.County_Codes_Reverse }
    $namesAll = if ($KeyName -ieq 'state') { $Script:AddrFmt.State_Codes_Name } else { $Script:AddrFmt.County_Codes_Name }

    if ($null -ne $reverseAll -and $reverseAll.ContainsKey($cc)) {
        $rev = $reverseAll[$cc]
        $name = [string]$Components[$KeyName]
        $uc_name = $name.ToUpperInvariant()

        $foundCode = $rev[$uc_name]
        if ($null -ne $foundCode) {
            $Components[$codeKey] = $foundCode
            # If the input was actually the code (e.g. state => 'NC'), set keyname to full name
            if ($uc_name -eq ([string]$foundCode).ToUpperInvariant()) {
                if ($null -ne $namesAll -and $namesAll.ContainsKey($cc)) {
                    $fullName = $namesAll[$cc][$foundCode]
                    if ($null -ne $fullName) { $Components[$KeyName] = $fullName }
                }
            }
        }

        # US odd variants
        if ($cc -eq 'US' -and $KeyName -ieq 'state' -and -not $Components.ContainsKey('state_code')) {
            $state = [string]$Components['state']
            if ($state -match '^united states') {
                $state2 = ($state -replace '^United States', 'US').ToUpperInvariant()
                $fc = $rev[$state2]
                if ($null -ne $fc) { $Components['state_code'] = $fc }
            }
            if (-not $Components.ContainsKey('state_code') -and $state -match '^washington,?\s*d\.?c\.?') {
                $Components['state_code'] = 'DC'
                $Components['state'] = 'District of Columbia'
                $Components['city'] = 'Washington'
            }
        }
    }

    if ($Components.ContainsKey($codeKey)) { return $Components[$codeKey] }
    return $null
}

function Apply-Replacements {
    param([hashtable]$Components, [object]$Rules)
    if ($null -eq $Rules) { return }

    # Detect "pre-compiled" rule format: a list whose first element is a hashtable
    # that has a 'Re' property (set by Compile-AllReplacements).
    $isCompiled = $false
    if ($Rules -is [System.Collections.IList] -and $Rules.Count -gt 0) {
        $first = $Rules[0]
        if ($first -is [System.Collections.IDictionary] -and $first.Contains('Re')) {
            $isCompiled = $true
        }
    }

    if ($isCompiled) {
        # Fast path: pre-compiled rules. Mirrors Geo::Address::Formatter::_apply_replacements.
        foreach ($component in @($Components.Keys)) {
            if ($component -eq 'country_code' -or $component -eq 'house_number') { continue }
            $cur = [string]$Components[$component]
            foreach ($rule in $Rules) {
                $compName = $rule['Component']
                if ($null -ne $compName) {
                    if ($compName -ne $component) { continue }
                    if ($cur -ieq [string]$rule['ExactMatch']) {
                        $cur = [string]$rule['Replacement']
                    } else {
                        $cur = $rule['Re'].Replace($cur, [string]$rule['Replacement'])
                    }
                } else {
                    $cur = $rule['Re'].Replace($cur, [string]$rule['Replacement'])
                }
            }
            $Components[$component] = $cur
        }
        return
    }

    # Legacy path: raw array-of-arrays (used by tests / unconverted templates)
    foreach ($component in @($Components.Keys)) {
        if ($component -eq 'country_code' -or $component -eq 'house_number') { continue }
        $compEqPrefix = $component + '='
        $compEqLen = $compEqPrefix.Length
        foreach ($ra in $Rules) {
            $regexp = $null
            $from = [string]$ra[0]
            $to = [string]$ra[1]

            if ($from.Length -ge $compEqLen -and $from.StartsWith($compEqPrefix)) {
                $keyFrom = $from.Substring($compEqLen)
                if ([string]$Components[$component] -ieq $keyFrom) {
                    $Components[$component] = $to
                } else {
                    $regexp = $keyFrom
                }
            } else {
                $regexp = $from
            }

            if ($regexp) {
                try {
                    $re = $Script:ApplyRxCache[$regexp]
                    if ($null -eq $re) {
                        $re = [regex]::new($regexp, ([System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor $Script:RxOpts))
                        $Script:ApplyRxCache[$regexp] = $re
                    }
                    $Components[$component] = $re.Replace([string]$Components[$component], $to)
                } catch {
                    Warn-If ('invalid replacement: ' + ($ra -join ', '))
                }
            }
        }
    }
}

function Invoke-Abbreviate {
    param([hashtable]$Components)
    if (-not $Components.ContainsKey('country_code')) {
        $msg = 'no country_code, unable to abbreviate'
        if ($Components.ContainsKey('country')) { $msg += " - country: $($Components['country'])" }
        Warn-If $msg
        return $null
    }

    $cc = $Components['country_code'].ToString()
    if (-not $Script:AddrFmt.Country2Lang.ContainsKey($cc)) { return $Components }
    $langs = [string]$Script:AddrFmt.Country2Lang[$cc]

    # Prefer pre-compiled abbreviations (built once at init by Compile-AllAbbreviations)
    $compiledAll = $Script:AddrFmt.CompiledAbbreviations

    foreach ($lang in $langs.Split(',')) {
        if ($null -ne $compiledAll -and $compiledAll.ContainsKey($lang)) {
            $perLang = $compiledAll[$lang]
            foreach ($compName in $perLang.Keys) {
                if (-not $Components.ContainsKey($compName)) { continue }
                $cur = [string]$Components[$compName]
                foreach ($rule in $perLang[$compName]) {
                    $cur = $rule['Re'].Replace($cur, ('$1' + $rule['Short']))
                }
                $Components[$compName] = $cur
            }
        }
    }
    return $Components
}

# Line-by-line normalization with \s (no multiline \h usage)
function Invoke-Clean {
    param([string]$Text)
    if ($null -eq $Text) { return '' }

    # Convert HTML apostrophe back
    $Text = $Script:RxCl_HtmlApos.Replace($Text, "'")

    # Split into lines (preserve logical lines)
    $rawLines = $Script:RxNewline.Split($Text)

    $normalizedLines = New-Object System.Collections.Generic.List[string]
    foreach ($line in $rawLines) {
        $l = $line

        $l = $Script:RxCl_LBracket.Replace($l, '')
        $l = $Script:RxCl_RBracket.Replace($l, '')
        $l = $Script:RxCl_LComma.Replace($l, '')
        $l = $Script:RxCl_RComma.Replace($l, '')
        $l = $Script:RxCl_CommaCommas.Replace($l, ',')
        $l = $Script:RxCl_CommaSpace.Replace($l, ', ')
        $l = $Script:RxCl_MultiSpace.Replace($l, ' ')
        $l = $l.Trim()

        $normalizedLines.Add($l)
    }

    # Final dedupe across and within lines
    $seenLines = @{}
    $afterLines = New-Object System.Collections.Generic.List[string]
    foreach ($line in $normalizedLines) {
        $l = $line
        if ([string]::IsNullOrWhiteSpace($l)) { continue }

        if ($seenLines.ContainsKey($l)) { continue }
        $seenLines[$l] = 1

        # Deduplicate comma-separated items in the line (except "New York")
        $words = $l.Split(',')
        $seenWords = @{}
        $afterWords = New-Object System.Collections.Generic.List[string]
        foreach ($w in $words) {
            $w2 = $w.Trim()
            if ($w2 -ine 'new york') {
                if ($seenWords.ContainsKey($w2)) { continue }
                $seenWords[$w2] = 1
            }
            $afterWords.Add($w2)
        }
        $afterLines.Add([string]::Join(', ', $afterWords))
    }

    $out = [string]::Join("`n", $afterLines)
    return $out.Trim()
}

function Set-DistrictAlias {
    param([Parameter(Mandatory)][string]$CountryCode)

    $ucc = $CountryCode
    if ($Script:AddrFmt.SetDistrictAlias.ContainsKey($ucc)) { return }

    $Script:AddrFmt.SetDistrictAlias[$ucc] = 1
    $oldalias = $null

    if ($Script:SmallDistrict.ContainsKey($ucc)) {
        $Script:AddrFmt.Component2Type['district'] = 'neighbourhood'
        $oldalias = 'state_district'
        if (-not $Script:AddrFmt.ComponentAliases.ContainsKey('neighbourhood')) {
            $Script:AddrFmt.ComponentAliases['neighbourhood'] = @()
        }
        $Script:AddrFmt.ComponentAliases['neighbourhood'] += 'district'
    } else {
        $Script:AddrFmt.Component2Type['district'] = 'state_district'
        $oldalias = 'neighbourhood'
        if (-not $Script:AddrFmt.ComponentAliases.ContainsKey('state_district')) {
            $Script:AddrFmt.ComponentAliases['state_district'] = @()
        }
        $Script:AddrFmt.ComponentAliases['state_district'] += 'district'
    }

    if ($oldalias -and $Script:AddrFmt.ComponentAliases.ContainsKey($oldalias)) {
        $Script:AddrFmt.ComponentAliases[$oldalias] = @(
            $Script:AddrFmt.ComponentAliases[$oldalias] | Where-Object { $_ -ine 'district' }
        )
    }
}

function Find-UnknownComponents {
    param([hashtable]$Components)
    $list = New-Object System.Collections.Generic.List[object]
    $known = $Script:AddrFmt.HKnown
    foreach ($k in $Components.Keys) {
        if (-not $known.ContainsKey($k)) {
            $list.Add($k)
        }
    }
    return , ($list.ToArray())
}

# ------------------------------
# Module Initialization - Runs automatically on Import-Module
# ------------------------------

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

# $PSScriptRoot is the path to the current module script (.psm1 file)
$confPath = Join-Path $PSScriptRoot 'subModules\OpenCageData\address-formatting\conf'

# The path to your configuration data must be relative to the module root.
# Assuming 'conf' folder is a peer to the .psm1 file.
New-AddressFormatter -ConfPath $confPath