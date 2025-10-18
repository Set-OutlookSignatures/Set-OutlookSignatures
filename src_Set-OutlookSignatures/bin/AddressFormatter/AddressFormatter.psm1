# This function is the new, exported entry point
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

    $rh_components = Clone-Deep $Components
    if (-not $rh_components) { return $null }

    # Reset final components at start
    $Script:AddrFmt.FinalComponents = $null

    # Merge options
    $opt = @{}
    if ($Options) { $opt = Clone-Deep $Options }
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

    # 5) Apply replacements (pre-render)
    if ($rh_config -and $rh_config.replace) {
        Apply-Replacements -Components $rh_components -Rules $rh_config.replace
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

    # 9) Prepare template (replace lambda)
    $templateText = Replace-TemplateLambdas -TemplateText $templateText
    # 10) Render
    $rendered = Render-Mustache -Template $templateText -Context $rh_components
    $rendered = Evaluate-TemplateLambdas -Rendered $rendered

    # 11) Postformat replacements and duplicate removal
    $rendered = Invoke-Postformat -Text $rendered -Rules $rh_config.postformat_replace

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

    return $rendered
}


$Script:AddrFmt = @{
    Templates         = @{}   # country templates + default
    ComponentAliases  = @{}   # primary -> [aliases]
    Component2Type    = @{}   # alias -> primary
    OrderedComponents = @()   # [ primary1, alias1a, alias1b, primary2, ... ]
    HKnown            = @{}   # set of known component keys
    State_Codes       = @{}   # country_code -> map
    County_Codes      = @{}   # country_code -> map
    Country2Lang      = @{}   # country_code -> "en,de"
    Abbreviations     = @{}   # lang -> component -> (long -> short)
    SetDistrictAlias  = @{}   # cache to avoid rework
    FinalComponents   = $null # set after successful Format-PostalAddress

    ShowWarnings      = $true
    OnlyAddress       = $false

    ConfPath          = $null
}

# Countries where plain "district" should be treated as "neighbourhood" (small district)
$Script:SmallDistrict = @{
    'BR' = 1; 'CR' = 1; 'ES' = 1; 'NI' = 1; 'PY' = 1; 'RO' = 1; 'TG' = 1; 'TM' = 1; 'XK' = 1
}

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
        if ($InputObject -is [hashtable]) {
            $clone = @{}
            foreach ($k in $InputObject.Keys) {
                $clone[$k] = Clone-Deep $InputObject[$k]
            }
            return $clone
        } elseif ($InputObject -is [System.Collections.IDictionary]) {
            $clone = @{}
            foreach ($k in $InputObject.Keys) {
                $clone[$k] = Clone-Deep $InputObject[$k]
            }
            return $clone
        } elseif ($InputObject -is [System.Collections.IEnumerable] -and
            $InputObject -isnot [string]) {
            $arr = @()
            foreach ($i in $InputObject) { $arr += , (Clone-Deep $i) }
            return , $arr
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

    if ($InputObject -is [hashtable]) {
        $h = @{}
        foreach ($k in $InputObject.Keys) {
            $h[$k] = ConvertTo-Hashtable -InputObject $InputObject[$k]
        }
        return $h
    } elseif ($InputObject -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $InputObject.Keys) {
            $h[$k] = ConvertTo-Hashtable -InputObject $InputObject[$k]
        }
        return $h
    } elseif ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $h[$p.Name] = ConvertTo-Hashtable -InputObject $p.Value
        }
        return $h
    } elseif ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        $arr = @()
        foreach ($i in $InputObject) { $arr += , (ConvertTo-Hashtable -InputObject $i) }
        return , $arr
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


# YAML loader
function Import-Yaml {
    param([Parameter(Mandatory)][string]$Path)

    Throw-IfNullOrMissingPath -Path $Path -What 'YAML file'
    # Use the nested powershell-yaml module's function
    return (ConvertFrom-Yaml -Yaml (Get-Content -LiteralPath $path -Raw -Encoding UTF8) -UseMergingParser -AllDocuments)
}

# Resolve "a.b.c" in nested hashtables
function Get-ContextValue {
    param([hashtable]$Context, [string]$KeyPath)
    if ($null -eq $Context) { return $null }
    if ([string]::IsNullOrWhiteSpace($KeyPath)) { return $null }
    $node = $Context
    foreach ($part in $KeyPath.Split('.')) {
        if ($node -is [hashtable] -or $node -is [System.Collections.IDictionary]) {
            if ($node.ContainsKey($part)) {
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

    # Expand sections (nested same-name not supported; nested different names ok)
    $sectionPattern = [regex]'(?s){{\#\s*([\w\.\-]+)\s*}}(.*?){{/\s*\1\s*}}'
    $invertedPattern = [regex]'(?s){{\^\s*([\w\.\-]+)\s*}}(.*?){{/\s*\1\s*}}'

    $output = $Template

    # Expand regular sections
    $m = $sectionPattern.Match($output)
    while ($m.Success) {
        $key = $m.Groups[1].Value
        $inner = $m.Groups[2].Value
        $val = Get-ContextValue -Context $Context -KeyPath $key

        $replacement = ''
        if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
            $acc = ''
            foreach ($item in $val) {
                if ($item -is [hashtable] -or $item -is [System.Collections.IDictionary]) {
                    $child = Clone-Deep $Context
                    foreach ($k in $item.Keys) { $child[$k] = $item[$k] }
                    $acc += Render-Mustache -Template $inner -Context $child
                } else {
                    $child = Clone-Deep $Context
                    $child['.'] = $item
                    $acc += Render-Mustache -Template $inner -Context $child
                }
            }
            $replacement = $acc
        } elseif (Test-Truthy $val) {
            $replacement = Render-Mustache -Template $inner -Context $Context
        }

        $output = $output.Substring(0, $m.Index) + $replacement + $output.Substring($m.Index + $m.Length)
        $m = $sectionPattern.Match($output)
    }

    # Expand inverted sections
    $m = $invertedPattern.Match($output)
    while ($m.Success) {
        $key = $m.Groups[1].Value
        $inner = $m.Groups[2].Value
        $val = Get-ContextValue -Context $Context -KeyPath $key
        $replacement = ''
        if (-not (Test-Truthy $val)) {
            $replacement = Render-Mustache -Template $inner -Context $Context
        }
        $output = $output.Substring(0, $m.Index) + $replacement + $output.Substring($m.Index + $m.Length)
        $m = $invertedPattern.Match($output)
    }

    # Variables
    $triple = [regex]'(?s){{{\s*([\w\.\-]+)\s*}}}'
    $amp = [regex]'(?s){{&\s*([\w\.\-]+)\s*}}'
    $simple = [regex]'(?s){{\s*([\w\.\-]+)\s*}}'

    $output = $triple.Replace($output, {
            param($match)
            $val = Get-ContextValue -Context $Context -KeyPath $match.Groups[1].Value
            if ($null -eq $val) { '' } else { [string]$val }
        })

    $output = $amp.Replace($output, {
            param($match)
            $val = Get-ContextValue -Context $Context -KeyPath $match.Groups[1].Value
            if ($null -eq $val) { '' } else { [string]$val }
        })

    $output = $simple.Replace($output, {
            param($match)
            $val = Get-ContextValue -Context $Context -KeyPath $match.Groups[1].Value
            if ($null -eq $val) { '' } else { ConvertTo-HtmlEscapedText ([string]$val) }
        })

    return $output
}

# Replace {{#first}}...{{/first}} with sentinel tags that we evaluate after render
function Replace-TemplateLambdas {
    param([string]$TemplateText)
    if ($null -eq $TemplateText) { return $TemplateText }
    $rx = [regex]'(?s){{\#first}}(.+?){{\/first}}'
    return $rx.Replace($TemplateText, { param($m) "FIRSTSTART$($m.Groups[1].Value)FIRSTEND" })
}

# Evaluate FIRSTSTART...FIRSTEND by selecting the first non-empty chunk (split on newline)
function Evaluate-TemplateLambdas {
    param([string]$Rendered)
    if ($null -eq $Rendered) { return '' }
    $rx = [regex]'(?s)FIRSTSTART\s*(.+?)\s*FIRSTEND'
    $m = $rx.Match($Rendered)
    while ($m.Success) {
        $replacement = Select-First -Text $m.Groups[1].Value
        $Rendered = $Rendered.Substring(0, $m.Index) + $replacement + $Rendered.Substring($m.Index + $m.Length)
        $m = $rx.Match($Rendered)
    }
    return $Rendered
}

function Select-First {
    param([string]$Text)
    if ($null -eq $Text) { return '' }

    # An '||' ODER an Zeilenumbrüchen trennen, je mit optionalen Spaces
    $parts = [regex]::Split($Text, '\s*(?:\|\||\r?\n)\s*')

    foreach ($p in $parts) {
        # Trimmen & leere Teile überspringen
        $candidate = $p -replace '^\s+', '' -replace '\s+$', ''
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
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
    $templatesRaw = Import-Yaml -Path $wwFile


    # Some forks wrap under 'countries:' or similar. Normalize to a flat map of countryCode -> config.
    $templatesMap = ConvertTo-Hashtable -InputObject $templatesRaw

    if ($templatesMap.ContainsKey('countries') -and ($templatesMap['countries'] -is [hashtable])) {
        $templatesMap = $templatesMap['countries']
    } elseif ($templatesMap.ContainsKey('worldwide') -and ($templatesMap['worldwide'] -is [hashtable])) {
        # Just in case a wrapper key is 'worldwide' instead of being the file name.
        $templatesMap = $templatesMap['worldwide']
    }


    # Fill Templates dictionary
    foreach ($k in $templatesMap.Keys) {
        $Script:AddrFmt.Templates[$k] = $templatesMap[$k]
    }

    # Build alias maps and ordered components
    foreach ($c in $components) {
        if ($c.name) {
            $Script:AddrFmt.ComponentAliases[$c.name] = @()
            if ($c.aliases) { $Script:AddrFmt.ComponentAliases[$c.name] = @($c.aliases) }
        }
    }
    foreach ($c in $components) {
        $name = $c.name
        $Script:AddrFmt.OrderedComponents += , $name
        $Script:AddrFmt.Component2Type[$name] = $name
        if ($c.aliases) {
            foreach ($a in $c.aliases) {
                $Script:AddrFmt.OrderedComponents += , $a
                $Script:AddrFmt.Component2Type[$a] = $name
            }
        }
    }
    $Script:AddrFmt.HKnown = @{}
    foreach ($k in $Script:AddrFmt.OrderedComponents) { $Script:AddrFmt.HKnown[$k] = 1 }

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
        Get-ChildItem -LiteralPath $abbrDir -File -Filter '*.yaml' | ForEach-Object {
            if ($_.Name -match '^(\w\w)\.yaml$') {
                $lang = $Matches[1]
                try {
                    $Script:AddrFmt.Abbreviations[$lang] = Import-Yaml -Path $_.FullName
                } catch {
                    Warn-If "Error parsing abbreviations in '$($_.FullName)': $($_.Exception.Message)"
                }
            }
        }
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

    $Text = @(
        $Text -split '\r?\n' | Where-Object { $_ }
    ) -join "`n"

    $Text = @(
        $Text -split '\r?\n' | Where-Object { $_ } | ForEach-Object {
            (
                $_.Trim() `
                    -replace '^- ', '' `
                    -replace ',\s*,', ', ' `
                    -replace '\s+,\s+', ', ' `
                    -replace '\s\s+', ' ' `
                    -replace '^,', '' `
                    -replace ',,+', ',' `
                    -replace ',$', ''
            ).Trim()
        }
    ) -join "`n"

    $Text = @(
        $Text -split '\r?\n' | Where-Object { $_ }
    ) -join "`n"

    # Remove duplicates across comma-separated pieces (keep first; except "new york")
    $before = $Text -split ', '
    $seen = @{}
    $after = New-Object System.Collections.Generic.List[string]
    foreach ($p in $before) {
        $piece = ($p -replace '^\s+', '')
        $key = $piece
        if ($piece -ine 'new york') {
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = 1
        }
        $after.Add($piece)
    }
    $Text = ($after -join ', ')


    # Country-specific regex replacements with $1/$2/$3 backrefs
    if ($Rules) {
        foreach ($rule in $Rules) {
            try {
                $from = [string]$rule[0]
                $to = [string]$rule[1]
                $rx = [regex]$from
                $Text = $rx.Replace($Text, $to)
            } catch {
                Warn-If ('invalid replacement: ' + ($rule -join ', '))
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
        } elseif ($pc -match '^\d+;\d+$') {
            $Components.Remove('postcode') | Out-Null
        } elseif ($pc -match '^(\d{5}),\d{5}') {
            $Components['postcode'] = $Matches[1]
        }
    }

    # Remove null/empty/no-word/URL values
    $keys = @($Components.Keys)
    foreach ($c in $keys) {
        $v = $Components[$c]
        if ($null -eq $v) { $Components.Remove($c) | Out-Null; continue }
        $sv = [string]$v
        if ($sv -notmatch '\w') { $Components.Remove($c) | Out-Null; continue }
        if ($sv -match '(?s)https?://') { $Components.Remove($c) | Out-Null; continue }
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

    $cc = $Components['country_code'].ToString()
    $maps = if ($KeyName -ieq 'state') { $Script:AddrFmt.State_Codes } else { $Script:AddrFmt.County_Codes }
    if (-not $maps.ContainsKey($cc)) { return $null }
    $mapping = $maps[$cc]
    $name = [string]$Components[$KeyName]
    $uc_name = $name

    foreach ($abbrv in $mapping.Keys) {
        $confval = $mapping.$abbrv
        $confNames = @()
        if ($confval -is [System.Collections.IDictionary] -or $confval -is [hashtable]) {
            $confNames += $confval.Values
        } else {
            $confNames += , $confval
        }

        foreach ($confname in $confNames) {
            if ($uc_name -ieq ([string]$confname)) {
                $Components[$codeKey] = $abbrv
                break
            }
            if ($uc_name -ieq $abbrv) {
                $Components[$KeyName] = [string]$confname
                $Components[$codeKey] = $abbrv
                break
            }
        }
        if ($Components.ContainsKey($codeKey)) { break }
    }

    # US odd variants
    if ($cc -ieq 'US' -and $KeyName -ieq 'state' -and -not $Components.ContainsKey('state_code')) {
        $state = [string]$Components['state']
        if ($state -match '^united states') {
            $state2 = $state -replace '^United States', 'US'
            foreach ($k in $mapping.PSObject.Properties.Name) {
                if ($state2 -ieq $k) {
                    $Components['state_code'] = [string]$mapping.$k
                    break
                }
            }
        }
        if (-not $Components.ContainsKey('state_code') -and $state -match '^washington,?\s*d\.?c\.?' ) {
            $Components['state_code'] = 'DC'
            $Components['state'] = 'District of Columbia'
            $Components['city'] = 'Washington'
        }
    }

    return $Components[$codeKey]
}

function Apply-Replacements {
    param([hashtable]$Components, [object]$Rules)

    foreach ($component in @($Components.Keys)) {
        if ($component -in @('country_code', 'house_number')) { continue }
        foreach ($ra in $Rules) {
            $regexp = $null
            $from = [string]$ra[0]
            $to = [string]$ra[1]

            if ($from -match "^$([regex]::escape($component))=") {
                $keyFrom = $from.Substring($component.Length + 1)
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
                    $re = [regex]::new($regexp, 'IgnoreCase')
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
    foreach ($lang in $langs.Split(',')) {
        if ($Script:AddrFmt.Abbreviations.ContainsKey($lang)) {
            $rh_abbr = $Script:AddrFmt.Abbreviations[$lang]
            foreach ($compName in $rh_abbr.Keys) {
                if (-not $Components.ContainsKey($compName)) { continue }
                $map = $rh_abbr.$compName
                foreach ($long in $map.Keys) {
                    $short = [string]$map.$long
                    [string]$Components[$compName] = [string]$Components[$compName] -ireplace "(^|\s)$([regex]::escape($long))\b", "`$1$($short)"
                }
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
    $Text = $Text -replace '&#39;', "'"

    # Split into lines (preserve logical lines)
    $rawLines = [regex]::Split($Text, '\r?\n')

    $normalizedLines = New-Object System.Collections.Generic.List[string]
    foreach ($line in $rawLines) {
        $l = $line

        # Remove stray leading/trailing bracket/comma blocks (best effort)
        $l = $l -replace '^[\[\{,\s]+', ''
        $l = $l -replace '[\}\],\s]+$', ''

        # Remove leading/trailing commas on the line
        $l = $l -replace '^\s*,+\s*', ''
        $l = $l -replace '\s*,+\s*$', ''

        # Reduce multiple commas to one, uniformly space around commas as ", "
        $l = $l -replace ',\s*,+', ','         # ",  ,," -> ","
        $l = $l -replace '\s*,\s*', ', '       # normalize spaces around commas

        # Collapse multiple spaces to one (within line)
        $l = $l -replace '\s{2,}', ' '

        # Trim leading/trailing whitespace
        $l = $l -replace '^\s+', ''
        $l = $l -replace '\s+$', ''

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
        $words = $l -split ','
        $seenWords = @{}
        $afterWords = New-Object System.Collections.Generic.List[string]
        foreach ($w in $words) {
            $w2 = $w -replace '^\s+', '' -replace '\s+$', ''
            if ($w2 -ine 'new york') {
                if ($seenWords.ContainsKey($w2)) { continue }
                $seenWords[$w2] = 1
            }
            $afterWords.Add($w2)
        }
        $l2 = ($afterWords -join ', ')
        $afterLines.Add($l2)
    }

    $out = ($afterLines -join "`n")
    $out = $out -replace '^\s+', ''
    $out = $out -replace '\s+$', ''

    return $out
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
    $unknown = @()
    foreach ($k in $Components.Keys) {
        if (-not $Script:AddrFmt.HKnown.ContainsKey($k)) {
            $unknown += , $k
        }
    }
    return , $unknown
}

# ------------------------------
# Module Initialization - Runs automatically on Import-Module
# ------------------------------

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

# $PSScriptRoot is the path to the current module script (.psm1 file)
$confPath = Join-Path $PSScriptRoot 'address-formatting\conf'

# The path to your configuration data must be relative to the module root.
# Assuming 'conf' folder is a peer to the .psm1 file.
New-AddressFormatter -ConfPath $confPath