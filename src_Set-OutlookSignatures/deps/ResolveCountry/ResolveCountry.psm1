# Requires -Version 5.1


# In-Memory Cache Variable Definition (Module Scope)
$script:CountryLookupCache = $null


# Compile C# Levenshtein distance function (Cross-platform and PS5.1 safe)
if (-not ('FuzzyMatcher' -as [type])) {
    Add-Type -TypeDefinition @'
        public class FuzzyMatcher {
            public static int Levenshtein(string s, string t) {
                if (string.IsNullOrEmpty(s)) return string.IsNullOrEmpty(t) ? 0 : t.Length;
                if (string.IsNullOrEmpty(t)) return s.Length;

                // Allocate only two 1D arrays instead of a full 2D matrix
                int[] v0 = new int[t.Length + 1];
                int[] v1 = new int[t.Length + 1];

                for (int i = 0; i <= t.Length; i++) v0[i] = i;

                for (int i = 0; i < s.Length; i++) {
                    v1[0] = i + 1;

                    for (int j = 0; j < t.Length; j++) {
                        int cost = (s[i] == t[j]) ? 0 : 1;
                        v1[j + 1] = System.Math.Min(
                            System.Math.Min(v1[j] + 1, v0[j + 1] + 1),
                            v0[j] + cost
                        );
                    }

                    // Copy current row to previous row
                    System.Array.Copy(v1, v0, v0.Length);
                }

                return v1[t.Length];
            }
        }
'@
}


# Text Normalization Helper (Internal)
function ConvertTo-NormalizedText ([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $decomposed = $text.Normalize([System.Text.NormalizationForm]::FormKD)
    $noDiacritics = $decomposed -replace '\p{M}', ''
    $noPunctuation = $noDiacritics -replace '[\p{P}\p{S}\p{C}]', ' '
    return ($noPunctuation -replace '\s+', ' ').Trim().ToLowerInvariant()
}


# Helper to read and decompress cache file into Hashtable (Internal)
function Import-CompressedCountryCache ([string]$Path) {
    $fileStream = [System.IO.File]::OpenRead($Path)
    $gzipStream = [System.IO.Compression.GZipStream]::new($fileStream, [System.IO.Compression.CompressionMode]::Decompress)
    $reader = [System.IO.StreamReader]::new($gzipStream, [System.Text.Encoding]::UTF8)

    try {
        $json = $reader.ReadToEnd()
        $rawObject = $json | ConvertFrom-Json

        # Convert PSCustomObject back to Hashtable for identical key lookup performance
        $hash = @{}
        foreach ($prop in $rawObject.psobject.properties) {
            $hash[$prop.Name] = $prop.Value
        }
        return $hash
    } finally {
        $reader.Dispose()
        $gzipStream.Dispose()
        $fileStream.Dispose()
    }
}


# Data Download and Compressed Cache Generation Function (Public)
function Update-ResolveCountryData {
    [CmdletBinding()]
    param (
        [string]$CacheFilePath = (Join-Path -Path $PSScriptRoot -ChildPath 'ResolveCountry.Data.json.gz')
    )

    $tempDir = New-Item -Path ([System.IO.Path]::GetTempPath()) -Name (New-Guid).Guid -ItemType Directory
    $originalLocation = Get-Location
    Set-Location -Path $tempDir.FullName

    try {
        Write-Verbose 'Downloading countries.json...'
        Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/mledoze/countries/refs/heads/master/dist/countries.json' -OutFile 'countries.json' -UseBasicParsing

        $ImportCountries = Get-Content 'countries.json' -Raw | ConvertFrom-Json
        $Result = @{}

        foreach ($ImportCountry in $ImportCountries) {
            if ([string]::IsNullOrWhiteSpace($ImportCountry.cca2)) { continue }

            # Extract native names
            $nativeValues = @()
            if ($null -ne $ImportCountry.name.native) {
                $nativeValues = foreach ($nat in $ImportCountry.name.native.psobject.properties) {
                    $nat.Value.official
                    $nat.Value.common
                }
            }

            # Extract built-in translations
            $translationValues = @()
            if ($null -ne $ImportCountry.translations) {
                $translationValues = foreach ($trans in $ImportCountry.translations.psobject.properties) {
                    $trans.Value.official
                    $trans.Value.common
                }
            }

            # Build unified lookup array
            $ImportCountryUniqueValues = @(
                $ImportCountry.name.official
                $ImportCountry.name.common
                $nativeValues
                $ImportCountry.cca2
                $ImportCountry.cca3
                $ImportCountry.ccn3
                $ImportCountry.altSpellings
                $translationValues
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique

            $Result["$($ImportCountry.cca2)"] = [pscustomobject]@{
                UniqueValues  = $ImportCountryUniqueValues
                CountryObject = $ImportCountry
            }
        }

        # Save minified, Gzip-compressed JSON
        $json = $Result | ConvertTo-Json -Depth 100 -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)

        $fileStream = [System.IO.File]::Create($CacheFilePath)
        $gzipStream = [System.IO.Compression.GZipStream]::new($fileStream, [System.IO.Compression.CompressionLevel]::Optimal)
        try {
            $gzipStream.Write($bytes, 0, $bytes.Length)
        } finally {
            $gzipStream.Dispose()
            $fileStream.Dispose()
        }

        Write-Verbose "Cache saved to compressed JSON: $CacheFilePath"

        return $Result
    } finally {
        Set-Location -Path $originalLocation.Path
        Remove-Item -Path $tempDir.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}


# Custom country data configuration
function Set-CountryLookupConfig {
    [CmdletBinding()]
    param (
        [Parameter()]
        [hashtable]$CustomCountries,

        [Parameter()]
        [hashtable]$CustomAliases
    )

    $cacheFilePath = Join-Path -Path $PSScriptRoot -ChildPath 'ResolveCountry.Data.json.gz'
    if ($null -eq $script:CountryLookupCache) {
        if (Test-Path -Path $cacheFilePath) {
            $script:CountryLookupCache = Import-CompressedCountryCache -Path $cacheFilePath
        } else {
            $script:CountryLookupCache = Update-ResolveCountryData -CacheFilePath $cacheFilePath
        }
    }

    if ($PSBoundParameters.ContainsKey('CustomCountries') -and $null -ne $CustomCountries) {
        foreach ($key in $CustomCountries.Keys) {
            $countryObj = $CustomCountries[$key]
            $cca2Key = $key.ToString().ToUpperInvariant()

            $nativeValues = if ($null -ne $countryObj.name.native) {
                foreach ($nat in $countryObj.name.native.psobject.properties) { $nat.Value.official; $nat.Value.common }
            } else { @() }

            $uniqueVals = @(
                $countryObj.name.official
                $countryObj.name.common
                $nativeValues
                $countryObj.cca2
                $countryObj.cca3
                $countryObj.ccn3
                $countryObj.altSpellings
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

            $script:CountryLookupCache[$cca2Key] = [pscustomobject]@{
                UniqueValues  = $uniqueVals
                CountryObject = $countryObj
            }
        }
    }

    if ($PSBoundParameters.ContainsKey('CustomAliases') -and $null -ne $CustomAliases) {
        foreach ($alias in $CustomAliases.Keys) {
            $targetCca2 = $CustomAliases[$alias].ToString().ToUpperInvariant()

            if ($script:CountryLookupCache.ContainsKey($targetCca2)) {
                $currentValues = [System.Collections.Generic.List[string]]::new([string[]]$script:CountryLookupCache[$targetCca2].UniqueValues)
                if (-not $currentValues.Contains($alias)) {
                    $currentValues.Add($alias)
                    $script:CountryLookupCache[$targetCca2].UniqueValues = $currentValues.ToArray()
                }
            } else {
                Write-Warning "Ziel-Land '$targetCca2' für Alias '$alias' wurde im Cache nicht gefunden."
            }
        }
    }
}


# Search Function with Compressed Disk Cache Integration (Public)
function Resolve-Country {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$InputString,

        [Parameter(Position = 1)]
        [ValidateSet(
            'cca2',
            'cca3',
            'ccn3',
            'name.common',
            'name.native.common',
            'name.native.official',
            'name.official',
            'tld',
            'country'
        )]
        [string]$ReturnType = 'cca2',

        [Parameter()]
        [object]$FallbackValue,

        [Parameter()]
        [switch]$Reload
    )

    begin {
        $cacheFilePath = Join-Path -Path $PSScriptRoot -ChildPath 'ResolveCountry.Data.json.gz'

        # Load cache once for the entire pipeline execution
        if ($Reload -or $null -eq $script:CountryLookupCache) {
            if (-not $Reload -and (Test-Path -Path $cacheFilePath)) {
                Write-Verbose 'Loading country lookup data from compressed JSON cache...'
                $script:CountryLookupCache = Import-CompressedCountryCache -Path $cacheFilePath
            } else {
                Write-Verbose 'Generating country lookup data...'
                $script:CountryLookupCache = Update-ResolveCountryData -CacheFilePath $cacheFilePath
            }
        }
    }

    process {
        if ([string]::IsNullOrWhiteSpace($InputString)) {
            $MatchedKey = $null
        } else {
            $CleanQuery = ConvertTo-NormalizedText $InputString
            $MatchedKey = $null

            # Stage 1: Exact / Case-Insensitive Search
            $FoundMatches = $script:CountryLookupCache.GetEnumerator() | Where-Object {
                $null -ne $_.Value.UniqueValues -and $_.Value.UniqueValues -icontains $InputString
            }

            if ($FoundMatches) {
                $MatchedKey = ($FoundMatches | Select-Object -First 1).Key
            } else {
                # Stage 2: Normalized Exact Search
                $FoundMatches = $script:CountryLookupCache.GetEnumerator() | Where-Object {
                    $uniqueVals = $_.Value.UniqueValues
                    if ($null -eq $uniqueVals) { return $false }

                    foreach ($val in $uniqueVals) {
                        if ((ConvertTo-NormalizedText $val) -eq $CleanQuery) { return $true }
                    }
                    return $false
                }

                if ($FoundMatches) {
                    $MatchedKey = ($FoundMatches | Select-Object -First 1).Key
                } else {
                    # Stage 3: Token-Aware Fuzzy Search
                    $MaxErrorRate = 0.25
                    $MaxDistance = [Math]::Max(1, [Math]::Ceiling($CleanQuery.Length * $MaxErrorRate))

                    $MatchResults = foreach ($entry in $script:CountryLookupCache.GetEnumerator()) {
                        $uniqueVals = $entry.Value.UniqueValues
                        if ($null -eq $uniqueVals) { continue }

                        $bestDistForCountry = 999

                        foreach ($val in $uniqueVals) {
                            $cleanAlias = ConvertTo-NormalizedText $val
                            if ([string]::IsNullOrWhiteSpace($cleanAlias)) { continue }

                            $dist = [FuzzyMatcher]::Levenshtein($CleanQuery, $cleanAlias)
                            if ($dist -lt $bestDistForCountry) {
                                $bestDistForCountry = $dist
                            }

                            $tokens = $cleanAlias -split ' '
                            if ($tokens.Count -gt 1) {
                                foreach ($token in $tokens) {
                                    $tokenDist = [FuzzyMatcher]::Levenshtein($CleanQuery, $token)
                                    if ($tokenDist -lt $bestDistForCountry) {
                                        $bestDistForCountry = $tokenDist
                                    }
                                }
                            }
                        }

                        if ($bestDistForCountry -le $MaxDistance) {
                            [PSCustomObject]@{
                                Key      = $entry.Key
                                Distance = $bestDistForCountry
                            }
                        }
                    }

                    $BestMatch = $MatchResults | Sort-Object Distance | Select-Object -First 1
                    if ($BestMatch) {
                        $MatchedKey = $BestMatch.Key
                    }
                }
            }
        }

        if (-not $MatchedKey) {
            if ($PSBoundParameters.ContainsKey('FallbackValue')) {
                return $FallbackValue
            }
            # Fallback to 'AT' or first entry as default target key
            $MatchedKey = if ($script:CountryLookupCache.ContainsKey('AT')) { 'AT' } else { ($script:CountryLookupCache.Keys | Select-Object -First 1) }
        }

        # Single return switch handles both matched search and default fallback!
        $country = $script:CountryLookupCache[$MatchedKey].CountryObject
        switch ($ReturnType) {
            'cca2' { return $country.cca2 }
            'cca3' { return $country.cca3 }
            'ccn3' { return $country.ccn3 }
            'name.common' { return $country.name.common }
            'name.native.common' { return ($country.name.native.psobject.properties | Select-Object -First 1).Value.common }
            'name.native.official' { return ($country.name.native.psobject.properties | Select-Object -First 1).Value.official }
            'name.official' { return $country.name.official }
            'tld' { return $country.tld }
            'country' { return $country }
        }
    }
}


# Export only the public functions; helper functions remain internal to the module
Export-ModuleMember -Function Resolve-Country, Set-CountryLookupConfig, Update-ResolveCountryData
