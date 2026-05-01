Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


Write-Host
Write-Host 'Import modules'
Write-Host '  AddressFormatter'
Import-Module (Split-Path $PSScriptRoot)
Write-Host '  powershell-yaml'
Import-Module (Join-Path $PSScriptRoot '..\nestedModules\powershell-yaml')

$submoduleID = 'subModules/OpenCageData/address-formatting'
$gitModulesFile = Join-Path $PSScriptRoot '..\.gitmodules'

# Query the .gitmodules file directly using git config
$relativeModulePath = git config -f $gitModulesFile --get "submodule.$($submoduleID).path"

if ($null -ne $relativeModulePath) {
    # Resolve to absolute path relative to the .gitmodules file
    $submoduleRoot = Join-Path (Split-Path $gitModulesFile) $relativeModulePath
    $gitInfo = git -C $submoduleRoot log -1 --format="%h|%cI" 2>$null
} else {
    $gitInfo = "Could not resolve submodule ID '$submoduleID'"
}


Write-Host
Write-Host 'Submodule OpenCageData/address-formatting'
Write-Host "  Commit $($gitInfo.Split('|')[0]), dated $($gitInfo.Split('|')[1])"


Write-Host
Write-Host 'Running test cases'
$TestCaseFilesCount = 0
$TestCaseCount = 0
$TestCaseErrorCount = 0

# Enumerate test files via .NET (no Get-ChildItem pipeline overhead)
$testRoot = Join-Path $PSScriptRoot '..\subModules\OpenCageData\address-formatting\testcases'
$testFiles = [System.IO.Directory]::GetFiles($testRoot, '*.yaml', [System.IO.SearchOption]::AllDirectories)

# Parse every test file ONCE. The previous version parsed each file twice
# (once to count, once to run); powershell-yaml is the dominant cost in this
# script so doing it once roughly halves total parse time.
# Pre-computing $isAbbrev per file also avoids a per-test-case Split-Path call.
$utf8 = [System.Text.Encoding]::UTF8
$parsedFiles = New-Object System.Collections.Generic.List[object]
foreach ($f in $testFiles) {
    $TestCaseFilesCount++
    $text = [System.IO.File]::ReadAllText($f, $utf8)
    $cases = @(ConvertFrom-Yaml -Yaml $text -AllDocuments)
    $TestCaseCount += $cases.Count
    $isAbbrev = ([System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($f))) -ieq 'abbreviations'
    $parsedFiles.Add([pscustomobject]@{
            File     = $f
            Cases    = $cases
            IsAbbrev = $isAbbrev
        })
}

Write-Host "  $TestCaseCount test cases from $TestCaseFilesCount files"

# Pre-compile the two regexes used per test case (compiled once, reused thousands of times)
$rxTrailingNL = [regex]::new('\r?\n$', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$rxAnyNL = [regex]::new('\r?\n', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$envNL = [System.Environment]::NewLine

# Collect errors in a List instead of growing an array via += (which reallocates each time)
$errorsList = New-Object System.Collections.Generic.List[string]

foreach ($entry in $parsedFiles) {
    $file = $entry.File
    $isAbbrev = $entry.IsAbbrev
    foreach ($TestCase in $entry.Cases) {
        if ($isAbbrev) {
            $result = Format-PostalAddress -Components $TestCase.components -Abbreviate
        } else {
            $result = Format-PostalAddress -Components $TestCase.components
        }

        # Normalize expected: drop trailing newline, then convert any \r?\n to env newline.
        $expected = $rxTrailingNL.Replace([string]$TestCase.expected, '')
        $expected = $rxAnyNL.Replace($expected, $envNL)

        if ($result -ne $expected) {
            $TestCaseErrorCount++
            $errorsList.Add($file)
            $errorsList.Add("  $($TestCase.description)")
            $errorsList.Add('    Expected lines:')
            foreach ($line in $rxAnyNL.Split($expected)) { $errorsList.Add("      '$line'") }
            $errorsList.Add('    Returned lines:')
            foreach ($line in $rxAnyNL.Split([string]$result)) { $errorsList.Add("      '$line'") }
        }
    }
}


Write-Host
Write-Host 'Test results'
Write-Host "  Passed: $($TestCaseCount - $TestCaseErrorCount)/$($TestCaseCount) ($((($TestCaseCount - $TestCaseErrorCount) * 100 / $TestCaseCount).ToString('F2')) %)"
Write-Host "  Failed: $($TestCaseErrorCount)/$($TestCaseCount) ($(($TestCaseErrorCount * 100 / $TestCaseCount).ToString('F2')) %)"

foreach ($line in $errorsList) {
    Write-Host "    $line"
}


Write-Host
Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
