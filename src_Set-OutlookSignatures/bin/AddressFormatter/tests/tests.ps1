Write-Host 'Import module AddressFormatter'
Import-Module (Split-Path $PSScriptRoot)

Write-Host 'Import module powershell-yaml'
Import-Module (Join-Path $PSScriptRoot '..\nestedModules\powershell-yaml')

$TestCaseFilesCount = 0
$TestCaseCount = 0
$TestCaseErrorCount = 0
$TestCaseErrors = @()

Write-Host 'Running test cases...'

foreach ($TestCaseFile in
    @(
        Get-ChildItem (Join-Path $PSScriptRoot '..\address-formatting\testcases') -Include '*.yaml' -File -Recurse
    )
) {
    $TestCaseFilesCount++

    foreach ($TestCase in @(ConvertFrom-Yaml -Yaml (Get-Content $TestCaseFile.fullname -Raw -Encoding UTF8) -AllDocuments)) {
        $TestCaseCount++

        if ((Split-Path (Split-Path $TestCaseFile.fullname) -Leaf) -ieq 'abbreviations') {
            $result = (Format-PostalAddress -Components $TestCase.components -Abbreviate)
        } else {
            $result = (Format-PostalAddress -Components $TestCase.components)
        }

        $TestCase.expected = $TestCase.expected -replace '\n$', '' # We do not add a trailing newline

        if ($result -ne $TestCase.expected) {
            $TestCaseErrorCount++

            $TestCaseErrors += $TestCaseFile.fullname
            $TestCaseErrors += "  $($TestCase.description)"

            @(
                ' Failed'
                ' Expected lines:'
                $TestCase.expected -split '\r?\n' | ForEach-Object {
                    "        '$($_)'"
                }
                ' Returned lines:'
                $result -split '\r?\n' | ForEach-Object {
                    "        '$($_)'"
                }
            ) | ForEach-Object {
                $TestCaseErrors += $_
            }
        }
    }
}

Write-Host "$TestCaseCount test cases from $TestCaseFilesCount files completed."
Write-Host "  Passed: $($TestCaseCount-$TestCaseErrorCount)/$($TestCaseCount) ($((($TestCaseCount-$TestCaseErrorCount)*100/$TestCaseCount).ToString('F2'))%)"
Write-Host "  Failed: $($TestCaseErrorCount)/$($TestCaseCount) ($(($TestCaseErrorCount*100/$TestCaseCount).ToString('F2'))%)"
$TestCaseErrors | ForEach-Object { Write-Host "    $($_)" }
