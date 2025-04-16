<#
This sample code shows how to use Intune detect and remediation scripts to deploy and regularly run Set-OutlookSignatures.

See FAQ 'How can I deploy and run Set-OutlookSignatures using Microsoft Intune?' in '.\docs\README' for details

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
#>


[CmdletBinding()] param ()


# Interval in hours between runs of Set-OutlookSignatures
$maximumAgeHours = 2


#
# Do not change anything from here on
#


if ($psISE) {
    Write-Host 'PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
    Write-Host 'Required features are not available in ISE. Exit.' -ForegroundColor Red
    exit 1
}

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

$logFile = (Get-ChildItem -Path $(Join-Path -Path $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\Logs') -ChildPath $('Set-OutlookSignatures_Log_*.txt')) -File -Force -ErrorAction SilentlyContinue | Sort-Object -Culture 127 -Property $_.CreationTime | Select-Object -Last 1).FullName

If ((-not $logFile) -or (-not (Test-Path -LiteralPath $logFile))) {
    Write-Host 'Log file not found, Set-OutlookSignatures has not yet run.'
    Write-Host 'Exit with error code 1 to trigger remediation script.'

    exit 1
} else {
    If ((Get-Date).AddHours(-$maximumAgeHours) -ge (Get-Item -LiteralPath $logFile).LastWriteTime) {
        Write-Host "Log file found, it is at least $maximumAgeHours hours old."
        Write-Host 'Exit with error code 1 to trigger remediation script.'

        # Exit code 1 in the detect script triggers the Intune remediation script
        exit 1
    } else {
        Write-Host "Log file found, it is younger than $maximumAgeHours hours."
        Write-Host 'Exit with error code 0 to not trigger remediation script.'
    }
}

# Exit code 0 in the detect script does not trigger the Intune remediation script
exit 0