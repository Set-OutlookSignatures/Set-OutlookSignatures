<#
This sample code shows how to use Intune detect and remediation scripts to deploy and regularly run Set-OutlookSignatures.

See FAQ 'How can I deploy and run Set-OutlookSignatures using Microsoft Intune?' in '.\docs\README' for details

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
#>

[CmdletBinding()] param ()

# Log file for Set-OutlookSignatures, must be identical in detection and remediation script
$logFile = $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath 'Set-OutlookSignatures_log.txt')

# Interval in hours between runs of Set-OutlookSignatures
$maximumAgeHours = 2


$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot


If (-not (Test-Path -LiteralPath $logFile)) {
    Write-Host 'Log file not found, Set-OutlookSignatures has not yet run.'
    Write-Host 'Exit with error code 1 to trigger remediation script.'

    exit 1
} else {
    If ((Get-Date).AddHours(-$maximumAgeHours) -ge (Get-Item -LiteralPath $logFile).LastWriteTime) {
        Write-Host "Log file found, it is at least $maximumAgeHours hours old."
        Write-Host 'Exit with error code 1 to trigger remediation script.'

        exit 1 # Exit code 1 in the detect script triggers the Intune remediation script
    } else {
        Write-Host "Log file found, it is younger than $maximumAgeHours hours."
        Write-Host 'Exit with error code 0 to not trigger remediation script.'

        exit 0 # Exit code 0 in the detect script does not trigger the Intune remediation script
    }
}
