# Intune-SetOutlookSignatures-Detect.ps1
# See FAQ 'How can I deploy and run Set-OutlookSignatures using Microsoft Intune?' in '.\docs\README' for details

# Log file for Set-OutlookSignatures, must be identical in detection and remediation script
$logFile = 'c:\path\to\the\user\specific\logfile.txt'

# Interval in hours between runs of Set-OutlookSignatures
$maximumAgeHours = 2


If (-not (Test-Path $logFile)) {
    Write-Host 'Log file not found, Set-OutlookSignatures has not yet run.'
    Write-Host 'Exit with error code 1 to trigger remediation script.'

    exit 1
} else {
    If ((Get-Date).AddHours(-$maximumAgeHours) -ge (Get-Item -LiteralPath $logFile).LastWriteTime) {
        Write-Host "Log file found, it is at least $maximumAgeHours hours old."
        Write-Host 'Exit with error code 1 to trigger remediation script.'

        exit 1
    } else {
        Write-Host "Log file found, it is younger than $maximumAgeHours hours."
        Write-Host 'Exit with error code 0 to not trigger remediation script.'

        exit 0
    }
}
