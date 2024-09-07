<#
This sample code does the following:
  - If New Outlook is running, ask the user to close it
  - Temporarily disable New Outlook in favor of Classic Outlook
  - Run Set-OutlookSignatures
  - Re-enable New Outlook if it was set as default
  - If New Outlook was running, inform the user that Outlook can be used again

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
#>

[CmdletBinding()] param ()

if ((-not $IsMacOS) -or (-not (Test-Path '/Applications/Microsoft Outlook.app' -PathType Container))) {
    Write-Host 'This script is only supported on macOS with Outlook. Exit.'
    exit 1
}

if ($psISE) {
    Write-Host 'PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
    Write-Host 'Required features are not available in ISE. Exit.' -ForegroundColor Red
    exit 1
}

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

$macOSSignaturesScriptable = @(@($(
            @'
tell application "Microsoft Outlook"
    set guid to do shell script "uuidgen"
    set newSignature to make new signature with properties {name:guid, content:"Set-OutlookSignatures test signature. Please delete."}

    if exists newSignature then
        delete newSignature
        return "Success"
    else
        return "Failure"
    end if
end tell
'@ | osascript *>&1)) | ForEach-Object { $_.tostring() })[0] -eq 'Success'

$DefaultIsRunningNewOutlook = ($(defaults read com.microsoft.Outlook IsRunningNewOutlook *>&1).ToString() -eq 1)

$DefaultEnableNewOutlook = $(defaults read com.microsoft.Outlook EnableNewOutlook *>&1).ToString()
If ($DefaultEnableNewOutlook -inotin @(0, 1, 2, 3)) { $DefaultEnableNewOutlook = 2 } else { $DefaultEnableNewOutlook = [int]$DefaultEnableNewOutlook }

$OutlookWasRunning = $false

# If New Outlook is enabled and running, ask the user to close New Outlook
# and then automatically temporarily disable New Outlook
If ((-not $macOSSignaturesScriptable) -and $DefaultIsRunningNewOutlook) {
    if ((Get-Process | Where-Object { $_.path -ilike '/Applications/Microsoft Outlook.app/*' }).count -gt 0) {
        $OutlookWasRunning = $true

        'display alert "Set-OutlookSignatures" message "To update your Outlook signatures, please close New Outlook at your convenience.\n\nOutlook will then be automatically started in classic mode to update signatures.\n\nYou will be informed when you can use Outlook again." buttons { "OK" } default button 1' | osascript *>$null

        while ((Get-Process | Where-Object { $_.path -ilike '/Applications/Microsoft Outlook.app/*' }).count -gt 0) {
            Start-Sleep -Seconds 1
        }
    }

    # Allow Classic Outlook only
    defaults write com.microsoft.Outlook EnableNewOutlook -integer 0
}


# Start Set-OutlookSignatures - adapt path and parameters to your needs
# & "/path/to/Set-OutlookSignatures/Set-OutlookSignatures.ps1"


# Restore New Outlook setting and inform user
If ((-not $macOSSignaturesScriptable) -and $DefaultIsRunningNewOutlook) {
    # Restore original setting
    defaults write com.microsoft.Outlook EnableNewOutlook -integer $DefaultEnableNewOutlook
    defaults write com.microsoft.Outlook IsRunningNewOutlook -integer 1

    if ($OutlookWasRunning) {
        'display alert "Set-OutlookSignatures" message "Signature updates completed.\n\nYou can now close Classic Outlook.\n\nThe next time you start Outlook, it will start with the New Outlook experience, as configured before." buttons { "OK" } default button 1' | osascript *>$null
    }
}
