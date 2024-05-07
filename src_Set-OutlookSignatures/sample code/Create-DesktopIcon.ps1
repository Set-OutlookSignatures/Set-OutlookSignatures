<#
This sample code shows how to create a desktop icon in Windows, allowing the user to start Set-OutlookSignatures.

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
#>


[CmdletBinding()] param ()

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

$pathSetOutlookSignatures = (Split-Path $PSScriptRoot -Parent)

if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut((Join-Path -Path $([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) -ChildPath 'Set Outlook signatures.lnk'))
    $Shortcut.WorkingDirectory = $pathSetOutlookSignatures
    $Shortcut.TargetPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
    $Shortcut.Arguments = "-File $(Join-Path -Path $pathSetOutlookSignatures -ChildPath 'Set-OutlookSignatures.ps1')"
    $Shortcut.IconLocation = $(Join-Path -Path $($pathSetOutlookSignatures) -ChildPath 'logo/Set-OutlookSignatures Icon.ico')
    $Shortcut.Description = 'Set Outlook signatures using Set-OutlookSignatures.ps1'
    $Shortcut.WindowStyle = 1 # 1 = undefined, 3 = maximized, 7 = minimized
    $Shortcut.Hotkey = ''
    $Shortcut.Save()
} elseif ($IsLinux) {
    $tempFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'Set Outlook signatures.desktop'

    @"
[Desktop Entry]
Version=1.0
Type=Application
Name[de]=Outlook-Signaturen setzen
Name[en]=Set Outlook signatures
Comment[de]=Outlook-Signaturen und Abwesenheitstexte mit Set-OutlookSignatures setzen
Comment[en]=Set Outlook signatures and out-of-office replies using Set-OutlookSignatures.ps1
Categories=Office;Utility;Email
Exec=pwsh -File '$(Join-Path -Path $pathSetOutlookSignatures -ChildPath 'Set-OutlookSignatures.ps1')'
Icon=$(Join-Path -Path $($pathSetOutlookSignatures) -ChildPath 'logo/Set-OutlookSignatures Icon.ico')
Terminal=true
"@ | Out-File $tempFile -Encoding UTF8 -Force

    xdg-desktop-icon uninstall (Split-Path $tempFile -Leaf)
    xdg-desktop-icon install --novendor $tempFile
    gio set $(Join-Path -Path ([System.Environment]::GetFolderPath('Desktop')) -ChildPath (Split-Path $tempFile -Leaf)) metadata::trusted true
    chmod a+x $(Join-Path -Path ([System.Environment]::GetFolderPath('Desktop')) -ChildPath (Split-Path $tempFile -Leaf))

    Remove-Item $tempFile -Force
} elseif ($IsMacOS) {
    $desktopFile = $(Join-Path -Path ([System.Environment]::GetFolderPath('Desktop')) -ChildPath 'Set Outlook signatures')
    $tempFile = $null

    @"
#!/usr/bin/env zsh

pwsh -File '$(Join-Path -Path $pathSetOutlookSignatures -ChildPath 'Set-OutlookSignatures.ps1')'
"@ | Out-File $desktopFile -Encoding UTF8 -Force

    chmod a+x $desktopFile

    if (-not (Get-Command fileicon -ErrorAction SilentlyContinue)) {
        try {
            $tempFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "fileicon_$((New-Guid).Guid)"

            Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/mklement0/fileicon/stable/bin/fileicon' -OutFile $tempFile

            chmod a+x $tempFile
        } catch {
            $tempFile = $null
        }
    }

    if (Get-Command fileicon -ErrorAction SilentlyContinue) {
        fileicon set $desktopFile $(Join-Path -Path $($pathSetOutlookSignatures) -ChildPath 'logo/Set-OutlookSignatures Icon.ico') -f
    } elseif ($tempFile -and (Test-Path $tempFile)) {
        & $tempfile set $desktopFile $(Join-Path -Path $($pathSetOutlookSignatures) -ChildPath 'logo/Set-OutlookSignatures Icon.ico') -f

        Remove-Item $tempFile -Force
    }
} else {
    Write-Host 'Unknown Operating System.'
}
