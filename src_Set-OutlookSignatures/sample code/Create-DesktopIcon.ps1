$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

$WshShell = New-Object -ComObject WScript.Shell  
$Shortcut = $WshShell.CreateShortcut((Join-Path -Path $([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) -ChildPath 'Set Outlook Signatures.lnk'))  
$Shortcut.WorkingDirectory = '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Set-OutlookSignatures'  
$Shortcut.TargetPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'  
$Shortcut.Arguments = "-NoExit -Command ""& '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Set-OutlookSignatures\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Templates\Signatures DOCX' -OOFTemplatePath '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Templates\Out of Office DOCX'"""  
$Shortcut.IconLocation = '\\Long-Server-Name\Long-Share-Name\Long-Folder-Name\Set-OutlookSignatures\logo\Set-OutlookSignatures Icon.ico'  
$Shortcut.Description = 'Set Outlook signatures using Set-OutlookSignatures.ps1'  
$Shortcut.WindowStyle = 1 # 1 = undefined, 3 = maximized, 7 = minimized  
$Shortcut.Hotkey = ''  
$Shortcut.Save()
