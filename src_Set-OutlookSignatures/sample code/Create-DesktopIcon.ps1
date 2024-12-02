<#
This sample code shows how to create a desktop icon in Windows, allowing the user to start Set-OutlookSignatures.

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
#>


[CmdletBinding()] param ()

if ($psISE) {
    Write-Host 'PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
    Write-Host 'Required features are not available in ISE. Exit.' -ForegroundColor Red
    exit 1
}

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

$pathSetOutlookSignatures = (Split-Path $PSScriptRoot -Parent)

if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) {
    if (-not ([System.Management.Automation.PSTypeName]'SetOutlookSignatures.ShellLink').Type) {
        Add-Type -TypeDefinition @'
namespace SetOutlookSignatures
{
    using System;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using System.Text;

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    [CoClass(typeof(CShellLinkW))]
    interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, IntPtr pfd, uint fFlags);
        IntPtr GetIDList();
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxName);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        ushort GetHotKey();
        void SetHotKey(ushort wHotKey);
        uint GetShowCmd();
        void SetShowCmd(uint iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, [Optional] uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    [ClassInterface(ClassInterfaceType.None)]
    class CShellLinkW { }

    public static class ShellLink
    {
        public static void CreateShortcut(
            string lnkPath,
            string targetPath,
            string arguments,
            string workingDirectory,
            string description,
            string iconPath,
            int iconIndex = 0,
            uint showCmd = 1)
        {
            if (string.IsNullOrWhiteSpace(lnkPath))
                throw new ArgumentNullException("lnkPath");

            if (string.IsNullOrWhiteSpace(targetPath))
                throw new ArgumentNullException("targetPath");

            IShellLinkW link = new IShellLinkW();

            link.SetPath(targetPath);

            if (!string.IsNullOrWhiteSpace(arguments))
            {
                link.SetArguments(arguments);
            }
            
            if (!string.IsNullOrWhiteSpace(workingDirectory))
            {
                link.SetWorkingDirectory(workingDirectory);
            }

            if (!string.IsNullOrWhiteSpace(description))
            {
                link.SetDescription(description);
            }

            if (!(iconPath == null))
            {
                link.SetIconLocation(iconPath, iconIndex);
            }

            link.SetShowCmd(showCmd);

            IPersistFile file = (IPersistFile)link;
            file.Save(lnkPath, true);

            Marshal.FinalReleaseComObject(file);
            Marshal.FinalReleaseComObject(link);
        }
    }
}
'@
    }

    [SetOutlookSignatures.ShellLink]::CreateShortcut(
        $(Join-Path -Path $([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)) -ChildPath 'Set Outlook signatures.lnk'), # lnkPath: Full path of the shortcut file to create
        'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe', # targetPath: Full path to the target file (the file the shortcut should open)
        "-File $(Join-Path -Path $pathSetOutlookSignatures -ChildPath 'Set-OutlookSignatures.ps1')", # arguments: Arguments to pass to the target file
        $pathSetOutlookSignatures, # workingDirectory: Full path of the working directory
        'Set Outlook signatures using Set-OutlookSignatures.ps1', # description: Description
        $(Join-Path -Path $($pathSetOutlookSignatures) -ChildPath 'logo/Set-OutlookSignatures Icon.ico'), # iconPath: Full path to the icon file
        0, # iconIndex: Index of the icon within the icon file
        1 # showCmd: Window mode: 1 = Normal, 3 = Maximized, 7 = Minimized
    )
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
