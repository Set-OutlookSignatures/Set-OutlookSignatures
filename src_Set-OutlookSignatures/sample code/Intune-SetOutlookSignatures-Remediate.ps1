<#
This sample code shows how to use Intune detect and remediation scripts to deploy and regularly run Set-OutlookSignatures.

See FAQ 'How can I deploy and run Set-OutlookSignatures using Microsoft Intune?' in '.\docs\README' for details

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
#>

[CmdletBinding()] param ()

# Log file for Set-OutlookSignatures, must be identical in detection and remediation script
$logFile = $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath 'Set-OutlookSignatures_log.txt')

# Version of Set-OutlookSignatures to use/download
$versionToUse = 'vX.X.X' # Must be a valid tag of a public release at https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases, for example 'v4.12.0'

# Where to find Set-OutlookSignatures locally
$softwarePath = 'c:\path\to\Set-OutlookSignatures'

# Parameters for the later execution of Set-OutlookSignatures.ps1
$parameters = @{
    SignatureTemplatePath       = 'c:\path\to\templates' # Path to folder containing the templates
    # Add more parameters here
    BenefactorCircleID          = 'xxx'
    BenefactorCircleLicenseFile = 'xxx'
}


$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot


try {
    # Get currently installed version
    $currentVersion = $null

    if (-not (Test-Path $softwarePath)) {
        New-Item -Path $softwarePath -ItemType Directory
    } else {
        if ((Test-Path (Join-Path -Path $softwarePath -ChildPath 'docs\releases.txt'))) {
            try {
                $currentVersion = @(Get-Content -LiteralPath (Join-Path -Path $softwarePath -ChildPath 'docs\releases.txt'))[-1]
            } catch {
                $currentVersion = $null
            }
        }
    }

    # Install Set-OutlookSignatures, if not already available in the required version
    if ($currentVersion -ine $versionToUse) {
        $tempFile = New-TemporaryFile | Rename-Item -NewName { [IO.Path]::ChangeExtension($_, '.zip') } -PassThru

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        Invoke-WebRequest -Uri "https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/download/$($versionToUse)/Set-OutlookSignatures_$($versionToUse).zip" -UseBasicParsing -OutFile $tempFile

        $ProgressPreference = $OldProgressPreference

        @(@(Get-ChildItem -LiteralPath $softwarePath -Recurse -Force) | Select-Object *, @{Name = 'FolderDepth'; Expression = { $_.DirectoryName.Split('\').Count } } | Sort-Object -Descending -Property FolderDepth, FullName) | Remove-Item -Force -Recurse

        Add-Type -Assembly System.IO.Compression.FileSystem

        $zip = [IO.Compression.ZipFile]::OpenRead($tempFile)

        $entries = $zip.Entries | Where-Object { $_.FullName -ilike "Set-OutlookSignatures_$($versionToUse)/*" } | Sort-Object

        $entries | ForEach-Object {
            $dest = $(Join-Path -Path $softwarePath -ChildPath ($_.FullName -ireplace "^$([regex]::escape("Set-OutlookSignatures_$($versionToUse)/"))"))

            if (($_.FullName.EndsWith('/')) -or (-not (Test-Path (Split-Path $dest)))) {
                $null = New-Item -Path $dest -ItemType Directory -Force
            } else {
                [IO.Compression.ZipFileExtensions]::ExtractToFile($_, $dest, $true)
            }
        }

        $zip.Dispose()

        Remove-Item -LiteralPath $tempFile -Force

        if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) {
            Get-ChildItem $softwarePath -Recurse | Unblock-File
        }
    }


    # Run Set-OutlookSignatures
    Start-Transcript -LiteralPath $logFile -Force # Required for detection script

    & (Join-Path -Path $softwarePath -ChildPath 'Set-OutlookSignatures.ps1') @parameters

    Stop-Transcript
} catch {
    $error[0]

    exit 1
}
