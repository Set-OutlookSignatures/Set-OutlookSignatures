# See FAQ 'How can I deploy and run Set-OutlookSignatures using Microsoft Intune?' in '.\docs\README' for details

# Log file for Set-OutlookSignatures, must be identical in detection and remediation script
$logFile = 'c:\path\to\the\user\specific\logfile.txt'

# Version of Set-OutlookSignatures to use/download
$versionToUse = 'v4.10.0'

# Where to find Set-OutlookSignatures locally
$softwarePath = 'c:\path\to\Set-OutlookSignatures'

# Parameters for the later execution of Set-OutlookSignatures.ps1
$parameters = @{
    SignatureTemplatePath       = 'https://URI/to/SharePoint/Libary/with/Templates' # Path to SharePoint document library containing the templates, ini files, ...
    BenefactorCircleId          = 'xxx'
    BenefactorCircleLicenseFile = 'xxx'
}


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

        Get-ChildItem $softwarePath -Recurse | Unblock-File
    }


    # Run Set-OutlookSignatures
    Start-Transcript -LiteralPath $logFile -Force # Required for detection script

    & (Join-Path -Path $softwarePath -ChildPath 'Set-OutlookSignatures.ps1') @parameters

    Stop-Transcript
} catch {
    $error[0]

    exit 1
}
