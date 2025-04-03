<#
This sample code shows how to use Intune detect and remediation scripts to deploy and regularly run Set-OutlookSignatures.

See FAQ 'How can I deploy and run Set-OutlookSignatures using Microsoft Intune?' in '.\docs\README' for details

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
#>


[CmdletBinding()] param ()


# Version of Set-OutlookSignatures to use/download
#   Must be a valid tag of a public release from https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases, for example 'XXXVersionStringXXX'
#   When using the Benefactor Circle add-on: Make sure that the same version of Set-OutlookSignatures and the add-on are used
#   Set to '' to not download Set-OutlookSignatures from GitHub but use the version available in $SoftwarePath
#   You can also use 'latest' at your own risk, as the latest version might bring breaking changes with it
$VersionToUse = 'XXXVersionStringXXX'

# Download Set-OutlookSignatures even if already available locally in the required version
#   Is ignored when $VersionToUse is empty
$ForceDownload = $true

# Where to find or download Set-OutlookSignatures locally
#   The path is created if it does not exist
#   You can use local paths or network shares, including SMB file shares in Azure Files
#     SharePoint is not supported
$SoftwarePath = $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath 'Set-OutlookSignatures\Set-OutlookSignatures')

# Parameters for the execution of Set-OutlookSignatures
$SetOutlookSignaturesParameters = @{
    # Add/modify parameters below
    SignatureTemplatePath       = 'c:\path\to\templates' # Path to folder containing the templates
    BenefactorCircleID          = 'xxx'
    BenefactorCircleLicenseFile = 'xxx'
    Verbose                     = $false
}


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


try {
    if (-not [string]::IsNullOrWhiteSpace($VersionToUse)) {
        # Get currently installed version
        $currentVersion = $null

        if (-not (Test-Path $SoftwarePath)) {
            New-Item -Path $SoftwarePath -ItemType Directory
        } else {
            if ((Test-Path (Join-Path -Path $SoftwarePath -ChildPath 'docs\releases.txt'))) {
                try {
                    $currentVersion = @(Get-Content -LiteralPath (Join-Path -Path $SoftwarePath -ChildPath 'docs\releases.txt') | Where-Object { $_ })[-1]
                } catch {
                    $currentVersion = $null
                }
            }
        }

        if ($VersionToUse -ieq 'latest') {
            $VersionToUse = (Invoke-WebRequest -Uri 'https://api.github.com/repos/Set-OutlookSignatures/Set-OutlookSignatures/releases/latest' -UseBasicParsing | ConvertFrom-Json).tag_name
        }

        # Download Set-OutlookSignatures if not already available locally in the required version
        if (($currentVersion -ine $VersionToUse) -or $ForceDownload) {
            $tempFile = New-TemporaryFile | Rename-Item -NewName { [IO.Path]::ChangeExtension($_, '.zip') } -PassThru

            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            try {
                Invoke-WebRequest -Uri "https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/download/$($VersionToUse)/Set-OutlookSignatures_$($VersionToUse).zip" -UseBasicParsing -OutFile $tempFile
            } catch {
                Write-Host $error[0]
    
                Write-Host "Error accessing '$("https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/releases/download/$($VersionToUse)/Set-OutlookSignatures_$($VersionToUse).zip")'."
                Write-Host "Variable '`$VersionToUse' might not be defined correctly (current value: '$($VersionToUse)')."

                exit 1
            }
            $ProgressPreference = $OldProgressPreference

            @(@(Get-ChildItem -LiteralPath $SoftwarePath -Recurse -Force) | Select-Object *, @{Name = 'FolderDepth'; Expression = { $_.DirectoryName.Split('\').Count } } | Sort-Object -Descending -Property FolderDepth, FullName) | Remove-Item -Force -Recurse

            Add-Type -Assembly System.IO.Compression.FileSystem

            $zip = [IO.Compression.ZipFile]::OpenRead($tempFile)

            $entries = $zip.Entries | Where-Object { $_.FullName -ilike "Set-OutlookSignatures_$($VersionToUse)/*" } | Sort-Object

            $entries | ForEach-Object {
                $dest = $(Join-Path -Path $SoftwarePath -ChildPath ($_.FullName -ireplace "^$([regex]::escape("Set-OutlookSignatures_$($VersionToUse)/"))"))

                if ($_.FullName.EndsWith('/')) {
                    if (-not (Test-Path $dest)) {
                        $null = New-Item -Path $dest -ItemType Directory -Force
                    }
                } else {
                    if (-not (Test-Path (Split-Path $dest -Parent))) {
                        $null = New-Item -Path (Split-Path $dest -Parent) -ItemType Directory -Force
                    }

                    [IO.Compression.ZipFileExtensions]::ExtractToFile($_, $dest, $true)
                }
            }

            $zip.Dispose()

            Remove-Item -LiteralPath $tempFile -Force

            if (-not $IsLinux) {
                Get-ChildItem $SoftwarePath -Recurse | Unblock-File
            }
        }
    }


    # Run Set-OutlookSignatures
    & (Join-Path -Path $SoftwarePath -ChildPath 'Set-OutlookSignatures.ps1') @SetOutlookSignaturesParameters

    # End script with exit code from Set-OutlookSignatures
    exit $LASTEXITCODE
} catch {
    $_

    exit 1
}
