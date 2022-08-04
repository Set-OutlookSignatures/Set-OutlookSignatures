<#
.SYNOPSIS
Set-OutlookSignatures XXXVersionStringXXX
Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.

.DESCRIPTION
Signatures and OOF messages can be:
- Generated from templates in DOCX or HTML file format
- Customized with a broad range of variables, including photos, from Active Directory and other sources
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific e-mail addresses, for every primary mailbox across all Outlook profiles
- Assigned time ranges within which they are valid
- Set as default signature for new e-mails, or for replies and forwards (signatures only)
- Set as default OOF message for internal or external recipients (OOF messages only)
- Set in Outlook Web for the currently logged in user
- Centrally managed only or exist along user created signatures (signatures only)
- Copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only)

Set-Outlooksignatures can be executed by users on clients, or on a server without end user interaction.
On clients, it can run as part of the logon script, as scheduled task, or on user demand via a desktop icon, start menu entry, link or any other way of starting a program.
Signatures and OOF messages can also be created and deployed centrally, without end user or client involvement.

Sample templates for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

Simulation mode allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.

The script is designed to work in big and complex environments (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). It works on premises, in hybrid and cloud-only environments.

It is multi-client capable by using different template paths, configuration files and script parameters.

Set-OutlookSignatures requires no installation on servers or clients. You only need a standard file share on a server, and PowerShell and Office on the client.

A documented implementation approach, based on real-life experience implementing the script in a multi-client environment with a five-digit number of mailboxes, contains proven procedures and recommendations for product managers, architects, operations managers, account managers and e-mail and client administrators.
The implementatin approach is suited for service providers as well as for clients, and covers several general overview topics, administration, support, training across the whole lifecycle from counselling to tests, pilot operation and rollout up to daily business.

The script is Free and Open-Source Software (FOSS). It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see '.\docs\LICENSE.txt' for copyright and MIT license details.

.LINK
Github: https://github.com/GruberMarkus/Set-OutlookSignatures

.PARAMETER SignatureTemplatePath
Path to centrally managed signature templates.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
Default value: '.\templates\Signatures DOCX'

.PARAMETER SignatureIniPath
Path to ini file containing signature template tags.
Must be UTF8 encoded.
See '.\templates\Signatures DOCX\_Signatures.ini' for a sample file with further explanations.
Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
The currently logged in user needs at least read access to the path
Default value: '.\templates\Signatures DOCX\_Signatures.ini'

.PARAMETER ReplacementVariableConfigFile
Path to a replacement variable config file.
Must be UTF8 encoded.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
Default value: '.\config\default replacement variables.txt'

.PARAMETER GraphConfigFile
Path to a Graph variable config file.
Must be UTF8 encoded.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signature')
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/config/default graph config.ps1' or '\\server.domain@SSL\SignatureSite\config\default graph config.ps1'
The currently logged in user needs at least read access to the path
Default value: '.\config\default graph config.ps1'

.PARAMETER TrustsToCheckForGroups
List of trusted domains to check for group membership across trusts.
If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.
If a string starts with a minus or dash ("-domain-a.local"), the domain after the dash or minus is removed from the list.
Subdomains of trusted domains are always considered.
Default value: '*'

.PARAMETER DeleteUserCreatedSignatures
Shall the script delete signatures which were created by the user itself?
Default value: $false

.PARAMETER DeleteScriptCreatedSignaturesWithoutTemplate
Shall the script delete signatures which were created by the script before but are no longer available as template?
default value: $true

.PARAMETER SetCurrentUserOutlookWebSignature
Shall the script set the Outlook Web signature of the currently logged in user?
If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used.
Default value: $true

.PARAMETER SetCurrentUserOOFMessage
Shall the script set the Out of Office (OOF) auto reply message of the currently logged in user?
If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used.
Default value: $true

.PARAMETER OOFTemplatePath
Path to centrally managed signature templates.
Local and remote paths are supported.
Local paths can be absolute ('C:\OOF templates') or relative to the script path ('.\templates\Out of Office').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/OOFTemplates' or '\\server.domain@SSL\SignatureSite\OOFTemplates'
The currently logged in user needs at least read access to the path.
Default value: '.\templates\Out of Office DOCX'

.PARAMETER OOFIniPath
Path to ini file containing signature template tags.
Must be UTF8 encoded.
See '.\templates\Out of Office DOCX\_OOF.ini' for a sample file with further explanations.
Local and remote paths are supported. Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
The currently logged in user needs at least read access to the path
Default value: '.\templates\Out of Office DOCX\_OOF.ini'

.PARAMETER AdditionalSignaturePath
An additional path that the signatures shall be copied to.
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.
This way, the user can easily copy-paste the preferred preconfigured signature for use in an e-mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.
Local and remote paths are supported.
Local paths can be absolute ('C:\Outlook signatures') or relative to the script path ('.\Outlook signatures').
WebDAV paths are supported (https only): 'https://server.domain/User' or '\\server.domain@SSL\User'
The currently logged in user needs at least write access to the path.
If the folder or folder structure does not exist, it is created.
Default value: "$([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures'))"

.PARAMETER UseHtmTemplates
With this parameter, the script searches for templates with the extension .htm instead of .docx.
Each format has advantages and disadvantages, please see "Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates." for a quick overview.
Templates in .htm format must be UTF8 encoded.
Default value: $false

.PARAMETER SimulateUser
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged in user.
Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

.PARAMETER SimulateMailboxes
SimulateMailboxes is optional for simulation mode, although highly recommended.
It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.

.PARAMETER GraphCredentialFile
Path to file containing Graph credential which should be used as alternative to other token acquisition methods
Makes only sense in combination with '.\sample code\SimulateAndDeploy.ps1', do not use this parameter for other scenarios
See '.\sample code\SimulateAndDeploy.ps1' for an example how to create this file
Default value: $null

.PARAMETER GraphOnly
Try to connect to Microsoft Graph only, ignoring any local Active Directory.
The default behavior is to try Active Directory first and fall back to Graph.
Default value: $false

.PARAMETER CreateRtfSignatures
Should signatures be created in RTF format?
Default value: $true

.PARAMETER CreateTxtSignatures
Should signatures be created in TXT format?
Default value: $true

.PARAMETER EmbedImagesInHtml
Should images be embedded into HTML files?
Outlook 2016 and newer can handle images embedded directly into an HTML file as BASE64 string ('<img src="data:image/[...]"').
Outlook 2013 and earlier can't handle these embedded images when composing HTML e-mails (there is no problem receiving such e-mails, or when composing RTF or TXT e-mails).
When setting EmbedImagesInHtml to $false, consider setting the Outlook registry value "Send Pictures With Document" to 1 to ensure that images are sent to the recipient (see https://support.microsoft.com/en-us/topic/inline-images-may-display-as-a-red-x-in-outlook-704ae8b5-b9b6-d784-2bdf-ffd96050dfd6 for details).
Default value: $true

.INPUTS
None. You cannot pipe objects to Set-OutlookSignatures.ps1.

.OUTPUTS
Set-OutlookSignatures.ps1 writes the current activities, warnings and error messages to the standard output stream.

.EXAMPLE
Run Set-OutlookSignatures with default values and sample templates
PS> .\Set-OutlookSignatures.ps1

.EXAMPLE
Use custom signature templates and custom ini file
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureIniPath '\\internal.example.com\share\Signature Templates\_Signatures.ini'

.EXAMPLE
Use custom signature templates, ignore trust to internal-test.example.com
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -TrustsToCheckForGroups '*', '-internal-test.example.com'

.EXAMPLE
Use custom signature templates, only check domains/trusts internal-test.example.com and company.b.com
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -TrustsToCheckForGroups 'internal-test.example.com', 'company.b.com'

.EXAMPLE
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. See '.\docs\README.html' for details.
PowerShell.exe -Command "& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -SignatureTemplatePath '\\internal.example.com\share\Signature Templates\_Signatures.ini' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -OOFTemplatePath '\\internal.example.com\share\Signature Templates\_OOF.ini' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1' "

.EXAMPLE
Please see '.\docs\README.html' and https://github.com/GruberMarkus/Set-OutlookSignatures for more details.

.NOTES
Script : Set-OutlookSignatures
Version: XXXVersionStringXXX
Web    : https://github.com/GruberMarkus/Set-OutlookSignatures
License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)
#>


[CmdletBinding(PositionalBinding = $false)]

Param(
    # Path to centrally managed signature templates
    [ValidateNotNullOrEmpty()]
    [string]$SignatureTemplatePath = '.\templates\Signatures DOCX',

    # Path to ini file containing signature template tags
    [ValidateNotNullOrEmpty()]
    [string]$SignatureIniPath = '.\templates\Signatures DOCX\_Signatures.ini',

    # Path to a replacement variable config file.
    [ValidateNotNullOrEmpty()]
    [string]$ReplacementVariableConfigFile = '.\config\default replacement variables.ps1',

    # Path to a Graph variable config file.
    [ValidateNotNullOrEmpty()]
    [string]$GraphConfigFile = '.\config\default graph config.ps1',

    # List of domains/forests to check for group membership across trusts
    [Alias('DomainsToCheckForGroups')]
    [string[]]$TrustsToCheckForGroups = ('*'),

    # Shall the script delete signatures which were created by the user itself?
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $DeleteUserCreatedSignatures = $false,

    # Shall the script delete signatures which were created by the script before but are no longer available as template?
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $DeleteScriptCreatedSignaturesWithoutTemplate = $true,

    # Shall the script set the Outlook Web signature of the currently logged in user?
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $SetCurrentUserOutlookWebSignature = $true,

    # Shall the script set the Out of Office (OOF) auto reply message(s) of the currently logged in user?
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $SetCurrentUserOOFMessage = $true,

    # Path to centrally managed Out of Office (OOF, automatic reply) templates
    [ValidateNotNullOrEmpty()]
    [string]$OOFTemplatePath = '.\templates\Out of Office DOCX',

    # Path to ini file containing OOF template tags
    [ValidateNotNullOrEmpty()]
    [string]$OOFIniPath = '.\templates\Out of Office DOCX\_OOF.ini',

    # An additional path that the signatures shall be copied to
    [string]$AdditionalSignaturePath = $(try { $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures')) }catch {}),

    # Use templates in .HTM file format instead of .DOCX
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $UseHtmTemplates = $false,

    # Simulate another user as currently logged in user
    [Alias('SimulationUser')]
    [validatescript( {
            $tempSimulateUser = $_
            if ($tempSimulateUser -match '^\S+@\S+$|^\S+\\\S+$') {
                $true
            } else {
                throw "'$tempSimulateUser' does not match the required format 'User@Domain' (UPN) or 'Domain\User'."
            }
        }
    )]
    [string]$SimulateUser = $null,

    # Simulate list of mailboxes instead of mailboxes configured in Outlook
    [Alias('SimulationMailboxes')]
    [mailaddress[]]$SimulateMailboxes = ($null),

    # Path to file containing Graph credential which should be used as alternative to other token acquisition methods
    [ValidateNotNullOrEmpty()]
    [string]$GraphCredentialFile = '',

    # Try to connect to Microsoft Graph only, ignoring any local Active Directory.
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $GraphOnly = $false,

    # Create RTF signatures
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $CreateRtfSignatures = $true,

    # Create TXT signatures
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $CreateTxtSignatures = $true,

    # Embed images in HTML
    [ValidateSet(1, 0, '1', '0', 'true', 'false', '$true', '$false', 'yes', 'no')]
    $EmbedImagesInHtml = $true
)


function main {
    Set-Location $PSScriptRoot | Out-Null


    Write-Host
    Write-Host "Script notes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host '  Script : Set-OutlookSignatures'
    Write-Host '  Version: XXXVersionStringXXX'
    Write-Host '  Web    : https://github.com/GruberMarkus/Set-OutlookSignatures'
    Write-Host "  License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)"


    Write-Host
    Write-Host "Check parameters and script environment @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  PowerShell: '$((($($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition), $($PSVersionTable.Platform), $($PSVersionTable.OS)) | Where-Object {$_}) -join "', '")'"

    Write-Host "  PowerShell bitness: $(if ([Environment]::Is64BitProcess -eq $false) {'Non-'})64-bit process on a $(if ([Environment]::Is64OperatingSystem -eq $false) {'Non-'})64-bit operating system"

    Write-Host "  PowerShell parameters: '$ScriptPassedParameters'"

    Write-Host "  Script path: '$PSCommandPath'"

    if ((Test-Path 'variable:IsWindows')) {
        # Automatic variable $IsWindows is available, must be cross-platform PowerShell version v6+
        if ($IsWindows -eq $false) {
            Write-Host "  Your OS: $($PSVersionTable.Platform), $($PSVersionTable.OS), $(Invoke-Expression '(lsb_release -ds || cat /etc/*release || uname -om) 2>/dev/null | head -n1')" -ForegroundColor Red
            Write-Host '  This script is supported on Windows only. Exit.' -ForegroundColor Red
            exit 1
        }
    } else {
        # Automatic variable $IsWindows is not available, must be PowerShell <v6 running on Windows
    }

    if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
        Write-Host "  This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode." -ForegroundColor Red
        Write-Host '  Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
        exit 1
    }

    $script:tempDir = [System.IO.Path]::GetTempPath()
    $script:jobs = New-Object System.Collections.ArrayList
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $Search = New-Object DirectoryServices.DirectorySearcher
    $Search.PageSize = 1000

    $HTMLMarkerTag = '<meta name=data-SignatureFileInfo content="Set-OutlookSignatures">'

    # Connected Files - description and folder name sources:
    #   https://docs.microsoft.com/en-us/windows/win32/shell/manage#connected-files
    #   https://docs.microsoft.com/en-us/office/vba/api/word.defaultweboptions.foldersuffix
    $ConnectedFilesFolderNames = ('.files', '_archivos', '_arquivos', '_bestanden', '_bylos', '_datoteke', '_dosyalar', '_elemei', '_failid', '_fails', '_fajlovi', '_ficheiros', '_fichiers', '_file', '_files', '_fitxategiak', '_fitxers', '_pliki', '_soubory', '_tiedostot', '-Dateien', '-filer')


    Write-Host ('  TrustsToCheckForGroups: ' + ('''' + $($TrustsToCheckForGroups -join ''', ''') + ''''))

    Write-Host "  SignatureTemplatePath: '$SignatureTemplatePath'" -NoNewline
    ConvertPath ([ref]$SignatureTemplatePath)
    CheckPath $SignatureTemplatePath

    Write-Host "  SignatureIniPath: '$SignatureIniPath'" -NoNewline
    if ($SignatureIniPath) {
        ConvertPath ([ref]$SignatureIniPath)
        CheckPath $SignatureIniPath
        $SignatureIniSettings = GetIniContent $SignatureIniPath

        Write-Verbose '    Parsed ini content'
        foreach ($section in $SignatureIniSettings.GetEnumerator()) {
            Write-Verbose "      Ini index #: '$($section.name)'"
            $local:tags = @()
            foreach ($key in $SignatureIniSettings[$($section.name)].GetEnumerator()) {
                if ($key.value) {
                    $local:tags += "$($key.name) = $($key.value)"
                } else {
                    $local:tags += "$($key.name)"
                }
            }
            Write-Verbose "        Tags: [$($local:tags -join '] [')]"
        }
    } else {
        $SignatureIniSettings = @{}
        Write-Host
    }

    Write-Host "  SetCurrentUserOutlookWebSignature: '$SetCurrentUserOutlookWebSignature'"
    if ($SetCurrentUserOutlookWebSignature -in (1, '1', 'true', '$true', 'yes')) {
        $SetCurrentUserOutlookWebSignature = $true
    } else {
        $SetCurrentUserOutlookWebSignature = $false
    }

    Write-Host "  SetCurrentUserOOFMessage: '$SetCurrentUserOOFMessage'"
    if ($SetCurrentUserOOFMessage -in (1, '1', 'true', '$true', 'yes')) {
        $SetCurrentUserOOFMessage = $true
    } else {
        $SetCurrentUserOOFMessage = $false
    }
    if ($SetCurrentUserOOFMessage) {
        Write-Host "  OOFTemplatePath: '$OOFTemplatePath'" -NoNewline
        ConvertPath ([ref]$OOFTemplatePath)
        CheckPath $OOFTemplatePath
        Write-Host "  OOFIniPath: '$OOFIniPath'" -NoNewline
        if ($OOFIniPath) {
            ConvertPath ([ref]$OOFIniPath)
            CheckPath $OOFIniPath
            $OOFIniSettings = GetIniContent $OOFIniPath

            Write-Verbose '    Parsed ini content'
            foreach ($section in $OOFIniSettings.GetEnumerator()) {
                Write-Verbose "      Ini index #: '$($section.name)'"
                $local:tags = @()
                foreach ($key in $OOFIniSettings[$($section.name)].GetEnumerator()) {
                    if ($key.value) {
                        $local:tags += "$($key.name) = $($key.value)"
                    } else {
                        $local:tags += "$($key.name)"
                    }
                }
                Write-Verbose "        Tags: [$($local:tags -join '] [')]"
            }
        } else {
            $OOFIniSettings = @{}
            Write-Host
        }
    }

    Write-Host "  UseHtmTemplates: '$UseHtmTemplates'"
    if ($UseHtmTemplates -in (1, '1', 'true', '$true', 'yes')) {
        $UseHtmTemplates = $true
    } else {
        $UseHtmTemplates = $false
    }

    Write-Host "  GraphOnly: '$GraphOnly'"
    if ($GraphOnly -in (1, '1', 'true', '$true', 'yes')) {
        $GraphOnly = $true
    } else {
        $GraphOnly = $false
    }

    Write-Host "  GraphConfigFile: '$GraphConfigFile'" -NoNewline
    if ($GraphConfigFile) {
        ConvertPath ([ref]$GraphConfigFile)
        CheckPath $GraphConfigFile
        foreach ($line in @(Get-Content -LiteralPath $GraphConfigFile -Encoding UTF8)) {
            Write-Verbose $line
        }
    } else {
        Write-Host
    }

    Write-Host "  GraphCredentialFile: '$GraphCredentialFile'" -NoNewline
    if ($GraphCredentialFile) {
        ConvertPath ([ref]$GraphCredentialFile)
        CheckPath $GraphCredentialFile
        foreach ($line in @(Get-Content -LiteralPath $GraphCredentialFile -Encoding UTF8)) {
            Write-Verbose $line
        }
    } else {
        Write-Host
    }

    Write-Host "  ReplacementVariableConfigFile: '$ReplacementVariableConfigFile'" -NoNewline
    if ($ReplacementVariableConfigFile) {
        ConvertPath ([ref]$ReplacementVariableConfigFile)
        CheckPath $ReplacementVariableConfigFile
        foreach ($line in @(Get-Content -LiteralPath $ReplacementVariableConfigFile -Encoding UTF8)) {
            Write-Verbose $line
        }
    } else {
        Write-Host
    }

    Write-Host "  EmbedImagesInHtml: '$EmbedImagesInHtml'"
    if ($EmbedImagesInHtml -in (1, '1', 'true', '$true', 'yes')) {
        $EmbedImagesInHtml = $true
    } else {
        $EmbedImagesInHtml = $false
    }

    Write-Host "  CreateRtfSignatures: '$CreateRtfSignatures'"
    if ($CreateRtfSignatures -in (1, '1', 'true', '$true', 'yes')) {
        $CreateRtfSignatures = $true
    } else {
        $CreateRtfSignatures = $false
    }

    Write-Host "  CreateTxtSignatures: '$CreateTxtSignatures'"
    if ($CreateTxtSignatures -in (1, '1', 'true', '$true', 'yes')) {
        $CreateTxtSignatures = $true
    } else {
        $CreateTxtSignatures = $false
    }

    Write-Host "  DeleteUserCreatedSignatures: '$DeleteUserCreatedSignatures'"
    if ($DeleteUserCreatedSignatures -in (1, '1', 'true', '$true', 'yes')) {
        $DeleteUserCreatedSignatures = $true
    } else {
        $DeleteUserCreatedSignatures = $false
    }

    Write-Host "  DeleteScriptCreatedSignaturesWithoutTemplate: '$DeleteScriptCreatedSignaturesWithoutTemplate'"
    if ($DeleteScriptCreatedSignaturesWithoutTemplate -in (1, '1', 'true', '$true', 'yes')) {
        $DeleteScriptCreatedSignaturesWithoutTemplate = $true
    } else {
        $DeleteScriptCreatedSignaturesWithoutTemplate = $false
    }

    Write-Host "  AdditionalSignaturePath: '$AdditionalSignaturePath'" -NoNewline
    ConvertPath ([ref]$AdditionalSignaturePath)
    checkpath $AdditionalSignaturePath -create

    Write-Host "  SimulateUser: '$SimulateUser'"

    $tempSimulateMailboxes = $SimulateMailboxes
    [string[]]$SimulateMailboxes = $null
    foreach ($tempSimulateMailbox in $tempSimulateMailboxes) {
        $SimulateMailboxes += $tempSimulateMailbox.Address
    }
    Write-Host ('  SimulateMailboxes: ' + ('''' + $($SimulateMailboxes -join ''', ''') + ''''))


    if ($SimulateUser) {
        Write-Host
        Write-Host 'Simulation mode enabled' -ForegroundColor Yellow
    }


    Write-Host
    Write-Host "Get basic Outlook and Word information @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled, skip Outlook checks' -ForegroundColor Yellow
    } else {
        Write-Host '  Outlook'

        $OutlookRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace 'Outlook.Application.', '') + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))

        try {
            $OutlookFilePath = Get-ChildItem (((Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\WOW6432NODE\CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32" -ErrorAction Stop).'(default)') -split ' \/')[0] -ErrorAction Stop
            $OutlookFileVersion = [System.Version]::Parse((((($OutlookFilePath.versioninfo.fileversion + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))
            $OutlookBitness = GetBitness $OutlookFilePath
        } catch {
            try {
                $OutlookFilePath = Get-ChildItem (((Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32" -ErrorAction Stop).'(default)') -split ' \/')[0] -ErrorAction Stop
                $OutlookFileVersion = [System.Version]::Parse((((($OutlookFilePath.versioninfo.fileversion + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))
                $OutlookBitness = GetBitness $OutlookFilePath
            } catch {
                $OutlookFileVersion = $null
                $OutlookBitness = $null
            }
        }

        if ($OutlookRegistryVersion.major -eq 0) {
            $OutlookRegistryVersion = $null
        } elseif ($OutlookRegistryVersion.major -gt 16) {
            Write-Host "    Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
            exit 1
        } elseif ($OutlookRegistryVersion.major -eq 16) {
            $OutlookRegistryVersion = '16.0'
        } elseif ($OutlookRegistryVersion.major -eq 15) {
            $OutlookRegistryVersion = '15.0'
        } elseif ($OutlookRegistryVersion.major -eq 14) {
            $OutlookRegistryVersion = '14.0'
        } elseif ($OutlookRegistryVersion.major -lt 14) {
            Write-Host "    Outlook version $OutlookRegistryVersion is older than Outlook 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
            exit 1
        }

        $OutlookDisableRoamingSignaturesTemporaryToggle = 0

        if ($null -ne $OutlookRegistryVersion) {
            $OutlookDefaultProfile = (Get-ItemProperty "hkcu:\software\microsoft\office\$OutlookRegistryVersion\Outlook" -ErrorAction SilentlyContinue).DefaultProfile

            foreach ($RegistryFolder in (
                    "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Setup",
                    "registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Setup",
                    "registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$OutlookRegistryVersion\Outlook\Setup",
                    "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$OutlookRegistryVersion\Outlook\Setup"
                )) {
                $x = (Get-ItemProperty $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignaturesTemporaryToggle'
                if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                    $OutlookDisableRoamingSignaturesTemporaryToggle = $x
                }
            }
        } else {
            $OutlookDefaultProfile = $null
        }

        Write-Host "    Registry version: $OutlookRegistryVersion"
        Write-Host "    File version: $OutlookFileVersion"
        if (($OutlookFileVersion -lt '16.0.0.0') -and ($EmbedImagesInHtml -eq $true)) {
            Write-Host '      Outlook 2013 or earlier detected.' -ForegroundColor Yellow
            Write-Host '      Consider parameter ''EmbedImagesInHtml false'' to avoid problems with images in templates.' -ForegroundColor Yellow
            Write-Host '      Microsoft supports Outlook 2013 until April 2023, older versions are already out of support.' -ForegroundColor Yellow
        }
        Write-Host "    Bitness: $OutlookBitness"
        Write-Host "    Default profile: $OutlookDefaultProfile"
        Write-Host "    Roaming signature toggle: $OutlookDisableRoamingSignaturesTemporaryToggle"
    }

    Write-Host '  Word'
    $WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace 'Word.Application.', '') + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))
    if ($WordRegistryVersion.major -eq 0) {
        $WordRegistryVersion = $null
    } elseif ($WordRegistryVersion.major -gt 16) {
        Write-Host "    Word version $WordRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
        exit 1
    } elseif ($WordRegistryVersion.major -eq 16) {
        $WordRegistryVersion = '16.0'
    } elseif ($WordRegistryVersion.major -eq 15) {
        $WordRegistryVersion = '15.0'
    } elseif ($WordRegistryVersion.major -eq 14) {
        $WordRegistryVersion = '14.0'
    } elseif ($WordRegistryVersion.major -lt 14) {
        Write-Host "    Word version $WordRegistryVersion is older than Word 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
        exit 1
    }

    try {
        $WordFilePath = Get-ChildItem (((Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\WOW6432NODE\CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32" -ErrorAction Stop).'(default)') -split ' \/')[0] -ErrorAction Stop
        $WordFileVersion = [System.Version]::Parse(((((((Get-ChildItem (((Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\WOW6432NODE\CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32" -ErrorAction Stop).'(default)') -split ' \/')[0] -ErrorAction Stop)).versioninfo.fileversion + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))
        $WordBitness = GetBitness $WordFilePath
    } catch {
        try {
            $WordFilePath = Get-ChildItem (((Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\CLSID\$((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32" -ErrorAction Stop).'(default)') -split ' \/')[0] -ErrorAction Stop
            $WordFileVersion = [System.Version]::Parse((((($WordFilePath.versioninfo.fileversion + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))
            $WordBitness = GetBitness $WordFilePath
        } catch {
            $WordFileVersion = $null
            $WordBitness = $null
        }
    }

    Write-Host "    Registry version: $WordRegistryVersion"
    Write-Host "    File version: $WordFileVersion"
    Write-Host "    Bitness: $WordBitness"


    Write-Host
    Write-Host "Get Outlook signature file path(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $SignaturePaths = @()
    if ($SimulateUser) {
        $SignaturePaths += $AdditionalSignaturePath
        Write-Host '  Simulation mode enabled. Skip task, use AdditionalSignaturePath instead' -ForegroundColor Yellow
    } else {
        $x = (Get-ItemProperty "hkcu:\software\microsoft\office\$OutlookRegistryVersion\common\general" -ErrorAction SilentlyContinue).'Signatures'
        if ($x) {
            Push-Location ((Join-Path -Path ($env:AppData) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))
            if (Test-Path $x -IsValid) {
                if (-not (Test-Path $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }
                $SignaturePaths += $x
                Write-Host "  $x"
            }
            Pop-Location
        }
    }


    Write-Host
    Write-Host "Get e-mail addresses from Outlook profiles and corresponding registry paths @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $MailAddresses = @()
    $RegistryPaths = @()
    $LegacyExchangeDNs = @()

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled. Skip task, use SimulateMailboxes instead' -ForegroundColor Yellow
        for ($i = 0; $i -lt $SimulateMailboxes.count; $i++) {
            $MailAddresses += $SimulateMailboxes[$i].ToLower()
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''
        }
    } else {
        foreach ($RegistryFolder in @(Get-ItemProperty "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*" -ErrorAction SilentlyContinue | Where-Object { if ($OutlookFileVersion -ge '16.0.0.0') { ($_.'Account Name' -like '*@*.*') } else { (($_.'Account Name' -join ',') -like '*,64,*,46,*') } })) {
            if ($OutlookFileVersion -ge '16.0.0.0') {
                $MailAddresses += ($RegistryFolder.'Account Name').ToLower()
            } else {
                $MailAddresses += (@(ForEach ($char in @(($RegistryFolder.'Account Name' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join '').ToLower()
            }
            $RegistryPaths += $RegistryFolder.PSPath
            if ($RegistryFolder.'Identity Eid') {
                $LegacyExchangeDN = ('/O=' + ((@(foreach ($char in @(($RegistryFolder.'Identity Eid' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join '') -split '/O=')[-1]).ToString().trim()
                if ($LegacyExchangeDN.length -le 3) {
                    $LegacyExchangeDN = ''
                }
            } else {
                $LegacyExchangeDN = ''
            }
            $LegacyExchangeDNs += $LegacyExchangeDN
            Write-Host "  $($RegistryFolder.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $RegistryFolder.PSDrive)"
            Write-Host "    $($MailAddresses[-1])"
        }
    }


    Write-Host
    Write-Host "Enumerate domains @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $x = $TrustsToCheckForGroups
    [System.Collections.ArrayList]$TrustsToCheckForGroups = @()
    if ($GraphOnly -eq $false) {
        # Users own domain/forest is always included
        try {
            $objTrans = New-Object -ComObject 'NameTranslate'
            $objNT = $objTrans.GetType()
            $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
            $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value))) # 12 = ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME
            $y = ([ADSI]"LDAP://$(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1) -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')

            if ($y -ne '') {
                Write-Host "  Current user forest: $y"
                $TrustsToCheckForGroups += $y

                # Other domains - either the list provided, or all outgoing and bidirectional trusts
                if ($x[0] -eq '*') {
                    $Search.SearchRoot = "GC://$($TrustsToCheckForGroups[0])"
                    $Search.Filter = '(ObjectClass=trustedDomain)'

                    foreach ($TrustedDomain in $Search.FindAll()) {
                        # DNS name of this side of the trust (could be the root domain or any subdomain)
                        # $TrustOrigin = ($TrustedDomain.properties.distinguishedname -split ',DC=')[1..999] -join '.'

                        # DNS name of the other side of the trust (could be the root domain or any subdomain)
                        # $TrustName = $TrustedDomain.properties.name

                        # Domain SID of the other side of the trust
                        # $TrustNameSID = (New-Object system.security.principal.securityidentifier($($TrustedDomain.properties.securityidentifier), 0)).tostring()

                        # Trust direction
                        # https://docs.microsoft.com/en-us/dotnet/api/system.directoryservices.activedirectory.trustdirection?view=net-5.0
                        # $TrustDirectionNumber = $TrustedDomain.properties.trustdirection

                        # Trust type
                        # https://docs.microsoft.com/en-us/dotnet/api/system.directoryservices.activedirectory.trusttype?view=net-5.0
                        # $TrustTypeNumber = $TrustedDomain.properties.trusttype

                        # Trust attributes
                        # https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-adts/e9a2d23c-c31e-4a6f-88a0-6646fdb51a3c
                        # $TrustAttributesNumber = $TrustedDomain.properties.trustattributes

                        # Which domains does the current user have access to?
                        # No intra-forest trusts, only bidirectional trusts and outbound trusts

                        if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3)) ) {
                            Write-Host "  Trusted domain: $($TrustedDomain.properties.name)"
                            $TrustsToCheckForGroups += $TrustedDomain.properties.name
                        }
                    }
                }

                for ($a = 0; $a -lt $x.Count; $a++) {
                    if (($a -eq 0) -and ($x[$a] -ieq '*')) {
                        continue
                    }

                    $y = ($x[$a] -replace ('DC=', '') -replace (',', '.'))

                    if ($y -eq $x[$a]) {
                        Write-Host "  User provided domain/forest: $y"
                    } else {
                        Write-Host "  User provided domain/forest: $($x[$a]) -> $y"
                    }

                    if (($a -ne 0) -and ($x[$a] -ieq '*')) {
                        Write-Host '    Entry * is only allowed at first position in list. Skip domain.' -ForegroundColor Red
                        continue
                    }

                    if ($y -match '[^a-zA-Z0-9.-]') {
                        Write-Host '    Allowed characters are a-z, A-Z, ., -. Skip domain.' -ForegroundColor Red
                        continue
                    }

                    if (-not ($y.StartsWith('-'))) {
                        if ($TrustsToCheckForGroups -icontains $y) {
                            Write-Host '    Domain already in list.' -ForegroundColor Yellow
                        } else {
                            $TrustsToCheckForGroups += $y
                        }
                    } else {
                        Write-Host '    Remove domain.'
                        for ($z = 0; $z -lt $TrustsToCheckForGroups.Count; $z++) {
                            if ($TrustsToCheckForGroups[$z] -ilike $y.substring(1)) {
                                $TrustsToCheckForGroups[$z] = ''
                            }
                        }
                    }
                }


                Write-Host
                Write-Host "Check for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                CheckADConnectivity $TrustsToCheckForGroups 'LDAP' '  ' | Out-Null


                Write-Host
                Write-Host "Check for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                CheckADConnectivity $TrustsToCheckForGroups 'GC' '  ' | Out-Null
            } else {
                Write-Host '  Problem connecting to logged in user''s Active Directory (no error message, but forest root domain name is empty), assuming Graph/Azure AD from now on.' -ForegroundColor Yellow
                $GraphOnly = $true
            }
        } catch {
            $y = ''
            Write-Verbose $error[0]
            Write-Host '  Problem connecting to logged in user''s Active Directory (see verbose stream for error message), assuming Graph/Azure AD from now on.' -ForegroundColor Yellow
            $GraphOnly = $true
        }
    } else {
        Write-Host "  Parameter GraphOnly set to '$GraphOnly', ignore user's Active Directory in favor of Graph/Azure AD."
    }


    Write-Host
    Write-Host "Get AD properties of currently logged in user and assigned manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (-not $SimulateUser) {
        Write-Host '  Currently logged in user'
    } else {
        Write-Host "  Simulate '$SimulateUser' as currently logged in user" -ForegroundColor Yellow
    }

    if ($GraphOnly -eq $false) {
        if ($null -ne $TrustsToCheckForGroups[0]) {
            try {
                if (-not $SimulateUser) {
                    $Search.SearchRoot = "GC://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"
                    $ADPropsCurrentUser = $Search.FindOne().Properties
                } else {
                    try {
                        $objTrans = New-Object -ComObject 'NameTranslate'
                        $objNT = $objTrans.GetType()
                        $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                        $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, $SimulateUser))
                        $SimulateUserDN = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1)
                        $Search.SearchRoot = "GC://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"
                        $ADPropsCurrentUser = $Search.FindOne().Properties
                    } catch {
                        Write-Verbose $error[0]
                        Write-Host "    Simulation user '$($SimulateUser)' not found. Exit." -ForegroundColor REd
                        exit 1
                    }
                }
            } catch {
                $ADPropsCurrentUser = $null
                Write-Host '    Problem connecting to Active Directory, or user is a local user. Exit.' -ForegroundColor Red
                $error[0]
                exit 1
            }
        }
    }

    if (($GraphOnly -eq $true) -or
        (($GraphOnly -eq $false) -and ($ADPropsCurrentUser.msexchrecipienttypedetails -ge 2147483648) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true))) -or
        (($GraphOnly -eq $false) -and ($null -eq $ADPropsCurrentUser))) {
        Write-Host "    Set up environment for connection to Microsoft Graph @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $script:CurrentUser = (Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value)\IdentityCache\$(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value)" -Name 'UserName' -ErrorAction SilentlyContinue)
        $script:msalPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))
        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\msal.ps')) -Destination (Join-Path -Path $script:msalPath -ChildPath 'msal.ps') -Recurse -ErrorAction SilentlyContinue
        Get-ChildItem $script:msalPath -Recurse | Unblock-File
        try {
            Import-Module (Join-Path -Path $script:msalPath -ChildPath 'msal.ps') -ErrorAction Stop
        } catch {
            Write-Host '        Problem importing MSAL.PS module. Exit.' -ForegroundColor Red
            $error[0]
            exit 1
        }

        if (Test-Path -Path $GraphConfigFile -PathType Leaf) {
            try {
                Write-Host "      Execute config file '$GraphConfigFile'"
                . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $GraphConfigFile -Encoding UTF8 -Raw)))
            } catch {
                Write-Host "        Problem executing content of '$GraphConfigFile'. Exit." -ForegroundColor Red
                $error[0]
                exit 1
            }
        } else {
            Write-Host "      Problem connecting to or reading from file '$GraphConfigFile'. Exit." -ForegroundColor Red
            exit 1
        }

        if ($($PSVersionTable.PSEdition) -ieq 'Desktop') {
            Write-Host "      MSAL.PS Graph token cache: '$([TokenCacheHelper]::CacheFilePath)'"
        }

        $GraphToken = GraphGetToken
        if ($GraphToken.error -eq $false) {
            Write-Verbose "Graph Token: $($GraphToken.AccessToken)"
            if ($SimulateUser) {
                $script:CurrentUser = $SimulateUser
            }
            if ($null -eq $script:CurrentUser) {
                $script:CurrentUser = (GraphGetMe).me.userprincipalname
            }

            $x = (GraphGetUserProperties $script:CurrentUser)
            if ($x.error -eq $false) {
                $AADProps = $x.properties
                $ADPropsCurrentUser = [PSCustomObject]@{}

                foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                    $z = $AADProps

                    foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                        $z = $z.$y
                    }

                    $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                }

                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $script:CurrentUser).photo
                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $script:CurrentUser).properties.userprincipalname

            } else {
                Write-Host "      Problem getting data for '$($script:CurrentUser)' from Microsoft Graph. Exit." -ForegroundColor Red
                $error[0]
                exit 1
            }
        } else {
            Write-Host '      Problem connecting to Microsoft Graph. Exit.' -ForegroundColor Red
            $error[0]
            exit 1
        }

        if (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true)) {
            if ($GraphCredentialFile) {
                $ExoToken = $GraphToken.AccessTokenExo
            } else {
                $ExoToken = ($script:msalClientApp | Get-MsalToken -LoginHint $script:CurrentUser -Scopes 'https://outlook.office.com/EWS.AccessAsUser.All' -Silent).accessToken
            }
            Write-Verbose "EXO Token: $ExoToken"

            if (-not $ExoToken) {
                Write-Host '      Problem connecting to Exchange Online with Graph token. Exit.' -ForegroundColor Red
                $error[0]
                exit 1
            }
        }
    }


    if ((($SetCurrentUserOutlookWebSignature -eq $true) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses -inotcontains $ADPropsCurrentUser.mail) -and (-not $SimulateUser)) {
        # OOF and/or Outlook web signature must be set, but user does not seem to have a mailbox in Outlook
        # Maybe this is a pure Outlook Web user, so we will add a helper entry
        # This entry fakes the users mailbox in his default Outlook profile, so it gets the highest priority later
        Write-Host "    User's mailbox not found in Outlook profiles, but Outlook Web signature and/or OOF message should be set. Add dummy mailbox entry." -ForegroundColor Yellow
        $script:CurrentUserDummyMailbox = $true
        $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname)) + $SignaturePaths
        $MailAddresses = @($ADPropsCurrentUser.mail.tolower()) + $MailAddresses
        $RegistryPaths = @("hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\") + $RegistryPaths
        $LegacyExchangeDNs = @('') + $LegacyExchangeDNs
    } else {
        $script:CurrentUserDummyMailbox = $false
    }
    if ($ADPropsCurrentUser.distinguishedname) {
        Write-Host "    $($ADPropsCurrentUser.distinguishedname)"
    } else {
        Write-Host "    $($ADPropsCurrentUser.userprincipalname)"
    }

    $CurrentUserSIDs = @()
    if (($ADPropsCurrentUser.objectsid -ne '') -and ($null -ne $ADPropsCurrentUser.objectsid)) {
        if ($GraphOnly) {
            $CurrentUserSIDs += $ADPropsCurrentUser.objectsid.tostring()
        } else {
            $CurrentUserSIDs += (New-Object System.Security.Principal.SecurityIdentifier $($ADPropsCurrentUser.objectsid), 0).value.tostring()
        }
    }
    foreach ($SidHistorySid in @($ADPropsCurrentUser.sidhistory | Where-Object { ($_ -ne '') -and ($null -ne $_ ) })) {
        if ($GraphOnly) {
            $CurrentUserSIDs += $SidHistorySid.tostring()
        } else {
            $CurrentUserSIDs += (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value.tostring()
        }
    }

    if (-not $SimulateUser) {
        Write-Host '  Manager of currently logged in user'
    } else {
        Write-Host '  Manager of simulated currently logged in user'
    }
    if ($null -ne $TrustsToCheckForGroups[0]) {
        try {
            $Search.SearchRoot = "GC://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
            $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"
            $ADPropsCurrentUserManager = $Search.FindOne().Properties
        } catch {
            $ADPropsCurrentUserManager = $null
        }
    } else {
        if ($ADPropsCurrentUser.manager) {
            $AADProps = (GraphGetUserProperties $ADPropsCurrentUser.manager).properties
            $ADPropsCurrentUserManager = [PSCustomObject]@{}

            foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                $z = $AADProps

                foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                    $z = $z.$y
                }

                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
            }

            $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsCurrentUserManager.userprincipalname).photo
            $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null
        }
    }

    if ($ADPropsCurrentUserManager) {
        if ($ADPropsCurrentUserManager.distinguishedname) {
            Write-Host "    $($ADPropsCurrentUserManager.distinguishedname)"
        } else {
            Write-Host "    $($ADPropsCurrentUserManager.userprincipalname)"
        }
    } else {
        Write-Host '    No manager found'
    }


    Write-Host
    Write-Host "Get AD properties of each mailbox @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $ADPropsMailboxes = @()
    $ADPropsMailboxesUserDomain = @()

    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        Write-Host "  Mailbox '$($MailAddresses[$AccountNumberRunning])'"

        $UserDomain = ''
        $ADPropsMailboxes += $null
        $ADPropsMailboxesUserDomain += $null

        if ((($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '') -or ($($MailAddresses[$AccountNumberRunning]) -ne ''))) {
            if ($null -ne $TrustsToCheckForGroups[0]) {
                # Loop through domains until the first one knows the legacyExchangeDN or the proxy address
                for ($DomainNumber = 0; (($DomainNumber -lt $TrustsToCheckForGroups.count) -and ($UserDomain -eq '')); $DomainNumber++) {
                    if (($TrustsToCheckForGroups[$DomainNumber] -ne '')) {
                        Write-Host "    Search for mailbox user object in domain/forest '$($TrustsToCheckForGroups[$DomainNumber])': " -NoNewline
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustsToCheckForGroups[$DomainNumber])")
                        if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                            $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(|(legacyExchangeDN=$($LegacyExchangeDNs[$AccountNumberRunning]))(&(legacyExchangeDN=*)(proxyaddresses=x500:$($LegacyExchangeDNs[$AccountNumberRunning])))))"
                        } elseif (($($MailAddresses[$AccountNumberRunning]) -ne '')) {
                            $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(legacyExchangeDN=*)(proxyaddresses=smtp:$($MailAddresses[$AccountNumberRunning])))"
                        }
                        $u = $Search.FindAll()
                        if ($u.count -eq 0) {
                            Write-Host 'Not found'
                        } elseif ($u.count -gt 1) {
                            Write-Host 'Ignore due to multiple matches' -ForegroundColor Red
                            foreach ($SingleU in $u) {
                                Write-Host "      $($SingleU.path)" -ForegroundColor Yellow
                            }
                            $LegacyExchangeDNs[$AccountNumberRunning] = ''
                            $MailAddresses[$AccountNumberRunning] = ''
                            $UserDomain = $null
                        } else {
                            # Connect to Domain Controller (LDAP), as Global Catalog (GC) does not have all attributes,
                            # for example tokenGroups including domain local groups
                            $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"
                            $ADPropsMailboxes[$AccountNumberRunning] = $Search.FindOne().Properties
                            $UserDomain = $TrustsToCheckForGroups[$DomainNumber]
                            $ADPropsMailboxesUserDomain[$AccountNumberRunning] = $TrustsToCheckForGroups[$DomainNumber]
                            $LegacyExchangeDNs[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].legacyexchangedn
                            $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()
                            Write-Host 'Found'
                            Write-Host "      $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                        }
                    }
                }
                if (-not $ADPropsMailboxes[$AccountNumberRunning]) {
                    $LegacyExchangeDNs[$AccountNumberRunning] = ''
                    $UserDomain = $null
                }
            } else {
                $AADProps = (GraphGetUserProperties $($MailAddresses[$AccountNumberRunning])).properties

                $ADPropsMailboxes[$AccountNumberRunning] = [PSCustomObject]@{}

                if ($AADProps) {
                    foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                        $z = $AADProps

                        foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                            $z = $z.$y
                        }

                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                    }

                    $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).photo
                    $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).properties.userprincipalname
                    $LegacyExchangeDNs[$AccountNumberRunning] = 'dummy'
                    $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()
                } else {
                    $LegacyExchangeDNs[$AccountNumberRunning] = ''
                    $UserDomain = $null
                }
            }
        } else {
            $ADPropsMailboxes[$AccountNumberRunning] = $null
        }
    }


    Write-Host
    Write-Host "Sort mailbox list: User's primary mailbox, mailboxes in default Outlook profile, others @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    # Get users primary mailbox
    $p = $null
    # First, check if the user has a mail attribute set
    if ($ADPropsCurrentUser.mail) {
        Write-Host "  AD mail attribute of currently logged in user: $($ADPropsCurrentUser.mail)"
        for ($i = 0; $i -lt $LegacyExchangeDNs.count; $i++) {
            if (($LegacyExchangeDNs[$i]) -and (($ADPropsMailboxes[$i].proxyaddresses) -contains $('SMTP:' + $ADPropsCurrentUser.mail))) {
                $p = $i
                break
            }
        }
        if ($p -ge 0) {
            Write-Host '    Matching mailbox found'
        } else {
            Write-Host '    No matching mailbox found' -ForegroundColor Yellow
        }
    } else {
        Write-Host '  AD mail attribute of currently logged in user is empty' -NoNewline
        if ($null -ne $TrustsToCheckForGroups[0]) {
            Write-Host ', searching msExchMasterAccountSid'
            # No mail attribute set, check for match(es) of user's objectSID and mailbox's msExchMasterAccountSid
            for ($i = 0; $i -lt $MailAddresses.count; $i++) {
                if ($ADPropsMailboxes[$i].msexchmasteraccountsid) {
                    if ((New-Object System.Security.Principal.SecurityIdentifier $ADPropsMailboxes[$i].msexchmasteraccountsid[0], 0).value -iin $CurrentUserSIDs) {
                        if ($p -ge 0) {
                            # $p already set before, there must be at least two matches, so set it to -1
                            $p = -1
                        } elseif (-not $p) {
                            $p = $i
                        }
                    }
                }
            }
            if ($p -ge 0) {
                Write-Host "    One matching mailbox found: $MailAddresses[$i]"
            } elseif ($null -eq $p) {
                Write-Host '    No matching mailbox found' -ForegroundColor Yellow
            } else {
                Write-Host '    Multiple matching mailboxes found, no prioritization possible' -ForegroundColor Yellow
            }
        } else {
            Write-Host
        }
    }

    $MailboxNewOrder = @()
    $PrimaryMailboxAddress = $null

    if ($p -ge 0) {
        $MailboxNewOrder += $p
        $PrimaryMailboxAddress = $MailAddresses[$p]
    }

    for ($i = 0; $i -le $RegistryPaths.length - 1; $i++) {
        if (($RegistryPaths[$i] -ilike "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\*") -and ($i -ne $p)) {
            $MailboxNewOrder += $i
        }
    }

    for ($i = 0; $i -le $RegistryPaths.length - 1; $i++) {
        if (($RegistryPaths[$i] -notlike "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\*") -and ($i -ne $p)) {
            $MailboxNewOrder += $i
        }
    }

    foreach ($VariableName in ('RegistryPaths', 'MailAddresses', 'LegacyExchangeDNs', 'ADPropsMailboxesUserDomain', 'ADPropsMailboxes')) {
        (Get-Variable -Name $VariableName).value = (Get-Variable -Name $VariableName).value[$MailboxNewOrder]
    }
    Write-Host '  Mailbox priority (highest to lowest)'
    foreach ($MailAddress in $MailAddresses) {
        Write-Host "    $MailAddress"
    }


    $TemplateFilesGroupSIDsOverall = @{}
    foreach ($SigOrOOF in ('signature', 'OOF')) {
        if (($SigOrOOF -eq 'OOF') -and ($SetCurrentUserOOFMessage -eq $false)) {
            break
        }
        Write-Host
        Write-Host "Get all $SigOrOOF template files and categorize them @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $TemplateFilesCommon = @{}
        $TemplateFilesGroup = @{}
        $TemplateFilesGroupFilePart = @{}
        $TemplateFilesMailbox = @{}
        $TemplateFilesMailboxFilePart = @{}
        $TemplateFilesDefaultnewOrInternal = @{}
        $TemplateFilesDefaultreplyfwdOrExternal = @{}

        $TemplateTemplatePath = Get-Variable -Name "$($SigOrOOF)TemplatePath" -ValueOnly
        $TemplateIniPath = Get-Variable -Name "$($SigOrOOF)IniPath" -ValueOnly
        $TemplateIniSettings = Get-Variable -Name "$($SigOrOOF)IniSettings" -ValueOnly

        $TemplateFiles = @((Get-ChildItem -LiteralPath $TemplateTemplatePath -File -Filter $(if ($UseHtmTemplates) { '*.htm' } else { '*.docx' })) | Sort-Object)
        if ($TemplateIniPath -ne '') {
            foreach ($Enumerator in $TemplateIniSettings.GetEnumerator().name) {
                if ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotin $TemplateFiles.name)) {
                        Write-Host "  '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' (ini index #$($Enumerator)) found in ini but not in signature template path, please check" -ForegroundColor Yellow
                    }
                }
            }

            foreach ($TemplateFile in $TemplateFiles) {
                if ($TemplateFile.name -inotin @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })) {
                    Write-Host "  '$($TemplateFile.name)' found in signature template path but not in ini, please check" -ForegroundColor Yellow
                }
            }

            try {
                $TemplateFilesSortCulture = (@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortCulture']
            } catch {
                $TemplateFilesSortCulture = $null
            }

            # Populate template files in the most complicated way first: SortOrder 'AsInThisFile'
            # This also considers that templates can be referenced multiple times in the INI file
            # If the setting in the ini file is different, we only need to sort $TemplateFiles
            $TemplateFiles = @($TemplateFiles | Where-Object { $_.name -iin @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] }) } | Select-Object *, TemplateIniSettingsIndex)
            $TemplateFilesSortOrder = @()
            $TemplateFilesIniIndex = @()

            if ($TemplateFiles) {
                foreach ($Enumerator in $TemplateIniSettings.GetEnumerator().name) {
                    if (@($TemplateFiles.name) -icontains $TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                        for ($x = 0; $x -lt $TemplateFiles.count; $x++) {
                            if ($TemplateFiles[$x].name -ieq $TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                                $TemplateFilesSortOrder += $x
                                $TemplateFilesIniIndex += $Enumerator
                            }
                        }
                    }
                }

                $TemplateFiles = @($TemplateFiles[$TemplateFilesSortOrder] | Select-Object *)

                if ($TemplateFiles.count -gt 0) {
                    foreach ($index In 0..($TemplateFiles.Count - 1)) {
                        $TemplateFiles[$index].TemplateIniSettingsIndex = $TemplateFilesIniIndex[$index]
                    }
                }

                if (($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' })) {
                    switch ((@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortOrder']) {
                        { $_ -iin 'AsInThisFile', 'AsListed' } {
                            # nothing to do, $TemplateFiles is already correctly populated and sorted
                            break
                        }

                        { $_ -iin ('a', 'asc', 'ascending', 'az', 'a-z', 'a..z', 'up') } {
                            $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                            break
                        }

                        { $_ -iin ('d', 'des', 'desc', 'descending', 'za', 'z-a', 'z..a', 'dn', 'down') } {
                            $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } } -Descending)
                            break
                        }

                        default {
                            # same as 'ascending'
                            $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                        }
                    }
                } else {
                    $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                }
            }
        }

        foreach ($TemplateFile in $TemplateFiles) {
            $TemplateIniSettingsIndex = $TemplateFile.TemplateIniSettingsIndex
            $TemplateFilesGroupSIDs = @{}
            Write-Host ("  '$($TemplateFile.Name)' (ini index #$($TemplateIniSettingsIndex))")
            if ($TemplateIniSettings[$TemplateIniSettingsIndex]['<Set-OutlookSignatures template>'] -ieq $TemplateFile.name) {
                $TemplateFilePart = ($TemplateIniSettings[$TemplateIniSettingsIndex].GetEnumerator().Name -join '] [')
                if ($TemplateFilePart) {
                    $TemplateFilePart = ($TemplateFilePart -split '\] \[' | Where-Object { $_ -inotin ('OutlookSignatureName', '<Set-OutlookSignatures template>') }) -join '] ['
                    $TemplateFilePart = '[' + $TemplateFilePart + ']'
                    $TemplateFilePart = $TemplateFilePart -replace '\[\]', ''
                }
                if ($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']) {
                    $TemplateFileTargetName = ($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'] + $(if ($UseHtmTemplates) { '.htm' } else { '.docx' }))
                } else {
                    $TemplateFileTargetName = $TemplateFile.Name
                }
            } else {
                $TemplateFilePart = ''
                $TemplateFileTargetName = $TemplateFile.Name
            }

            Write-Host "    Outlook signature name: '$([System.IO.Path]::ChangeExtension($TemplateFileTargetName, $null) -replace '\.$')'"

            $TemplateFilePartRegexTimeAllow = '\[(?!-:)\d{12}-\d{12}\]'
            $TemplateFilePartRegexTimeDeny = '\[-:\d{12}-\d{12}\]'
            $TemplateFilePartRegexGroupAllow = '\[(?!-:)\S+?(?<!]) .+?\]'
            $TemplateFilePartRegexGroupDeny = '\[-:\S+?(?<!]) .+?\]'
            $TemplateFilePartRegexMailaddressAllow = '\[(?!-:)(\S+?)@(\S+?)\.(\S+?)\]'
            $TemplateFilePartRegexMailaddressDeny = '\[-:(\S+?)@(\S+?)\.(\S+?)\]'
            if ($SigOrOOF -ieq 'signature') {
                $TemplateFilePartRegexDefaultneworinternal = '(?i)\[DefaultNew\]'
                $TemplateFilePartRegexDefaultreplyfwdorexternal = '(?i)\[DefaultReplyFwd\]'
            } else {
                $TemplateFilePartRegexDefaultneworinternal = '(?i)\[internal\]'
                $TemplateFilePartRegexDefaultreplyfwdorexternal = '(?i)\[external\]'
            }
            $TemplateFilePartRegexKnown = '(' + (($TemplateFilePartRegexTimeAllow, $TemplateFilePartRegexTimeDeny, $TemplateFilePartRegexGroupAllow, $TemplateFilePartRegexGroupDeny, $TemplateFilePartRegexMailaddressAllow, $TemplateFilePartRegexMailaddressDeny, $TemplateFilePartRegexDefaultneworinternal, $TemplateFilePartRegexDefaultreplyfwdorexternal) -join '|') + ')'

            # time based template
            $TemplateFileTimeActive = $true
            if (($TemplateFilePart -match $TemplateFilePartRegexTimeAllow) -or ($TemplateFilePart -match $TemplateFilePartRegexTimeDeny)) {
                Write-Host '    Time based template'
                if (([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexTimeAllow).captures.value).count -gt 0) {
                    $TemplateFileTimeActive = $false
                } else {
                    $TemplateFileTimeActive = $true
                }
                foreach ($TemplateFilePartTag in @((([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexTimeAllow).captures.value) + ([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexTimeDeny).captures.value)) | Where-Object { $_ })) {
                    Write-Host "      $($TemplateFilePartTag): " -NoNewline
                    try {
                        if (-not ($TemplateFilePartTag.startswith('[-:'))) {
                            $DateTimeTagStart = [System.DateTime]::ParseExact(($TemplateFilePartTag.tostring().Substring(1, 12)), 'yyyyMMddHHmm', $null)
                            $DateTimeTagEnd = [System.DateTime]::ParseExact(($TemplateFilePartTag.tostring().Substring(14, 12)), 'yyyyMMddHHmm', $null)

                            if (((Get-Date) -ge $DateTimeTagStart) -and ((Get-Date) -le $DateTimeTagEnd)) {
                                Write-Host 'Current DateTime is in allowed range'
                                $TemplateFileTimeActive = $true
                            } else {
                                Write-Host 'Current DateTime is not in allowed range'
                            }
                        } else {
                            $DateTimeTagStart = [System.DateTime]::ParseExact(($TemplateFilePartTag.tostring().Substring(3, 12)), 'yyyyMMddHHmm', $null)
                            $DateTimeTagEnd = [System.DateTime]::ParseExact(($TemplateFilePartTag.tostring().Substring(16, 12)), 'yyyyMMddHHmm', $null)

                            if (((Get-Date) -ge $DateTimeTagStart) -and ((Get-Date) -le $DateTimeTagEnd)) {
                                Write-Host 'Current DateTime is in denied range'
                                $TemplateFileTimeActive = $false
                            } else {
                                Write-Host 'Current DateTime is not in denied range'
                            }
                        }
                    } catch {
                        Write-Host 'Invalid DateTime, ignore tag' -ForegroundColor Red
                    }
                }
                if ($TemplateFileTimeActive -eq $true) {
                    Write-Host "      Current DateTime is in allowed time ranges, use $SigOrOOF template"
                } else {
                    Write-Host "      Current DateTime is not in allowed time ranges, ignore $SigOrOOF template" -ForegroundColor Yellow
                }
            }
            if ($TemplateFileTimeActive -ne $true) {
                continue
            }

            # common template
            if (($TemplateFilePart -notmatch $TemplateFilePartRegexGroupAllow) -and ($TemplateFilePart -notmatch $TemplateFilePartRegexMailaddressAllow)) {
                Write-Host '    Common template (no group or e-mail address allow tags specified)'
                if (-not $TemplateFilesCommon.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesCommon.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesCommon[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                }
                $TemplateClassificationDisplayOrder = ('group', 'mail')
            } elseif ($TemplateFilePart -match $TemplateFilePartRegexGroupAllow) {
                $TemplateClassificationDisplayOrder = ('group', 'mail')
            } elseif ($TemplateFilePart -match $TemplateFilePartRegexMailaddressAllow) {
                $TemplateClassificationDisplayOrder = ('mail', 'group')
            }

            foreach ($TemplateClassificationDisplayOrderEntry in $TemplateClassificationDisplayOrder) {
                # group specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'group') {
                    if (($TemplateFilePart -match $TemplateFilePartRegexGroupAllow) -or ($TemplateFilePart -match $TemplateFilePartRegexGroupDeny)) {
                        foreach ($TemplateFilePartTag in @((([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexGroupAllow).captures.value) + ([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexGroupDeny).captures.value)) | Where-Object { $_ })) {
                            if (-not $TemplateFilesGroup.ContainsKey($TemplateIniSettingsIndex)) {
                                if ($TemplateFilePart -match $TemplateFilePartRegexGroupAllow) {
                                    Write-Host '    Group specific template'
                                } else {
                                    Write-Host '    Group specific exclusions'
                                }
                                $TemplateFilesGroup.add($TemplateIniSettingsIndex, @{})
                                $TemplateFilesGroup[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                            }
                            Write-Host "      $($TemplateFilePartTag) = " -NoNewline
                            $NTName = (((($TemplateFilePartTag -replace '\[', '') -replace '^-:', '') -replace '\]$', '') -replace '(.*?) (.*)', '$1\$2')

                            # Check cache (only contains [xxx], not [-:xxx])
                            if ($TemplateFilePartTag.startswith('[-:')) {
                                if ($TemplateFilesGroupSIDsOverall.ContainsKey(($TemplateFilePartTag -replace '^\[-:', '['))) {
                                    $TemplateFilesGroupSIDs.add($TemplateFilePartTag, ('-:' + $TemplateFilesGroupSIDsOverall[($TemplateFilePartTag -replace '^\[-:', '[')]))
                                }
                            } else {
                                if ($TemplateFilesGroupSIDsOverall.ContainsKey($TemplateFilePartTag)) {
                                    $TemplateFilesGroupSIDs.add($TemplateFilePartTag, $TemplateFilesGroupSIDsOverall[$TemplateFilePartTag])
                                }
                            }

                            if ((-not $TemplateFilesGroupSIDs.ContainsKey($TemplateFilePartTag))) {
                                if (($null -ne $TrustsToCheckForGroups[0]) -and (-not ($NTName.startswith('AzureAD\', 'CurrentCultureIgnorecase')))) {
                                    try {
                                        if ($TemplateFilePartTag.startswith('[-:')) {
                                            $TemplateFilesGroupSIDs.add($TemplateFilePartTag, ('-:' + (New-Object System.Security.Principal.NTAccount($NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value))
                                            $TemplateFilesGroupSIDsOverall.add(($TemplateFilePartTag -replace '^\[-:', '['), (New-Object System.Security.Principal.NTAccount($NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value)
                                        } else {
                                            $TemplateFilesGroupSIDs.add($TemplateFilePartTag, (New-Object System.Security.Principal.NTAccount($NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value)
                                            $TemplateFilesGroupSIDsOverall.add($TemplateFilePartTag, (New-Object System.Security.Principal.NTAccount($NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value)
                                        }
                                    } catch {
                                        # No group with this sAMAccountName found. Maybe it's a display name?
                                        try {
                                            Write-Verbose $error[0]
                                            $objTrans = New-Object -ComObject 'NameTranslate'
                                            $objNT = $objTrans.GetType()
                                            $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                                            $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (4, ($NTName -split '\\')[1]))
                                            if ($TemplateFilePartTag.startswith('[-:')) {
                                                $TemplateFilesGroupSIDs.add($TemplateFilePartTag, ('-:' + ((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value))
                                                $TemplateFilesGroupSIDsOverall.add(($TemplateFilePartTag -replace '^\[-:', '['), ((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                                            } else {
                                                $TemplateFilesGroupSIDs.add($TemplateFilePartTag, ((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                                                $TemplateFilesGroupSIDsOverall.add($TemplateFilePartTag, ((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                                            }
                                        } catch {
                                            Write-Verbose $error[0]
                                        }
                                    }
                                } else {
                                    $tempFilterOrder = @(
                                        "((onPremisesNetBiosName eq '$($NTName.Split('\')[0])') and (onPremisesSamAccountName eq '$($NTName.Split('\')[1])'))"
                                        "((onPremisesNetBiosName eq '$($NTName.Split('\')[0])') and (displayName eq '$($NTName.Split('\')[1])'))"
                                        "(proxyAddresses/any(x:x eq 'smtp:$($NTName.Split('\')[1])'))"
                                        "(mailNickname eq '$($NTName.Split('\')[1])')"
                                        "(displayName eq '$($NTName.Split('\')[1])')"
                                    )
                                    ForEach ($tempFilter in $tempFilterOrder) {
                                        $tempResults = (GraphFilterGroups $tempFilter)
                                        if (($tempResults.error -eq $false) -and ($tempResults.groups.count -eq 1 )) {
                                            if ($TemplateFilePartTag.startswith('[-:')) {
                                                $TemplateFilesGroupSIDs.add($TemplateFilePartTag, ('-:' + $tempResults.groups[0].securityidentifier))
                                                $TemplateFilesGroupSIDsOverall.add(($TemplateFilePartTag -replace '^\[-:', '['), $tempResults.groups[0].securityidentifier)
                                            } else {
                                                $TemplateFilesGroupSIDs.add($TemplateFilePartTag, $tempResults.groups[0].securityidentifier)
                                                $TemplateFilesGroupSIDsOverall.add($TemplateFilePartTag, $tempResults.groups[0].securityidentifier)
                                            }
                                            break
                                        }
                                    }
                                }
                            }

                            if ($TemplateFilesGroupSIDs.containskey($TemplateFilePartTag)) {
                                if ($null -ne $TemplateFilesGroupSIDs[$TemplateFilePartTag]) {
                                    Write-Host "$($TemplateFilesGroupSIDs[$TemplateFilePartTag] -replace '^-:', '')"
                                    $TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] + '[' + $TemplateFilesGroupSIDs[$TemplateFilePartTag] + ']')
                                } else {
                                    Write-Host 'Not found, please check' -ForegroundColor Yellow
                                }
                            } else {
                                Write-Host 'Not found, please check' -ForegroundColor Yellow
                                if ($TemplateFilePartTag.startswith('[-:')) {
                                    $TemplateFilesGroupSIDsOverall.add(($TemplateFilePartTag -replace '^\[-:', '['), $null)
                                } else {
                                    $TemplateFilesGroupSIDsOverall.add($TemplateFilePartTag, $null)
                                }

                            }
                        }
                    }
                }

                # mailbox specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'mail') {
                    if (($TemplateFilePart -match $TemplateFilePartRegexMailaddressAllow) -or ($TemplateFilePart -match $TemplateFilePartRegexMailaddressDeny)) {
                        foreach ($TemplateFilePartTag in @((([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexMailaddressAllow).captures.value) + ([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexMailaddressDeny).captures.value)) | Where-Object { $_ })) {
                            if (-not $TemplateFilesMailbox.ContainsKey($TemplateIniSettingsIndex)) {
                                if ($TemplateFilePart -match $TemplateFilePartRegexmailaddressAllow) {
                                    Write-Host '    Mailbox specific template'
                                } else {
                                    Write-Host '    Mailbox specific exclusions'
                                }
                                $TemplateFilesMailbox.add($TemplateIniSettingsIndex, @{})
                                $TemplateFilesMailbox[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                            }
                            Write-Host "      $($TemplateFilePartTag)"
                            $TemplateFilesMailboxFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesMailboxFilePart[$TemplateIniSettingsIndex] + $TemplateFilePartTag)
                        }
                    }
                }
            }

            # DefaultNew, DefaultReplyFwd, Internal, External
            if ($TemplateFilePart -match $TemplateFilePartRegexDefaultneworinternal) {
                foreach ($TemplateFilePartTag in @(([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexDefaultneworinternal).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '    Default signature for new e-mails'
                        Write-Host "      $($TemplateFilePartTag)"
                    } else {
                        Write-Host '    Default internal OOF message'
                        Write-Host "      $($TemplateFilePartTag)"
                    }
                }

                if (-not $TemplateFilesDefaultnewOrInternal.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesDefaultnewOrInternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultnewOrInternal[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                }
            }

            if ($TemplateFilePart -match $TemplateFilePartRegexDefaultreplyfwdorexternal) {
                foreach ($TemplateFilePartTag in @(([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexDefaultreplyfwdorexternal).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '    Default signature for replies and forwards'
                        Write-Host "      $($TemplateFilePartTag)"
                    } else {
                        Write-Host '    Default external OOF message'
                        Write-Host "      $($TemplateFilePartTag)"
                    }
                }

                if (-not $TemplateFilesDefaultreplyfwdOrExternal.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesDefaultreplyfwdOrExternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultreplyfwdOrExternal[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                }
            }

            if ($SigOrOOF -ieq 'OOF') {
                if (($TemplateFilePart -notmatch $TemplateFilePartRegexDefaultreplyfwdorexternal) -and ($TemplateFilePart -notmatch $TemplateFilePartRegexDefaultneworinternal)) {
                    $TemplateFilesDefaultnewOrInternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultnewOrInternal[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                    Write-Host '    Default internal OOF message (neither internal nor external tag specified)'
                    $TemplateFilesDefaultreplyfwdOrExternal.add($TemplateFile.FullName, $TemplateFileTargetName)
                    Write-Host '    Default external OOF message (neither internal nor external tag specified)'
                }
            }

            # unknown tags
            $x = ($TemplateFilePart -replace $TemplateFilePartRegexKnown, '').trim()
            if ($x) {
                Write-Host '    Unknown tags, please check' -ForegroundColor yellow
                Write-Host "      $x"
            }

            Set-Variable -Name "$($SigOrOOF)Files" -Value $TemplateFiles
            Set-Variable -Name "$($SigOrOOF)FilesCommon" -Value $TemplateFilesCommon
            Set-Variable -Name "$($SigOrOOF)FilesGroup" -Value $TemplateFilesGroup
            Set-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -Value $TemplateFilesGroupFilePart
            Set-Variable -Name "$($SigOrOOF)FilesMailbox" -Value $TemplateFilesMailbox
            Set-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -Value $TemplateFilesMailboxFilePart
            if ($SigOrOOF -ieq 'signature') {
                $SignatureFilesDefaultNew = $TemplateFilesDefaultnewOrInternal
                $SignatureFilesDefaultReplyFwd = $TemplateFilesDefaultreplyfwdOrExternal
            } else {
                $OOFFilesInternal = $TemplateFilesDefaultnewOrInternal
                $OOFFilesExternal = $TemplateFilesDefaultreplyfwdOrExternal
            }
        }
    }


    Write-Host
    Write-Host "Start Word background process @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (($UseHtmTemplates -eq $true) -and (($CreateRtfSignatures -eq $false) -and ($CreateTxtSignatures -eq $false))) {
        Write-Host '  Do not start Word: UseHtmTemplates = $true, CreateRtfSignatures = $false, CreateTxtSignatures = $false'
    } else {
        # Start Word dummy object, start real Word object, close dummy object - this seems to avoid a rare problem where a manually started Word instance connects to the Word process created by the script
        try {
            $script:COMWordDummy = New-Object -ComObject Word.Application

            $script:COMWord = New-Object -ComObject Word.Application

            if ($script:COMWordDummy) {
                $script:COMWordDummy.Quit([ref]$false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWordDummy) | Out-Null
                Remove-Variable COMWordDummy -Scope 'script'
            }

            $script:COMWordShowFieldCodesOriginal = $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes

            Add-Type -Path (Get-ChildItem -LiteralPath ((Join-Path -Path ($env:SystemRoot) -ChildPath 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word')) -Filter 'Microsoft.Office.Interop.Word.dll' -Recurse | Select-Object -ExpandProperty FullName -Last 1)
        } catch {
            Write-Host '  Word not installed or not working correctly. Exit.' -ForegroundColor Red
            $error[0]
            exit 1
        }
    }


    # Process each e-mail address only once
    $script:SignatureFilesDone = @()
    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        if (($AccountNumberRunning -le $MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])) -and ($($MailAddresses[$AccountNumberRunning]) -like '*@*')) {
            Write-Host
            Write-Host "Mailbox $($MailAddresses[$AccountNumberRunning]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

            $UserDomain = ''

            Write-Host "  Get group membership of mailbox @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            if ($($ADPropsMailboxesUserDomain[$AccountNumberRunning])) {
                Write-Host "    $($ADPropsMailboxesUserDomain[$AccountNumberRunning]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            }
            $GroupsSIDs = @()

            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                $ADPropsCurrentMailbox = $ADPropsMailboxes[$AccountNumberRunning]
                if ($null -ne $TrustsToCheckForGroups[0]) {
                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")
                    try {
                        $Search.filter = "(distinguishedname=$($ADPropsCurrentMailbox.manager))"
                        $ADPropsCurrentMailboxManager = ([ADSI]"$(($Search.FindOne()).path)").Properties
                    } catch {
                        $ADPropsCurrentMailboxManager = @()
                    }

                    $UserDomain = $ADPropsMailboxesUserDomain[$AccountNumberRunning]
                    $SIDsToCheckInTrusts = @()

                    if ($ADPropsCurrentMailbox.objectsid) {
                        $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier $($ADPropsCurrentMailbox.objectsid), 0).value.tostring()
                    }

                    foreach ($SidHistorySid in @($ADPropsCurrentMailbox.sidhistory | Where-Object { $_ })) {
                        $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier $SidHistorySid, 0).value.tostring()
                    }

                    try {
                        # Security groups, no matter if enabled for mail or not
                        $UserAccount = [ADSI]"LDAP://$($ADPropsCurrentMailbox.distinguishedname)"
                        $UserAccount.GetInfoEx(@('tokengroups'), 0)
                        foreach ($sidBytes in $UserAccount.Properties.tokengroups) {
                            $sid = New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)
                            $GroupsSIDs += $sid.tostring()
                            $SIDsToCheckInTrusts += $sid.tostring()
                            Write-Verbose "      $sid"
                        }

                        # Distribution groups (static only)
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$(($($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.')")
                        $Search.filter = "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648))(member:1.2.840.113556.1.4.1941:=$($ADPropsCurrentMailbox.distinguishedname)))"
                        foreach ($DistributionGroup in $search.findall()) {
                            if ($DistributionGroup.properties.objectsid) {
                                $sid = (New-Object System.Security.Principal.SecurityIdentifier $($DistributionGroup.properties.objectsid), 0).value.tostring()
                                Write-Verbose "      $sid"
                                $GroupsSIDs += $sid.tostring()
                                $SIDsToCheckInTrusts += $sid.tostring()
                            }

                            foreach ($SidHistorySid in @($DistributionGroup.properties.sidhistory | Where-Object { $_ })) {
                                $sid = (New-Object System.Security.Principal.SecurityIdentifier $$SidHistorySid, 0).value.tostring()
                                Write-Verbose "      $sid"
                                $GroupsSIDs += $sid.tostring()
                                $SIDsToCheckInTrusts += $sid.tostring()
                            }
                        }
                    } catch {
                        Write-Host "      Error getting group information from $((($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.'), check firewalls, DNS and AD trust" -ForegroundColor Red
                        $error[0]
                    }

                    # Loop through all domains to check if the mailbox account has a group membership there
                    # Across a trust, a user can only be added to a domain local group.
                    # Domain local groups can not be used outside their own domain, so we don't need to query recursively
                    if ($SIDsToCheckInTrusts.count -gt 0) {
                        $LdapFilterSIDs = '(|'
                        foreach ($SidToCheckInTrusts in $SIDsToCheckInTrusts) {
                            try {
                                $SidHex = @()
                                $ot = New-Object System.Security.Principal.SecurityIdentifier($SidToCheckInTrusts)
                                $c = New-Object 'byte[]' $ot.BinaryLength
                                $ot.GetBinaryForm($c, 0)
                                foreach ($char in $c) {
                                    $SidHex += $('\{0:x2}' -f $char)
                                }
                                # Foreign Security Principals have an objectSID, but no sIDHistory
                                # The sIDHistory of the current mailbox is part of $SIDsToCheckInTrusts and therefore also considered in $LdapFilterSIDs
                                $LdapFilterSIDs += ('(objectsid=' + $($SidHex -join '') + ')')
                            } catch {
                                Write-Host '      Error creating LDAP filter for search across trusts.' -ForegroundColor Red
                                $error[0]
                            }
                        }
                        $LdapFilterSIDs += ')'
                    } else {
                        $LdapFilterSIDs = ''
                    }

                    if ($LdapFilterSids -ilike '*(objectsid=*') {
                        for ($DomainNumber = 0; $DomainNumber -lt $TrustsToCheckForGroups.count; $DomainNumber++) {
                            if (($TrustsToCheckForGroups[$DomainNumber] -ne '') -and ($TrustsToCheckForGroups[$DomainNumber] -ine $UserDomain) -and ($UserDomain -ne '')) {
                                Write-Host "    $($TrustsToCheckForGroups[$DomainNumber]) (mailbox group membership across trusts, takes some time) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustsToCheckForGroups[$DomainNumber])")
                                $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                                foreach ($fsp in $Search.FindAll()) {
                                    if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                                        # Foreign Security Principals do not have the tokengroups attribute
                                        # We need to switch to another, slower search method
                                        # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member
                                        # A Foreign Security Principal ist created in each (sub)domain, in which it is granted permissions,
                                        # and it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal.
                                        Write-Host "      Found $($fsp.properties.cn) in $((($fsp.path -split ',DC=')[1..999] -join '.'))"
                                        try {
                                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                            $Search.filter = "(&(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                            foreach ($group in $Search.findall()) {
                                                $sid = New-Object System.Security.Principal.SecurityIdentifier($group.properties.objectsid[0], 0)
                                                $GroupsSIDs += $sid.tostring()
                                                Write-Verbose "        $sid"
                                            }
                                        } catch {
                                            Write-Host "        Error: $($error[0].exception)" -ForegroundColor red
                                        }
                                    }
                                }
                            }
                        }
                    }
                } else {
                    try {
                        $AADProps = $null
                        if ($ADPropsCurrentMailbox.manager) {
                            $AADProps = (GraphGetUserProperties $ADPropsCurrentMailbox.manager).properties

                            $ADPropsCurrentMailboxManager = [PSCustomObject]@{}

                            foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                                $z = $AADProps

                                foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                    $z = $z.$y
                                }

                                $ADPropsCurrentMailboxManager | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z
                            }

                            $ADPropsCurrentMailboxManager | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsCurrentMailboxManager.userprincipalname).photo
                            $ADPropsCurrentMailboxManager | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null
                        }
                        Write-Host '    Microsoft Graph'
                        foreach ($sid in @((GraphGetUserTransitiveMemberOf $ADPropsCurrentMailbox.userPrincipalName).memberof.securityidentifier)) {
                            $GroupsSIDs += $sid
                            Write-Verbose "      $sid"
                        }
                    } catch {
                        $ADPropsCurrentMailboxManager = @()
                        Write-Host '    Skipping, mailbox not in Microsoft Graph.' -ForegroundColor yellow
                    }
                }
            } else {
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor yellow
            }

            Write-Host "  Get SMTP addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $CurrentMailboxSMTPAddresses = @()
            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                foreach ($ProxyAddress in $ADPropsCurrentMailbox.proxyaddresses) {
                    if ([string]$ProxyAddress -ilike 'smtp:*') {
                        $CurrentMailboxSMTPAddresses += [string]$ProxyAddress -ireplace 'smtp:', ''
                        Write-Verbose "    $($CurrentMailboxSMTPAddresses[-1])"
                    }
                }
            } else {
                $CurrentMailboxSMTPAddresses += $($MailAddresses[$AccountNumberRunning])
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor Yellow
                Write-Host '    Use mailbox name as single known SMTP address' -ForegroundColor Yellow
            }

            Write-Host "  Get data for replacement variables @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $ReplaceHash = @{}
            if (Test-Path -Path $ReplacementVariableConfigFile -PathType Leaf) {
                try {
                    Write-Host "    Execute config file '$ReplacementVariableConfigFile'"
                    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $ReplacementVariableConfigFile -Encoding UTF8 -Raw)))
                } catch {
                    Write-Host "    Problem executing content of '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                    $error[0]
                    exit 1
                }
            } else {
                Write-Host "    Problem connecting to or reading from file '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                exit 1
            }
            foreach ($replaceKey in @($replaceHash.Keys | Sort-Object)) {
                if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                    if ($($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): $($replaceHash[$replaceKey])"
                    }
                } else {
                    if ($null -ne $($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): Photo available"
                    }
                }
            }

            # Export pictures if available
            $CURRENTMAILBOXMANAGERPHOTOGUID = (New-Guid).guid
            $CURRENTMAILBOXPHOTOGUID = (New-Guid).guid
            $CURRENTUSERMANAGERPHOTOGUID = (New-Guid).guid
            $CURRENTUSERPHOTOGUID = (New-Guid).guid

            foreach ($VariableName in (('$CURRENTMAILBOXMANAGERPHOTO$', $CURRENTMAILBOXMANAGERPHOTOGUID) , ('$CURRENTMAILBOXPHOTO$', $CURRENTMAILBOXPHOTOGUID), ('$CURRENTUSERMANAGERPHOTO$', $CURRENTUSERMANAGERPHOTOGUID), ('$CURRENTUSERPHOTO$', $CURRENTUSERPHOTOGUID))) {
                if ($null -ne $ReplaceHash[$VariableName[0]]) {
                    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        $ReplaceHash[$VariableName[0]] | Set-Content -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -AsByteStream -Force
                    } else {
                        $ReplaceHash[$VariableName[0]] | Set-Content -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -Encoding Byte -Force
                    }
                }
            }

            if (
                # Outlook is installed
                # and $OutlookFileVersion is high enough (exact value is unknown yet) or it is a suiting beta version (-or (($OutlookFileVersion -ge '16.0.13430.20000') -and ($OutlookFileVersion.revision -in 20000..20199)))
                # and $OutlookDisableRoamingSignaturesTemporaryToggle equals 0
                # and the mailbox is in the cloud ((connected to AD AND $ADPropsCurrentMailbox.msexchrecipienttypedetails is like \*remote\*) OR (connected to Graph and $ADPropsCurrentMailbox is not like \*remote\*))
                # and the current mailbox is the personal mailbox of the currently logged in user
                ($null -ne $OutlookFileVersion) `
                    -and (($OutlookFileVersion -ge [system.version]::parse('99.0.99999.99999'))) `
                    -and ($OutlookDisableRoamingSignaturesTemporaryToggle -eq 0) `
                    -and ((($null -ne $ADPropsCurrentMailbox.msexchrecipienttypedetails) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -ilike 'remote*')) -or ($null -ne $ADPropsCurrentMailbox.mailboxsettings)) `
                    -and ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress)
            ) {
                # Microsoft signature roaming available
                $CurrentMailboxUseSignatureRoaming = $true
            } else {
                $CurrentMailboxUseSignatureRoaming = $false
            }


            EvaluateAndSetSignatures


            # Delete photos from file system
            foreach ($VariableName in (('$CURRENTMAILBOXMANAGERPHOTO$', $CURRENTMAILBOXMANAGERPHOTOGUID) , ('$CURRENTMAILBOXPHOTO$', $CURRENTMAILBOXPHOTOGUID), ('$CURRENTUSERMANAGERPHOTO$', $CURRENTUSERMANAGERPHOTOGUID), ('$CURRENTUSERPHOTO$', $CURRENTUSERPHOTOGUID))) {
                Remove-Item -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -Force -ErrorAction SilentlyContinue
                $ReplaceHash.Remove($VariableName[0])
                $ReplaceHash.Remove(($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$')
            }

        }

        # Set OOF message and Outlook Web signature
        if (((($SetCurrentUserOutlookWebSignature -eq $true) -and ($CurrentMailboxUseSignatureRoaming -eq $false)) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress)) {
            if ((-not $SimulateUser) ) {
                Write-Host "  Set up environment for connection to Outlook Web @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                $script:dllPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))
                try {
                    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS.NetStandard\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:dllPath -Force
                        Unblock-File -LiteralPath $script:dllPath
                    } else {
                        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\EWS\Microsoft.Exchange.WebServices.dll')) -Destination $script:dllPath -Force
                        Unblock-File -LiteralPath $script:dllPath
                    }
                } catch {
                }

                $error.clear()

                try {
                    Import-Module -Name $script:dllPath -Force -ErrorAction Stop
                    $exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
                    Write-Host "  Connect to Outlook Web @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    try {
                        Write-Host '    Try Windows Integrated Auth'
                        $exchService.UseDefaultCredentials = $true
                        $exchService.AutodiscoverUrl($PrimaryMailboxAddress, { $true }) | Out-Null
                    } catch {
                        try {
                            Write-Host '    Try OAuth with Autodiscover'
                            $exchService.UseDefaultCredentials = $false
                            $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $ExoToken
                            $exchService.AutodiscoverUrl($PrimaryMailboxAddress, { $true }) | Out-Null
                        } catch {
                            Write-Host '    Try OAuth with fixed URL'
                            $exchService.UseDefaultCredentials = $false
                            $exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $ExoToken
                            $exchService.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
                        }
                    }

                    $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchservice, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar)
                    if ($Calendar.DisplayName) {
                        $error.clear()
                    } else {
                        Write-Host '    Could not connect to Outlook Web, although the EWS DLL threw no error.' -ForegroundColor Red
                        throw
                    }
                } catch {
                    Write-Host "    Error connecting to Outlook Web: $_" -ForegroundColor Red

                    if ($SetCurrentUserOutlookWebSignature) {
                        Write-Host '    Outlook Web signature can not be set' -ForegroundColor Red
                        $SetCurrentUserOutlookWebSignature = $false
                    }

                    if ($SetCurrentUserOOFMessage -and (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648))) {
                        Write-Host '    Out of Office (OOF) auto reply message(s) can not be set' -ForegroundColor Red
                        $SetCurrentUserOOFMessage = $false
                    }
                }
            } else {
                $error.Clear()
            }

            if ($SetCurrentUserOutlookWebSignature -and ($CurrentMailboxUseSignatureRoaming -eq $false)) {
                Write-Host "  Set Outlook Web signature @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                if ($SimulateUser) {
                    Write-Host '    Simulation mode enabled, skip task' -ForegroundColor Yellow
                } else {
                    # If this is the primary mailbox, set OWA signature
                    for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                        if ($MailAddresses[$j] -ieq $PrimaryMailboxAddress) {
                            try {
                                if ($script:CurrentUserDummyMailbox -ne $true) {
                                    $TempNewSig = Get-ItemPropertyValue -LiteralPath $RegistryPaths[$j] -Name 'New Signature'
                                    if ($OutlookFileVersion -lt '16.0.0.0') {
                                        $TempNewSig = @(foreach ($char in @(($TempNewSig -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join ''
                                    }
                                } else {
                                    $TempNewSig = $script:CurrentUserDummyMailboxDefaultSigNew
                                }
                            } catch {
                                $TempNewSig = ''
                            }
                            try {
                                if ($script:CurrentUserDummyMailbox -ne $true) {
                                    $TempReplySig = Get-ItemPropertyValue -LiteralPath $RegistryPaths[$j] -Name 'Reply-Forward Signature'
                                    if ($OutlookFileVersion -lt '16.0.0.0') {
                                        $TempReplySig = @(foreach ($char in @(($TempReplySig -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$(char)" }) -join ''
                                    }
                                } else {
                                    $TempReplySig = $script:CurrentUserDummyMailboxDefaultSigReply
                                }
                            } catch {
                                $TempReplySig = ''
                            }
                            if (($TempNewSig -eq '') -and ($TempReplySig -eq '')) {
                                Write-Host '    No default signatures defined, nothing to do'
                                $TempOWASigFile = $null
                                $TempOWASigSetNew = $null
                                $TempOWASigSetReply = $null
                            }

                            if (($TempNewSig -ne '') -and ($TempReplySig -eq '')) {
                                Write-Host "    Only default signature for new e-mails is set: '$TempNewSig'"
                                $TempOWASigFile = $TempNewSig
                                $TempOWASigSetNew = $true
                                $TempOWASigSetReply = $false
                            }

                            if (($TempNewSig -eq '') -and ($TempReplySig -ne '')) {
                                Write-Host "    Only default signature for reply/forward is set: '$TempReplySig'"
                                $TempOWASigFile = $TempReplySig
                                $TempOWASigSetNew = $false
                                $TempOWASigSetReply = $true
                            }


                            if ((($TempNewSig -ne '') -and ($TempReplySig -ne '')) -and ($TempNewSig -ine $TempReplySig)) {
                                Write-Host "    Different default signatures for new and reply/forward set, use new one: '$TempNewSig'"
                                $TempOWASigFile = $TempNewSig
                                $TempOWASigSetNew = $true
                                $TempOWASigSetReply = $false
                            }

                            if ((($TempNewSig -ne '') -and ($TempReplySig -ne '')) -and ($TempNewSig -ieq $TempReplySig)) {
                                Write-Host "    Same default signature for new and reply/forward: '$TempNewSig'"
                                $TempOWASigFile = $TempNewSig
                                $TempOWASigSetNew = $true
                                $TempOWASigSetReply = $true
                            }
                            if (($null -ne $TempOWASigFile) -and ($TempOWASigFile -ne '')) {
                                try {
                                    if (Test-Path -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.htm'))) -PathType Leaf) {
                                        if ($EmbedImagesInHtml -eq $false) {
                                            $x = (New-Guid).guid.tostring()
                                            ConvertToSingleFileHTML ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.htm'))) (Join-Path -Path $script:tempDir -ChildPath $x)
                                            $hsHtmlSignature = (Get-Content -LiteralPath (Join-Path -Path $script:tempDir -ChildPath $x) -Encoding UTF8 -Raw).ToString()
                                            Remove-Item (Join-Path -Path $script:tempDir -ChildPath $x) -Force
                                        } else {
                                            $hsHtmlSignature = (Get-Content -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.htm'))) -Encoding UTF8 -Raw).ToString()
                                        }
                                    } else {
                                        $hsHtmlSignature = ''
                                        Write-Host "      Signature file '$($TempOWASigFile + '.htm')' not found. Outlook Web HTML signature will be blank." -ForegroundColor Yellow
                                    }
                                    if (Test-Path -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.txt'))) -PathType Leaf) {
                                        $stTextSig = (Get-Content -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.txt'))) -Encoding UTF8 -Raw).ToString()
                                    } else {
                                        $stTextSig = ''
                                        Write-Host "      Signature file '$($TempOWASigFile + '.txt')' not found. Outlook Web text signature will be blank." -ForegroundColor Yellow
                                    }
                                    $OutlookWebHash = @{}
                                    # Keys are case sensitive when setting them
                                    $OutlookWebHash.Add('signaturehtml', $hsHtmlSignature)
                                    $OutlookWebHash.Add('signaturetext', $stTextSig)
                                    $OutlookWebHash.Add('signaturetextonmobile', $stTextSig)
                                    $OutlookWebHash.Add('autoaddsignature', $TempOWASigSetNew)
                                    $OutlookWebHash.Add('autoaddsignatureonmobile', $TempOWASigSetNew)
                                    $OutlookWebHash.Add('autoaddsignatureonreply', $TempOWASigSetReply)

                                    #Specify the Root folder where the FAI Item is
                                    $folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $($PrimaryMailboxAddress))
                                    $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($exchService, 'OWA.UserOptions', $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)

                                    foreach ($OutlookWebHashKey in $OutlookWebHash.Keys) {
                                        if ($UsrConfig.Dictionary.ContainsKey($OutlookWebHashKey)) {
                                            $UsrConfig.Dictionary[$OutlookWebHashKey] = $OutlookWebHash.$OutlookWebHashKey
                                        } else {
                                            $UsrConfig.Dictionary.Add($OutlookWebHashKey, $OutlookWebHash.$OutlookWebHashKey)
                                        }
                                    }
                                    $UsrConfig.Update() | Out-Null
                                } catch {
                                    Write-Host '    Error setting Outlook Web signature' -ForegroundColor Red
                                    $error[0]
                                }
                            }
                        }
                    }
                }
            }

            if ($SetCurrentUserOOFMessage) {
                Write-Host "  Process Out of Office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                $OOFInternalGUID = (New-Guid).guid
                $OOFExternalGUID = (New-Guid).guid
                $OOFDisabled = $null

                if ($SimulateUser) {
                    Write-Host '    Simulation mode enabled, process OOF templates without changing OOF settings' -ForegroundColor Yellow
                } else {
                    if (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648)) {
                        $OOFSettings = $exchService.GetUserOOFSettings($PrimaryMailboxAddress)
                        if ($($PSVersionTable.PSEdition) -ieq 'Core') { $OOFSettings = $OOFSettings.result }
                        if ($OOFSettings.STATE -eq [Microsoft.Exchange.WebServices.Data.OOFState]::Disabled) { $OOFDisabled = $true }
                    } else {
                        $OOFSettings = $ADPropsCurrentUser.mailboxsettings.automaticRepliesSetting
                        if ($OOFSettings.status -ieq 'disabled') { $OOFDisabled = $true }
                    }
                }

                if (($OOFDisabled -and (-not $SimulateUser)) -or ($SimulateUser)) {
                    EvaluateAndSetSignatures -ProcessOOF:$true

                    if (-not $SimulateUser) {
                        Write-Host "    Set Out of Office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    } else {
                        Write-Host "    Copy Out of Office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    }
                    if (-not $SimulateUser) {
                        if (Test-Path -LiteralPath (Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) {
                            if (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648)) {
                                $OOFSettings.InternalReply = New-Object Microsoft.Exchange.WebServices.Data.OOFReply((Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) -Encoding UTF8 -Raw).tostring())
                            } else {
                                $x = GraphPatchUserMailboxsettings -user $PrimaryMailboxAddress -OOFInternal (Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) -Encoding UTF8 -Raw).tostring()
                                if ($x.error -ne $false) {
                                    Write-Host "      Error setting Outlook Web Out of Office (OOF) auto reply message(s): $($x.error)" -ForegroundColor Red
                                }
                            }
                        }
                        if (Test-Path -LiteralPath (Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) {
                            if (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648)) {
                                $OOFSettings.ExternalReply = New-Object Microsoft.Exchange.WebServices.Data.OOFReply((Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) -Encoding UTF8 -Raw).tostring())
                            } else {
                                $x = GraphPatchUserMailboxsettings -user $PrimaryMailboxAddress -OOFExternal (Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) -Encoding UTF8 -Raw).tostring()
                                if ($x.error -ne $false) {
                                    Write-Host "      Error setting Outlook Web Out of Office (OOF) auto reply message(s): $($x.error)" -ForegroundColor Red
                                }
                            }
                        }
                    } else {
                        foreach ($SignaturePath in $SignaturePaths) {
                            if (Test-Path -LiteralPath (Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) {
                                Copy-Item -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePath) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'OOF Internal.htm')) -Force
                            }
                            if (Test-Path (Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) {
                                Copy-Item -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePath) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'OOF External.htm')) -Force
                            }
                        }
                    }

                    if ((-not $SimulateUser) -and (($null -ne $TrustsToCheckForGroups[0]) -and ($ADPropsCurrentMailbox.msexchrecipienttypedetails -lt 2147483648))) {
                        try {
                            $exchService.SetUserOOFSettings($PrimaryMailboxAddress, $OOFSettings) | Out-Null
                        } catch {
                            Write-Host '      Error setting Outlook Web Out of Office (OOF) auto reply message(s)' -ForegroundColor Red
                        }
                    }
                } else {
                    Write-Host '    Out of Office (OOF) auto reply currently active or scheduled, do not change settings'
                }

                # Delete temporary OOF files from file system
                foreach ($FileName in ("$OOFInternalGUID OOFInternal", "$OOFExternalGUID OOFExternal")) {
                    Remove-Item ((Join-Path -Path $script:tempDir -ChildPath ($FileName + '.*'))) -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }


    # Delete old signatures created by this script, which are no longer available in $SignatureTemplatePath
    # We check all local signatures for a specific marker in HTML code, so we don't touch user created signatures
    if ($DeleteScriptCreatedSignaturesWithoutTemplate -eq $true) {
        Write-Host
        Write-Host "Remove old signatures created by this script, which are no longer centrally available @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        foreach ($SignaturePath in $SignaturePaths) {
            foreach ($file in @(Get-ChildItem -LiteralPath $SignaturePath -Filter '*.htm' -File)) {
                if ((Get-Content -LiteralPath $file.fullname -Encoding UTF8 -Raw) -like ('*' + $HTMLMarkerTag + '*')) {
                    if ($file.name -notin $script:SignatureFilesDone) {
                        Write-Host ("  '" + $([System.IO.Path]::ChangeExtension($file.fullname, '')) + "*'")
                        Remove-Item -LiteralPath $file.fullname -Force -ErrorAction silentlycontinue
                        Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($file.fullname, '.rtf'))) -Force -ErrorAction silentlycontinue
                        Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($file.fullname, '.txt'))) -Force -ErrorAction silentlycontinue
                        Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($file.fullname, '.files'))) -Recurse -Force -ErrorAction silentlycontinue
                    }
                }
            }
        }
    }

    # Delete user created signatures if $DeleteUserCreatedSignatures -eq $true
    if ($DeleteUserCreatedSignatures -eq $true) {
        Write-Host
        Write-Host "Remove user created signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        foreach ($SignaturePath in $SignaturePaths) {
            foreach ($file in @(Get-ChildItem -LiteralPath $SignaturePath -Filter '*.htm' -File)) {
                if ((Get-Content -LiteralPath $file.fullname -Encoding UTF8 -Raw) -notlike ('*' + $HTMLMarkerTag + '*')) {
                    Write-Host ("  '" + $([System.IO.Path]::ChangeExtension($file.fullname, '')) + "*'")
                    Remove-Item -LiteralPath $file.fullname -Force -ErrorAction silentlycontinue
                    Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($file.fullname, '.rtf'))) -Force -ErrorAction silentlycontinue
                    Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($file.fullname, '.txt'))) -Force -ErrorAction silentlycontinue
                    foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                        Remove-Item -LiteralPath ($([System.IO.Path]::GetFileNameWithoutExtension($file.fullname)) + $ConnectedFilesFolderName) -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }


    # Copy signatures to additional path if $AdditionalSignaturePath is set
    if ($AdditionalSignaturePath) {
        Write-Host
        Write-Host "Copy signatures to AdditionalSignaturePath @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        Write-Host "  '$AdditionalSignaturePath'"
        if ($SimulateUser) {
            Write-Host '    Simulation mode enabled, AdditionalSignaturePath already used as output directory' -ForegroundColor Yellow
        } else {
            if (-not (Test-Path $AdditionalSignaturePath -PathType Container -ErrorAction SilentlyContinue)) {
                New-Item -Path $AdditionalSignaturePath -ItemType Directory -Force | Out-Null
                if (-not (Test-Path $AdditionalSignaturePath -PathType Container -ErrorAction SilentlyContinue)) {
                    Write-Host '  Path could not be accessed or created. Ignore path.' -ForegroundColor Yellow
                } else {
                    Copy-Item -Path (Join-Path -Path $SignaturePaths[0] -ChildPath '*') -Destination $AdditionalSignaturePath -Recurse -Force -ErrorAction SilentlyContinue
                }
            } else {
                (Get-ChildItem -Path $AdditionalSignaturePath -Recurse -Force).fullname | Remove-Item -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                Copy-Item -Path (Join-Path -Path $SignaturePaths[0] -ChildPath '*') -Destination $AdditionalSignaturePath -Recurse -Force
            }
        }
    }

    if ($script:CurrentUserDummyMailbox -eq $true) {
        Remove-Item $SignaturePaths[0] -Recurse -Force
    }
}


function GetBitness {
    [CmdletBinding(ConfirmImpact = 'none')]
    param (
        [Parameter(HelpMessage = 'Enter binary file(s) to examine', Position = 0,
            Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({ Test-Path -Path ((Get-Item -Path $_).FullName) })]
        [string[]]  $Path,

        [Alias('PassThru')]
        [switch] $PassThrough
    )

    begin {
        $paths = Resolve-Path -Path $path | Select-Object -ExpandProperty Path
        try {
            $enumString = @'
                public enum BinaryType
                {
                    BIT32 = 0, // A 32-bit Windows-based application, SCS_32BIT_BINARY
                    DOS = 1, // An MS-DOS - based application, SCS_DOS_BINARY
                    WOW = 2, // A 16-bit Windows-based application, SCS_WOW_BINARY
                    PIF = 3, // A PIF file that executes an MS-DOS based application, SCS_PIF_BINARY
                    POSIX = 4, // A POSIX based application, SCS_POSIX_BINARY
                    OS216 = 5, // A 16-bit OS/2-based application, SCS_OS216_BINARY
                    BIT64 = 6 // A 64-bit Windows-based application, SCS_64BIT_BINARY
                }
'@

            Add-Type -TypeDefinition $enumString
        } catch {
        }

        try {
            $Signature = @'
                    [DllImport("kernel32.dll")]
                    public static extern bool GetBinaryType(
                                        string lpApplicationName,
                                        ref int lpBinaryType
                    );
'@

            Add-Type -MemberDefinition $Signature -Name BinaryType -Namespace PFWin32Utils
        } catch {
        }
    }

    process {
        foreach ($Item in $Paths) {
            $ReturnedType = -1
            $Result = [PFWin32Utils.BinaryType]::GetBinaryType($Item, [ref] $ReturnedType)

            if (!$Result -or ($ReturnedType -eq -1)) {
                Write-Error -Message "Failed to get binary type for file $($Item)"
            } else {
                $ToReturn = [BinaryType] $ReturnedType
                if ($PassThrough) {
                    Get-Item -Path $Item.FullName -Force |
                    Add-Member -MemberType noteproperty -Name BinaryType -Value $ToReturn -Force -PassThru
                } else {
                    $ToReturn
                }
            }
        }
    }
}


Function ConvertToSingleFileHTML([string]$inputfile, [string]$outputfile) {
    $tempFileContent = Get-Content -LiteralPath $inputfile -Encoding UTF8 -Raw

    $src = @()
    foreach ($regex in @(([regex]'(?i)src="(.*?)"').Matches($tempFileContent))) {
        $src += $regex.Groups[0].Value
        if ($regex.Groups[0].Value.StartsWith('src="data:')) {
            $src += ''
        } else {
            $src += (Join-Path -Path (Split-Path -Path ($inputfile) -Parent) -ChildPath ([uri]::UnEscapeDataString($regex.Groups[1].Value)))
        }
    }
    for ($x = 0; $x -lt $src.count; $x = $x + 2) {
        if ($src[$x].StartsWith('src="data:')) {
        } elseif (Test-Path -LiteralPath $src[$x + 1] -PathType leaf) {
            $fmt = $null
            switch ((Get-ChildItem -LiteralPath $src[$x + 1]).Extension) {
                '.apng' { $fmt = 'data:image/apng;base64,' }
                '.avif' { $fmt = 'data:image/avif;base64,' }
                '.gif' { $fmt = 'data:image/gif;base64,' }
                '.jpg' { $fmt = 'data:image/jpeg;base64,' }
                '.jpeg' { $fmt = 'data:image/jpeg;base64,' }
                '.jfif' { $fmt = 'data:image/jpeg;base64,' }
                '.pjpeg' { $fmt = 'data:image/jpeg;base64,' }
                '.pjp' { $fmt = 'data:image/jpeg;base64,' }
                '.png' { $fmt = 'data:image/png;base64,' }
                '.svg' { $fmt = 'data:image/svg+xml;base64,' }
                '.webp' { $fmt = 'data:image/webp;base64,' }
                '.css' { $fmt = 'data:text/css;base64,' }
                '.less' { $fmt = 'data:text/css;base64,' }
                '.js' { $fmt = 'data:text/javascript;base64,' }
                '.otf' { $fmt = 'data:font/otf;base64,' }
                '.sfnt' { $fmt = 'data:font/sfnt;base64,' }
                '.ttf' { $fmt = 'data:font/ttf;base64,' }
                '.woff' { $fmt = 'data:font/woff;base64,' }
                '.woff2' { $fmt = 'data:font/woff2;base64,' }
            }
            if ($fmt) {
                if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                    $tempFileContent = $tempFileContent.replace($src[$x], ('src="' + $fmt + [Convert]::ToBase64String((Get-Content -LiteralPath $src[$x + 1] -AsByteStream)) + '"'))
                } else {
                    $tempFileContent = $tempFileContent.replace($src[$x], ('src="' + $fmt + [Convert]::ToBase64String((Get-Content -LiteralPath $src[$x + 1] -Encoding Byte)) + '"'))
                }
            }
        }
    }

    [System.IO.File]::WriteAllLines($outputfile, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))
}


function EvaluateAndSetSignatures {
    Param(
        [switch]$ProcessOOF = $false
    )

    if ($ProcessOOF -eq $true) {
        $SigOrOOF = 'OOF'
        $Indent = '  '
    } else {
        $SigOrOOF = 'Signature'
        $Indent = ''
    }

    foreach ($TemplateGroup in ('common', 'group', 'mailbox')) {
        Write-Host "$Indent  Process $TemplateGroup $(if($TemplateGroup -iin ('group', 'mailbox')){'specific '})templates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

        if (-not (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly -ErrorAction SilentlyContinue)) {
            continue
        }

        for ($TemplateFileIndex = 0; $TemplateFileIndex -lt (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly).count; $TemplateFileIndex++) {
            $TemplateFile = (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly)[$TemplateFileIndex]
            $TemplateIniSettingsIndex = $TemplateFile.TemplateIniSettingsIndex

            if (-not $TemplateIniSettingsIndex) {
                continue
            }

            if (-not (Get-Variable -Name "$($SigOrOOF)Files$($TemplateGroup)" -ValueOnly).containskey($TemplateIniSettingsIndex)) {
                continue
            } else {
                $Template = (Get-Variable -Name "$($SigOrOOF)Files$($TemplateGroup)" -ValueOnly)[$TemplateIniSettingsIndex].GetEnumerator() | Select-Object -First 1
            }

            Write-Host "$Indent    '$([System.IO.Path]::GetFileName($Template.key))' (ini index #$($TemplateIniSettingsIndex)) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            Write-Host "$Indent      Check permissions"
            $TemplateAllowed = $false


            # check for allow entries
            Write-Host "$Indent        Allows"
            if ($TemplateGroup -ieq 'common') {
                $TemplateAllowed = $true
                Write-Host "$Indent          Common: Template is classified as common template valid for all mailboxes"
            } elseif ($TemplateGroup -ieq 'group') {
                $tempAllowCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($GroupsSid)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        $tempSearchSting = $GroupsSid
                        Write-Host "$Indent          First group match: $(($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $tempSearchSting }).name -join '/') = $($GroupsSid)"
                        break
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox is not member of any allowed group"
                }
            } elseif ($TemplateGroup -ieq 'mailbox') {
                $tempAllowCount = 0

                foreach ($CurrentMailboxSmtpAddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        Write-Host "$Indent          First e-mail address match: $($CurrentMailboxSmtpAddress)"
                        break
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          E-mail address: Mailbox does not have any allowed e-mail address"
                }
            }


            # check for deny entries
            if ($TemplateAllowed -eq $true) {
                Write-Host "$Indent        Denies"
                # check for group deny
                $tempDenyCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($GroupsSid)``]*") {
                        $TemplateAllowed = $false
                        $tempDenyCount++
                        $tempSearchSting = $($GroupsSid)
                        Write-Host "$Indent          First group match: $((($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $tempSearchSting }).name -replace '^\[', '[-:') -join '/') = $($GroupsSid)"
                        break
                    }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox is not member of any denied group"
                }

                # check for mail address deny
                $tempDenyCount = 0

                foreach ($CurrentMailboxSmtpAddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $false
                        $tempDenyCount++
                        Write-Host "$Indent          First e-mail address match: $($CurrentMailboxSmtpAddress)"
                        break
                    }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          E-Mail address: Mailbox does not have any denied e-mail address"
                }
            }

            # result
            if ($Template -and ($TemplateAllowed -eq $true)) {
                Write-Host "$Indent        Use template as there is at least one allow and no deny for this mailbox"
                if ($ProcessOOF) {
                    if ($OOFFilesInternal.contains($TemplateIniSettingsIndex)) {
                        $OOFInternal = $Template
                    }

                    if ($OOFFilesExternal.contains($TemplateIniSettingsIndex)) {
                        $OOFExternal = $Template
                    }
                } else {
                    $Signature = $Template
                    SetSignatures -ProcessOOF:$ProcessOOF
                }
            } else {
                Write-Host "$Indent        Do not use template as there is no allow or at least one deny for this mailbox"
            }
        }
    }

    if ($ProcessOOF) {
        # Internal OOF message
        if ($OOFInternal -or $OOFExternal) {
            Write-Host "$Indent  Convert final OOF templates to HTM format @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        }

        if ($OOFInternal) {
            $Signature = $OOFInternal

            if ($OOFExternal -eq $OOFInternal) {
                Write-Host "$Indent    Common OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            } else {
                Write-Host "$Indent    Internal OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            }
            if ($UseHtmTemplates) {
                $Signature.value = "$OOFInternalGUID OOFInternal.htm"
            } else {
                $Signature.value = "$OOFInternalGUID OOFInternal.docx"
            }

            SetSignatures -ProcessOOF:$ProcessOOF

            if ($OOFExternal -eq $OOFInternal) {
                Copy-Item -Path (Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm") -Destination (Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")
            }
        }
    }

    # External OOF message
    if ($OOFExternal -and ($OOFExternal -ne $OOFInternal)) {
        $Signature = $OOFExternal

        Write-Host "$Indent    External OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

        if ($UseHtmTemplates) {
            $Signature.value = "$OOFExternalGUID OOFExternal.htm"
        } else {
            $Signature.value = "$OOFExternalGUID OOFExternal.docx"
        }

        SetSignatures -ProcessOOF:$ProcessOOF
    }
}


function SetSignatures {
    Param(
        [switch]$ProcessOOF = $false
    )

    if ($ProcessOOF) {
        $Indent = '  '
    }

    if (-not $ProcessOOF) {
        Write-Host "      Outlook signature name: '$([System.IO.Path]::ChangeExtension($($Signature.value), $null) -replace '\.$')'"
    }

    if (-not $ProcessOOF) {
        $SignatureFileAlreadyDone = ($script:SignatureFilesDone -contains $TemplateIniSettingsIndex)

        if ($SignatureFileAlreadyDone) {
            Write-Host "$Indent      Template already processed before with higher priority, no need to update signature"
        } else {
            $script:SignatureFilesDone += $TemplateIniSettingsIndex
        }
    }
    if (($SignatureFileAlreadyDone -eq $false) -or $ProcessOOF) {
        Write-Host "$Indent      Create temporary file copy"

        $pathGUID = (New-Guid).guid.tostring()
        $path = Join-Path -Path $script:tempDir -ChildPath "$($pathGUID).htm"
        $pathConnectedFolderNames = @()
        foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
            $pathConnectedFolderNames += "$($pathGUID)$($ConnectedFilesFolderName)"
        }

        if ($UseHtmTemplates) {
            # use .html for temporary file, .htm for final file
            try {
                if ($EmbedImagesInHtml -eq $false) {
                    Copy-Item -LiteralPath $Signature.name -Destination $path
                    foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                        if (Test-Path (Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName")) {
                            Copy-Item (Join-Path -Path (Split-Path $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName") (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files") -Recurse -Force
                            return
                        }
                    }
                } else {
                    ConvertToSingleFileHTML $Signature.Name $path
                }
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                Write-Host $error[0]
                continue
            }
        } else {
            $path = $([System.IO.Path]::ChangeExtension($($path), '.docx'))
            try {
                Copy-Item -LiteralPath $Signature.Name -Destination $path -Force
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                continue
            }
        }

        $Signature.value = $([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))
        if (-not $ProcessOOF) {
            $script:SignatureFilesDone += $Signature.Value
        }

        if ($UseHtmTemplates) {
            Write-Host "$Indent      Replace picture variables"
            $html = New-Object -ComObject 'HTMLFile'
            $HTML.IHTMLDocument2_write((Get-Content -LiteralPath $path -Encoding UTF8 -Raw))

            foreach ($image in @($html.images)) {
                foreach ($VariableName in (('$CURRENTMAILBOXMANAGERPHOTO$', $CURRENTMAILBOXMANAGERPHOTOGUID) , ('$CURRENTMAILBOXPHOTO$', $CURRENTMAILBOXPHOTOGUID), ('$CURRENTUSERMANAGERPHOTO$', $CURRENTUSERMANAGERPHOTOGUID), ('$CURRENTUSERPHOTO$', $CURRENTUSERPHOTOGUID))) {
                    if (($image.src -clike "*$($VariableName[0])*") -or ($image.alt -clike "*$($VariableName[0])*")) {
                        if ($null -ne $ReplaceHash[$VariableName[0]]) {
                            if ($EmbedImagesInHtml -eq $false) {
                                Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -replace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                Copy-Item (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$($VariableName[0]).jpeg") -Force
                                $image.src = [System.Web.HttpUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($VariableName[0]).jpeg")
                                if ($image.alt) {
                                    $image.alt = $($image.alt).replace($VariableName[0], '')
                                }
                            } else {
                                $image.src = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                if ($image.alt) {
                                    $image.alt = $($image.alt).replace($VariableName[0], '')
                                }
                            }
                        } else {
                            $image.src = "$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -replace '^about:', '')))))"
                        }
                    } elseif (($image.src -clike "*$(($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$')*") -or ($image.alt -clike "*$(($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$')*")) {
                        if ($null -ne $ReplaceHash[$VariableName[0]]) {
                            if ($EmbedImagesInHtml -eq $false) {
                                Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -replace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                                Copy-Item (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$($VariableName[0]).jpeg") -Force
                                $image.src = [System.Web.HttpUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($VariableName[0]).jpeg")
                                if ($image.alt) {
                                    $image.alt = $($image.alt).replace((($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$'), '')
                                }
                            } else {
                                $image.src = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                if ($image.alt) {
                                    $image.alt = $($image.alt).replace((($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$'), '')
                                }
                            }
                        } else {
                            Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Web.HttpUtility]::UrlDecode(($image.src -replace '^about:', '')))))") -Force -ErrorAction SilentlyContinue
                            $image.removenode() | Out-Null
                        }
                    }
                }
            }

            Write-Host "$Indent      Replace non-picture variables"
            $tempFileContent = $html.documentelement.outerhtml
            foreach ($replaceKey in $replaceHash.Keys) {
                if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                    $tempFileContent = $tempFileContent.replace($replacekey, $replaceHash.$replaceKey)
                }
            }

            if (-not $ProcessOOF) {
                $tempFileContent | Out-File -LiteralPath $path -Encoding UTF8 -Force
            } else {
                $tempFileContent | Out-File -LiteralPath (Join-Path -Path $script:tempDir -ChildPath $Signature.Value) -Encoding UTF8 -Force
            }
        }

        if ($CreateRtfSignatures -or $CreateTxtSignatures) {
            $script:COMWord.Documents.Open($path, $false) | Out-Null
        }

        if (-not $UseHtmTemplates) {
            Write-Host "$Indent      Replace picture variables"
            foreach ($image in @($script:COMWord.ActiveDocument.Shapes + $script:COMWord.ActiveDocument.InlineShapes)) {
                try {
                    if ($image.linkformat.sourcefullname) {
                        foreach ($Variablename in (('$CURRENTMAILBOXMANAGERPHOTO$', $CURRENTMAILBOXMANAGERPHOTOGUID) , ('$CURRENTMAILBOXPHOTO$', $CURRENTMAILBOXPHOTOGUID), ('$CURRENTUSERMANAGERPHOTO$', $CURRENTUSERMANAGERPHOTOGUID), ('$CURRENTUSERPHOTO$', $CURRENTUSERPHOTOGUID))) {
                            if (([System.IO.Path]::GetFileName($image.linkformat.sourcefullname).contains($Variablename[0])) -or $(if ($image.alternativetext) { (($image.alternativetext).contains($Variablename[0])) })) {
                                if ($null -ne $ReplaceHash[$Variablename[0]]) {
                                    $image.linkformat.sourcefullname = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                    if ($image.alternativetext) {
                                        $image.alternativetext = $($image.alternativetext).replace($Variablename[0], '')
                                    }
                                }
                            } elseif (([System.IO.Path]::GetFileName($image.linkformat.sourcefullname).contains(($Variablename[0][-999..-2] -join '') + 'DELETEEMPTY$')) -or $(if ($image.alternativetext) { ($image.alternativetext.contains(($Variablename[0][-999..-2] -join '') + 'DELETEEMPTY$')) })) {
                                if ($null -ne $ReplaceHash[$Variablename[0]]) {
                                    $image.linkformat.sourcefullname = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                    if ($image.alternativetext) {
                                        $image.alternativetext = $($image.alternativetext).replace((($Variablename[0][-999..-2] -join '') + 'DELETEEMPTY$'), '')
                                    }
                                } else {
                                    $image.delete()
                                }
                            }
                        }
                    }
                } catch {
                }

                # Setting the values in word is very slow, so we use temporay variables
                $tempImageAlternativeText = $image.alternativetext
                $tempImageHyperlinkAddress = $image.hyperlink.Address
                $tempImageHyperlinkSubAddress = $image.hyperlink.SubAddress
                $tempImageHyperlinkEmailSubject = $image.hyperlink.EmailSubject
                $tempImageHyperlinkScreenTip = $image.hyperlink.ScreenTip

                foreach ($replaceKey in $replaceHash.Keys) {
                    if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                        if ($null -ne $tempimagealternativetext) {
                            $tempimagealternativetext = $tempimagealternativetext.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkAddress) {
                            $tempimagehyperlinkAddress = $tempimagehyperlinkAddress.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkSubAddress) {
                            $tempimagehyperlinkSubAddress = $tempimagehyperlinkSubAddress.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkEmailSubject) {
                            $tempimagehyperlinkEmailSubject = $tempimagehyperlinkEmailSubject.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkScreenTip) {
                            $tempimagehyperlinkScreenTip = $tempimagehyperlinkScreenTip.replace($replaceKey, $replaceHash.replaceKey)
                        }
                    }
                }

                if ($null -ne $tempimagealternativetext) {
                    $image.alternativetext = $tempImageAlternativeText
                }
                if ($null -ne $tempimagehyperlinkAddress) {
                    $image.hyperlink.Address = $tempImageHyperlinkAddress
                }
                if ($null -ne $tempimagehyperlinkSubAddress) {
                    $image.hyperlink.SubAddress = $tempImageHyperlinkSubAddress
                }
                if ($null -ne $tempimagehyperlinkEmailSubject) {
                    $image.hyperlink.EmailSubject = $tempImageHyperlinkEmailSubject
                }
                if ($null -ne $tempimagehyperlinkScreenTip) {
                    $image.hyperlink.ScreenTip = $tempImageHyperlinkScreenTip
                }
            }

            Write-Host "$Indent      Replace non-picture variables"
            $wdFindContinue = 1
            $MatchCase = $true
            $MatchWholeWord = $true
            $MatchWildcards = $False
            $MatchSoundsLike = $False
            $MatchAllWordForms = $False
            $Forward = $True
            $Wrap = $wdFindContinue
            $Format = $False
            $wdFindContinue = 1
            $ReplaceAll = 2

            # Replace in current view (show or hide field codes)
            foreach ($replaceKey in $replaceHash.Keys) {
                if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                    $FindText = $replaceKey
                    $ReplaceWith = (($replaceHash.$replaceKey -replace "`r`n", '^p') -replace "`n", '^l')
                    $script:COMWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, `
                            $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                            $Wrap, $Format, $ReplaceWith, $ReplaceAll) | Out-Null
                }
            }

            # Invert current view (show or hide field codes)
            # This is neccessary to be able to replace variables in hyperlinks and quicktips of hyperlinks
            $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = (-not $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes)
            foreach ($replaceKey in $replaceHash.Keys) {
                if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                    $FindText = $replaceKey
                    $ReplaceWith = (($replaceHash.$replaceKey -replace "`r`n", '^p') -replace "`n", '^l')
                    $script:COMWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, `
                            $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                            $Wrap, $Format, $ReplaceWith, $ReplaceAll) | Out-Null
                }
            }

            # Restore original view
            $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = (-not $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes)

            # Exports
            Write-Host "$Indent      Export to HTM format"
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatFilteredHTML')
            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
            $script:COMWord.ActiveDocument.Weboptions.encoding = 65001
            # Overcome Word security warning when export contains embedded pictures
            if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value 0 -ErrorAction Ignore | Out-Null
            }
            $WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
            if (($null -eq $WordDisableWarningOnIncludeFieldsUpdate) -or ($WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -PropertyType DWord -Value 1 -ErrorAction Ignore | Out-Null
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value 1 -ErrorAction Ignore | Out-Null
            }
            try {
                $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
            } catch {
                Start-Sleep -Seconds 2
                $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
            }
            # Restore original security setting
            if ($null -eq $WordDisableWarningOnIncludeFieldsUpdate) {
                Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
            } else {
                Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            }
        }

        if (-not $ProcessOOF) {
            if ($CreateRtfSignatures -eq $true) {
                Write-Host "$Indent      Export to RTF format"
                $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatRTF')
                $path = $([System.IO.Path]::ChangeExtension($path, '.rtf'))
                # Overcome Word security warning when export contains embedded pictures
                if ((Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security\DisableWarningOnIncludeFieldsUpdate") -eq $false) {
                    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value 0 -ErrorAction Ignore | Out-Null
                }
                $WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                if (($null -eq $WordDisableWarningOnIncludeFieldsUpdate) -or ($WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -PropertyType DWord -Value 1 -ErrorAction Ignore | Out-Null
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value 1 -ErrorAction Ignore | Out-Null
                }
                try {
                    $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
                } catch {
                    Start-Sleep -Seconds 2
                    $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
                }
                # Restore original security setting
                if ($null -eq $WordDisableWarningOnIncludeFieldsUpdate) {
                    Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                } else {
                    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }

                $script:COMWord.ActiveDocument.Close($false)

                # RTF files with embedded images get really huge
                # See https://support.microsoft.com/kb/224663 for a system-wide workaround
                # The following workaround is from https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_mac-mso_mac2011/huge-rtf-files-solved-on-windows-but-searching-for/58e54b37-cfd0-4a07-ac62-1cfc2769cad5
                $openFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdOpenFormat], 'wdOpenFormatUnicodeText')
                $script:COMWord.Documents.Open($path, $false, $false, $false, '', '', $true, '', '', $openFormat) | Out-Null
                $FindText = '\{\\nonshppict*\}\}'
                $ReplaceWith = ''
                $script:COMWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, `
                        $true, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                        $Wrap, $Format, $ReplaceWith, $ReplaceAll) | Out-Null
                try {
                    $script:COMWord.ActiveDocument.Save()
                } catch {
                    Start-Sleep -Seconds 2
                    $script:COMWord.ActiveDocument.Save()
                }
            }

            if ($CreateRtfSignatures -or $CreateTxtSignatures) {
                $script:COMWord.ActiveDocument.Close($false)
            }

            if ($CreateTxtSignatures -eq $true) {
                Write-Host "$Indent      Export to TXT format"
                # We work with the .htm file to avoid problems with empty lines at the end of exported .txt files. Details: https://eileenslounge.com/viewtopic.php?t=16703
                $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
                $script:COMWord.Documents.Open($path, $false) | Out-Null
                $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatUnicodeText')
                $script:COMWord.ActiveDocument.TextEncoding = 1200
                $path = $([System.IO.Path]::ChangeExtension($path, '.txt'))
                try {
                    $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
                } catch {
                    Start-Sleep -Seconds 2
                    $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
                }
                $script:COMWord.ActiveDocument.Close($false)
            }
        } else {
            if ($CreateRtfSignatures -or $CreateTxtSignatures) {
                $script:COMWord.ActiveDocument.Close($false)
            }
        }

        Write-Host "$Indent      Embed local files in HTM format and add marker"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

        $tempFileContent = Get-Content -LiteralPath $path -Encoding UTF8 -Raw

        if ($tempFileContent -notlike "*$HTMLMarkerTag*") {
            if ($tempFileContent -like '*<head>*') {
                $tempFileContent = $tempFileContent -ireplace ('<HEAD>', ('<HEAD>' + $HTMLMarkerTag))
            } else {
                $tempFileContent = $tempFileContent -ireplace ('<HTML>', ('<HTML><HEAD>' + $HTMLMarkerTag + '</HEAD>'))
            }
        }

        if (-not $ProcessOOF) {
            if ($EmbedImagesInHtml -eq $false) {
                foreach ($pathConnectedFolderName in $pathConnectedFolderNames) {
                    if (Test-Path (Join-Path -Path (Split-Path $path) -ChildPath $($pathConnectedFolderName))) {
                        $tempFileContent = $tempFileContent -replace ('(\s*src=")(' + $pathConnectedFolderName + '\/)'), ('$1' + "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value)).files/")
                        Rename-Item (Join-Path -Path (Split-Path $path) -ChildPath $($pathConnectedFolderName)) $([System.IO.Path]::GetFileNameWithoutExtension($Signature.value) + '.files') -ErrorAction SilentlyContinue
                        return
                    }
                }
                [System.IO.File]::WriteAllText($path, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))
            } else {
                [System.IO.File]::WriteAllText($path, $tempFileContent, (New-Object System.Text.UTF8Encoding($False)))
                ConvertToSingleFileHTML $path $path
            }
        } else {
            ConvertToSingleFileHTML $path ((Join-Path -Path $script:tempDir -ChildPath $Signature.Value))
        }


        if (-not $ProcessOOF) {
            foreach ($SignaturePath in $SignaturePaths) {
                if ($CurrentMailboxUseSignatureRoaming -eq $true) {
                    # Microsoft signature roaming available
                    Write-Host "$Indent      Microsoft signature roaming enabled. What to do now?" -ForegroundColor Red
                } else {
                    # Microsoft signature roaming not available
                    Write-Host "$Indent      Copy signature files to '$SignaturePath'"
                    foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                        Remove-Item -LiteralPath ((Join-Path -Path $SignaturePath -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value))") + $ConnectedFilesFolderName) -Recurse -Force -ErrorAction SilentlyContinue
                    }
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.htm')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))) -Force
                    if ($EmbedImagesInHtml -eq $false) {
                        if (Test-Path (Join-Path -Path (Split-Path $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))")) {
                            Copy-Item -LiteralPath (Join-Path -Path (Split-Path $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))") -Destination $SignaturePath -Force -Recurse
                        }
                    }
                    if ($CreateRtfSignatures -eq $true) {
                        Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.rtf')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))) -Force
                    } else {
                        Remove-Item ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))) -Force -ErrorAction SilentlyContinue
                    }
                    if ($CreateTxtSignatures -eq $true) {
                        Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.txt')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))) -Force
                    } else {
                        Remove-Item ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))) -Force -ErrorAction SilentlyContinue

                    }
                }
            }
        }

        Write-Host "$Indent      Remove temporary files"
        foreach ($extension in ('.docx', '.htm', '.rtf', '.txt')) {
            Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, $extension)) -ErrorAction SilentlyContinue | Out-Null
        }

        Foreach ($file in @(Get-ChildItem -Path ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($path) + '*') -Directory).FullName) {
            Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
        }
        Remove-Item (Join-Path -Path (Split-Path $path) -ChildPath $([System.IO.Path]::ChangeExtension($signature.value, '.files'))) -Force -Recurse -ErrorAction SilentlyContinue
    }

    if ((-not $ProcessOOF)) {
        # Set default signature for new e-mails
        if ($SignatureFilesDefaultNew.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if (-not $SimulateUser) {
                        Write-Host "$Indent      Set signature as default for new messages"
                        if ($script:CurrentUserDummyMailbox -ne $true) {
                            if ($OutlookFileVersion -ge '16.0.0.0') {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -PropertyType String -Value ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + $(if ($CurrentMailboxUseSignatureRoaming -eq $true) { " ($($MailAddresses[$AccountNumberRunning]))" })) -Force | Out-Null
                            } else {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + $(if ($CurrentMailboxUseSignatureRoaming -eq $true) { " ($($MailAddresses[$AccountNumberRunning]))" })) + "`0")))) -Force | Out-Null
                            }
                        } else {
                            $script:CurrentUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        Copy-Item -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + '.htm')) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'Default New.htm')) -Force
                        Copy-Item -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + '.rtf')) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'Default New.rtf')) -Force
                        Copy-Item -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + '.txt')) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'Default New.txt')) -Force
                    }
                }
            }
        }

        # Set default signature for replies and forwarded e-mails
        if ($SignatureFilesDefaultReplyFwd.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if (-not $SimulateUser) {
                        Write-Host "$Indent      Set signature as default for reply/forward messages"
                        if ($script:CurrentUserDummyMailbox -ne $true) {
                            if ($OutlookFileVersion -ge '16.0.0.0') {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType String -Value ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + $(if ($CurrentMailboxUseSignatureRoaming -eq $true) { " ($($MailAddresses[$AccountNumberRunning]))" })) -Force | Out-Null
                            } else {
                                New-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + $(if ($CurrentMailboxUseSignatureRoaming -eq $true) { " ($($MailAddresses[$AccountNumberRunning]))" })) + "`0")))) -Force | Out-Null
                            }
                        } else {
                            $script:CurrentUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        Copy-Item -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + '.htm')) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'Default Reply-Forward.htm')) -Force
                        Copy-Item -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + '.rtf')) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'Default Reply-Forward.rtf')) -Force
                        Copy-Item -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + '.txt')) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($SignaturePaths[0]) -ChildPath "$($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath 'Default Reply-Forward.txt')) -Force
                    }
                }
            }
        }
    }
}


function CheckADConnectivity {
    param (
        [array]$CheckDomains,
        [string]$CheckProtocolText,
        [string]$Indent
    )
    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 10)
    $RunspacePool.Open()

    for ($DomainNumber = 0; $DomainNumber -lt $CheckDomains.count; $DomainNumber++) {
        if ($($CheckDomains[$DomainNumber]) -eq '') {
            continue
        }

        $PowerShell = [powershell]::Create()
        $PowerShell.RunspacePool = $RunspacePool

        [void]$PowerShell.AddScript( {
                Param (
                    [string]$CheckDomain,
                    [string]$CheckProtocolText
                )
                $DebugPreference = 'Continue'
                Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"
                Write-Output "$CheckDomain"
                $Search = New-Object DirectoryServices.DirectorySearcher
                $Search.PageSize = 1000
                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($CheckProtocolText)://$CheckDomain")
                $Search.filter = '(objectclass=user)'
                try {
                    $UserAccount = ([ADSI]"$(($Search.FindOne()).path)")
                    Write-Output 'QueryPassed'
                } catch {
                    Write-Output 'QueryFailed'
                }
            }).AddArgument($($CheckDomains[$DomainNumber])).AddArgument($CheckProtocolText)
        $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
        $Handle = $PowerShell.BeginInvoke($Object, $Object)
        $temp = '' | Select-Object PowerShell, Handle, Object, StartTime, Done
        $temp.PowerShell = $PowerShell
        $temp.Handle = $Handle
        $temp.Object = $Object
        $temp.StartTime = $null
        $temp.Done = $false
        [void]$script:jobs.Add($Temp)
    }
    while (($script:jobs.Done | Where-Object { $_ -eq $false }).count -ne 0) {
        foreach ($job in $script:jobs) {
            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -match 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -replace '[^0-9]'
                $job.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $job.StartTime) {
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $job.Object[0..$(($job.object).count - 1)]
                    Write-Host "$Indent$($data[0])"
                    if ($data -icontains 'QueryPassed') {
                        Write-Host "$Indent  $CheckProtocolText query successful"
                        $returnvalue = $true
                    } else {
                        Write-Host "$Indent  $CheckProtocolText query failed, remove domain from list." -ForegroundColor Red
                        Write-Host "$Indent  If this error is permanent, check firewalls, DNS and AD trust. Consider parameter TrustsToCheckForGroups." -ForegroundColor Red
                        $TrustsToCheckForGroups.remove($data[0])
                        $returnvalue = $false
                    }
                    $job.Done = $true
                }
            }
        }
    }
    return $returnvalue
}


function CheckPath([string]$path, [switch]$silent = $false, [switch]$create = $false) {
    if ($create -eq $false) {
        if (($path.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($path -ilike '*@ssl\*')) {
            $path = $path -ireplace '@ssl\\', '\'
            $path = ([uri]::UnescapeDataString($path) -ireplace ('https://', '\\'))
            $path = ([System.URI]$path).AbsoluteURI -replace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\$2' -replace '/', '\'
            $path = [uri]::UnescapeDataString($path)
        } else {
            try {
                $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
                $path = ([System.URI]$path).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -replace '/', '\'
                $path = [uri]::UnescapeDataString($path)
            } catch {
                if ($silent -eq $false) {
                    Write-Host ': ' -NoNewline
                    Write-Host "Problem connecting to or reading from folder '$path'. Exit." -ForegroundColor Red
                    exit 1
                }
            }
        }

        if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
            # Reconnect already connected network drives at the OS level
            # New-PSDrive is not enough for this
            foreach ($NetworkConnection in @(Get-CimInstance Win32_NetworkConnection)) {
                & net use $NetworkConnection.LocalName $NetworkConnection.RemoteName 2>&1 | Out-Null
            }

            if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
                # Connect network drives
                '`r`n' | & net use "$path" 2>&1 | Out-Null
                try {
                    (Test-Path -LiteralPath $path -ErrorAction Stop) | Out-Null
                } catch {
                    if ($_.CategoryInfo.Category -eq 'PermissionDenied') {
                        & net use "$path" 2>&1
                    }
                }
                & net use "$path" /d 2>&1 | Out-Null
            }

            if (($path -ilike '*@ssl\*') -and (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue))) {
                Try {
                    # Add site to trusted sites in internet options
                    New-Item ('HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\' + (New-Object System.Uri -ArgumentList ($path -ireplace ('@SSL', ''))).Host) -Force | New-ItemProperty -Name * -Value 1 -Type DWORD -Force | Out-Null

                    # Open site in new IE process
                    $oIE = New-Object -com InternetExplorer.Application
                    $oIE.Visible = $false
                    $oIE.Navigate2('https://' + ((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))
                    $oIE = $null

                    # Wait until an IE tab with the corresponding URL is open
                    $app = New-Object -com shell.application
                    $i = 0
                    while ($i -lt 1) {
                        $i += @($app.windows() | Where-Object { $_.LocationURL -like ('*' + ([uri]::EscapeUriString(((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') }).count
                        Start-Sleep -Milliseconds 50
                    }

                    # Wait until the corresponding URL is fully loaded, then close the tab
                    foreach ($window in @($app.windows() | Where-Object { $_.LocationURL -like ('*' + ([uri]::EscapeUriString(((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') })) {
                        while ($window.busy) {
                            Start-Sleep -Milliseconds 50
                        }
                        $window.quit([ref]$false)
                    }

                    $app = $null
                } catch {
                }
            }
        }

        if ((Test-Path -LiteralPath $path) -eq $false) {
            if ($silent -eq $false) {
                Write-Host ': ' -NoNewline
                Write-Host "Problem connecting to or reading from folder '$path'. Exit." -ForegroundColor Red
                exit 1
            } else {
                return $false
            }
        } else {
            if ($silent -eq $false) {
                Write-Host
            } else {
                return $true
            }
        }
    } else {
        if ($path.StartsWith('https://', 'CurrentCultureIgnoreCase')) {
            $path = ((([uri]::UnescapeDataString($path) -ireplace ('https://', '\\')) -replace ('(.*?)/(.*)', '${1}@SSL\$2')) -replace ('/', '\'))
        }
        $pathTemp = $path
        for ($i = (($path.ToCharArray() | Where-Object { $_ -eq '\' } | Measure-Object).Count); $i -ge 0; $i--) {
            if ((CheckPath $pathTemp -Silent) -eq $true) {
                if (-not (Test-Path $pathTemp -PathType Container -ErrorAction SilentlyContinue)) {
                    Write-Host ': ' -NoNewline
                    Write-Host "'$pathTemp' is a file, '$path' not valid. Exit." -ForegroundColor Red
                    exit 1
                }
                if ($pathTemp -eq $path) {
                    break
                } else {
                    New-Item -ItemType Directory -Path $path -ErrorAction SilentlyContinue | Out-Null
                    if (Test-Path -Path $path -PathType Container) {
                        break
                    }
                }
            } else {
                $pathTemp = Split-Path ($pathTemp -ireplace '@SSL', '') -Parent
            }
        }
        if ((checkpath $path -silent) -ne $true) {
            Write-Host ': ' -NoNewline
            Write-Host "Problem connecting to or reading from folder '$path'. Exit." -ForegroundColor Red
            exit 1
        } else {
            Write-Host
        }
    }
}


function GraphGetToken {
    if ($GraphCredentialFile) {
        try {
            $auth = Import-Clixml -Path $GraphCredentialFile
            $script:authorizationHeader = @{
                Authorization = $auth.authHeader
            }
            return @{
                error          = $false
                accessToken    = $auth.AccessToken
                accessTokenExo = $auth.AccessTokenExo
                authHeader     = $auth.authHeader
            }
        } catch {
            return @{
                error       = ($error | Out-String)
                accessToken = $null
                authHeader  = $null
            }
        }
    } else {
        $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -TenantId $(if ($null -ne $script:CurrentUser) { ($script:CurrentUser -split '@')[1] } else { 'organizations' }) -RedirectUri 'http://localhost' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

        try {
            $auth = $script:msalClientApp | Get-MsalToken -LoginHint $(if ($null -ne $script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes 'https://graph.microsoft.com/openid', 'https://graph.microsoft.com/email', 'https://graph.microsoft.com/profile', 'https://graph.microsoft.com/user.read.all', 'https://graph.microsoft.com/group.read.all', 'https://graph.microsoft.com/mailboxsettings.readwrite', 'https://graph.microsoft.com/EWS.AccessAsUser.All' -IntegratedWindowsAuth
        } catch {
            try {
                $auth = $script:msalClientApp | Get-MsalToken -LoginHint $(if ($null -ne $script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes ('https://graph.microsoft.com/openid', 'https://graph.microsoft.com/email', 'https://graph.microsoft.com/profile', 'https://graph.microsoft.com/user.read.all', 'https://graph.microsoft.com/group.read.all', 'https://graph.microsoft.com/mailboxsettings.readwrite', 'https://graph.microsoft.com/EWS.AccessAsUser.All') -Silent -ForceRefresh
            } catch {
                try {
                    $auth = $script:msalClientApp | Get-MsalToken -LoginHint $(if ($null -ne $script:CurrentUser) { $script:CurrentUser } else { '' }) -Scopes ('https://graph.microsoft.com/openid', 'https://graph.microsoft.com/email', 'https://graph.microsoft.com/profile', 'https://graph.microsoft.com/user.read.all', 'https://graph.microsoft.com/group.read.all', 'https://graph.microsoft.com/mailboxsettings.readwrite', 'https://graph.microsoft.com/EWS.AccessAsUser.All') -Interactive -Timeout (New-TimeSpan -Minutes 2) -Prompt 'NoPrompt' -UseEmbeddedWebView:$false
                } catch {
                }
            }
        }

        try {
            $script:authorizationHeader = @{
                Authorization = $auth.CreateAuthorizationHeader()
            }
            return @{
                error       = $false
                accessToken = $auth.AccessToken
                authHeader  = $script:authorizationHeader
            }
        } catch {
            return @{
                error       = ($error | Out-String)
                accessToken = $null
                authHeader  = $null
            }
        }
    }
}


function GraphGetMe {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s): User.Read
    # https://docs.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#properties
    # Microsoft Graph REST API v1.0
    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/me`?`$select=" + [System.Web.HttpUtility]::UrlEncode(($GraphUserProperties -join ', '))
            Headers     = $script:authorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $local:x = (Invoke-RestMethod @requestBody)
        $ProgressPreference = $OldProgressPreference
    } catch {
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            me    = $local:x
        }
    } else {
        return @{
            error = $error | Out-String
            me    = $null
        }
    }
}


function GraphGetUserProperties($user) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s): User.Read
    # https://docs.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#properties
    # Microsoft Graph REST API v1.0
    try {
        $local:x = $GraphUserProperties
        if (($user -eq $script:CurrentUser) -and (-not $SimulateUser)) {
            $local:x += 'mailboxsettings'
        }
        $local:x = $local:x -join ','

        $requestBody = @{
            Method      = 'Get'
            Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/users/$user`?`$select=" + [System.Web.HttpUtility]::UrlEncode($local:x)
            Headers     = $script:authorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }
        $local:x = $null
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $local:x = (Invoke-RestMethod @requestBody)
        $ProgressPreference = $OldProgressPreference
    } catch {
    }

    if ($null -ne $local:x) {
        return @{
            error      = $false
            properties = $local:x
        }
    } else {
        return @{
            error      = $error | Out-String
            properties = $null
        }
    }
}


function GraphGetUserManager($user) {
    # Current mailbox manager
    # https://docs.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0&tabs=http
    # Required permission(s): User.Read.All
    # Microsoft Graph REST API v1.0

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/users/$user/manager"
            Headers     = $script:authorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $local:x = Invoke-RestMethod @requestBody
        $ProgressPreference = $OldProgressPreference
    } catch {
    }

    if ($null -ne $local:x) {
        return @{
            error      = $false
            properties = $local:x
        }
    } else {
        return @{
            error      = $error | Out-String
            properties = $null
        }
    }

}


function GraphGetUserTransitiveMemberOf($user) {
    # https://docs.microsoft.com/en-us/graph/api/user-getmembergroups?view=graph-rest-1.0&tabs=http
    # Required permission(s): User.Read
    # Microsoft Graph REST API v1.0
    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/users/$user/transitiveMemberOf"
            Headers     = $script:authorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $x = (Invoke-RestMethod @requestBody).value
        $ProgressPreference = $OldProgressPreference
    } catch {
    }

    if ($null -ne $local:x) {
        return @{
            error    = $false
            memberof = $local:x
        }
    } else {
        return @{
            error    = $error | Out-String
            memberof = $null
        }
    }
}


function GraphGetUserPhoto($user) {
    # https://docs.microsoft.com/en-us/graph/api/profilephoto-get?view=graph-rest-1.0
    # Required permission(s): User.Read
    # Microsoft Graph REST API v1.0
    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/users/$user/photo/" + '$value'
            Headers     = $script:authorizationHeader
            ContentType = 'image/jpg'
        }
        $local:tempFile = (Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ((New-Guid).Guid))
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-RestMethod @requestBody -OutFile $local:tempFile
        $ProgressPreference = $OldProgressPreference

        if ($($PSVersionTable.PSEdition) -ieq 'Core') {
            $local:x = (Get-Content -LiteralPath $local:tempFile -AsByteStream -Raw)
        } else {
            $local:x = (Get-Content -LiteralPath $local:tempFile -Encoding Byte -Raw)
        }

        Remove-Item $local:tempFile -Force
    } catch {
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            photo = $local:x
        }
    } else {
        return @{
            error = $error | Out-String
            photo = $null
        }
    }
}


function GraphPatchUserMailboxsettings($user, $OOFInternal, $OOFExternal) {
    try {
        if ($OOFInternal -or $OOFExternal) {
            $body = @{}
            $body.add('automaticRepliesSetting', @{})
            if ($OOFInternal) { $Body.'automaticRepliesSetting'.add('internalReplyMessage', $OOFInternal ) }
            if ($OOFExternal) { $Body.'automaticRepliesSetting'.add('externalReplyMessage', $OOFExternal ) }
            $body = $body | ConvertTo-Json
            $requestBody = @{
                Method      = 'Patch'
                Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/users/$user/mailboxsettings"
                Headers     = $script:authorizationHeader
                ContentType = 'Application/Json; charset=utf-8'
                Body        = $body
            }
            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            Invoke-RestMethod @requestBody
            $ProgressPreference = $OldProgressPreference
        }

        return @{
            error = $false
        }
    } catch {
        return @{
            error = $error | Out-String
        }
    }
}


function GraphFilterGroups($filter) {
    # https://docs.microsoft.com/en-us/graph/api/group-get?view=graph-rest-1.0&tabs=http
    # Required permission(s): User.Read

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "https://graph.microsoft.com/$GraphEndpointVersion/groups`?`$filter=" + [System.Web.HttpUtility]::UrlEncode($filter)
            Headers     = $script:authorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        $local:x = (Invoke-RestMethod @requestBody).value
        $ProgressPreference = $OldProgressPreference
    } catch {
    }

    if ($null -ne $local:x) {
        return @{
            error  = $false
            groups = $local:x
        }
    } else {
        return @{
            error  = $error | Out-String
            groups = $null
        }
    }
}


function GetIniContent ($filePath) {
    $local:ini = [ordered]@{}
    $local:SectionIndex = -1
    if ($filePath -ne '') {
        try {
            Write-Verbose '    Original ini content'

            foreach ($line in @(Get-Content -LiteralPath $FilePath -Encoding UTF8 -ErrorAction Stop)) {
                Write-Verbose "      $line"
                switch -regex ($line) {
                    # Comments starting with ; or #, or empty line, whitespace(s) before are ignored
                    '(^\s*(;|#))|(^\s*$)' { continue }

                    # Section in square brackets, whitespace(s) before and after brackets are ignored
                    '^\s*\[(.+)\]\s*' {
                        $local:section = ($matches[1]).trim().trim('"').trim('''')
                        if ($null -ne $local:section) {
                            $local:SectionIndex++
                            $local:ini["$($local:SectionIndex)"] = @{ '<Set-OutlookSignatures template>' = $local:section }
                        }
                        continue
                    }

                    # Key and value, whitespace(s) before and after brackets are ignored
                    '^\s*(.+?)\s*=\s*(.*)\s*' {
                        if ($null -ne $local:section) {
                            $local:ini["$($local:SectionIndex)"][($matches[1]).trim().trim('"').trim('''')] = ($matches[2]).trim().trim('"').trim('''')
                            continue
                        }
                    }

                    # Key only, whitespace(s) before and after brackets are ignored
                    '^\s*(.*)\s*' {
                        if ($null -ne $local:section) {
                            $local:ini["$($local:SectionIndex)"][($matches[1]).trim().trim('"').trim('''')] = $null
                            continue
                        }
                    }
                }
            }
        } catch {
            Write-Host
            Write-Host "Error accessing '$FilePath'. Exit." -ForegroundColor red
            $Error[0]
            exit 1
        }
    }
    return $local:ini
}


function ConvertPath ([ref]$path) {
    if ($path) {
        if (($path.value.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($path.value -ilike '*@ssl\*')) {
            $path.value = $path.value -ireplace '@ssl\\', '\'
            $path.value = ([uri]::UnescapeDataString($path.value) -ireplace ('https://', '\\'))
            $path.value = ([System.URI]$path.value).AbsoluteURI -replace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\$2' -replace '/', '\'
            $path.value = [uri]::UnescapeDataString($path.value)
        } else {
            $path.value = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path.value)
            $path.value = ([System.URI]$path.value).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -replace '/', '\'
            $path.value = [uri]::UnescapeDataString($path.value)
        }
    }
}


#
# All functions have been defined above
# Initially executed code starts here
#

try {
    Write-Host
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    $ScriptPassedParameters = $MyInvocation.Line

    main
} catch {
    Write-Host
    Write-Host 'Unexpected error. Exit.' -ForegroundColor red
    $Error[0]
    exit 1
} finally {
    Write-Host
    Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    # Restore original security setting
    if ($null -eq $WordDisableWarningOnIncludeFieldsUpdate) {
        Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction SilentlyContinue | Out-Null
    } else {
        Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $WordDisableWarningOnIncludeFieldsUpdate.DisableWarningOnIncludeFieldsUpdate -ErrorAction SilentlyContinue | Out-Null
    }

    if ($script:COMWord) {
        try {
            $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
        } catch {
        }
        $script:COMWord.Quit([ref]$false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null
        Remove-Variable -Name 'COMWord' -Scope 'script'
    }

    Remove-Module -Name Microsoft.Exchange.WebServices -Force -ErrorAction SilentlyContinue
    Remove-Module -Name MSAL.PS -Force -ErrorAction SilentlyContinue
    if ($script:dllPath) {
        Remove-Item $script:dllPath -Force -ErrorAction SilentlyContinue
    }
    if ($script:msalPath) {
        Remove-Item $script:msalPath -Recurse -Force -ErrorAction SilentlyContinue
    }


    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
}
