<#
.SYNOPSIS
Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.

.DESCRIPTION
Centrally manage and deploy Outlook text signatures and Out of Office auto reply messages.

Signatures and OOF messages can be:
- Generated from templates in DOCX or HTML file format
- Customized with a broad range of variables, including photos, from Active Directory and other sources
- Applied to all mailboxes (including shared mailboxes), specific mailbox groups or specific e-mail addresses, for every primary mailbox across all Outlook profiles
- Assigned time ranges within which they are valid
- Set as default signature for new mails, or for replies and forwards (signatures only)
- Set as default OOF message for internal or external recipients (OOF messages only)
- Set in Outlook Web for the currently logged on user
- Centrally managed only or exist along user created signatures (signatures only)
- Copied to an alternate path for easy access on mobile devices not directly supported by this script (signatures only)

Sample templates for signatures and OOF messages demonstrate all available features and are provided as .docx and .htm files.

Simulation mode allows content creators and admins to simulate the behavior of the script and to inspect the resulting signature files before going live.

The script is designed to work in big and complex environments (Exchange resource forest scenarios, across AD trusts, multi-level AD subdomains, many objects). The script is **multi-client capable** by using different template paths, configuration files and script parameters.

The script is **Free and Open-Source Software (FOSS)**. It is published under the MIT license which is approved, among others, by the Free Software Foundation (FSF) and the Open Source Initiative (OSI), and is compatible with the General Public License (GPL) v3. Please see `'.\docs\LICENSE.txt'` for copyright and MIT license details.

.LINK
Github: https://github.com/GruberMarkus/Set-OutlookSignatures

.PARAMETER SignatureTemplatePath
Path to centrally managed signature templates.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
Default value: '.\templates\Signatures DOCX'

.PARAMETER ReplacementVariableConfigFile
Path to a replacement variable config file.
Local and remote paths are supported.
Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
Default value: '.\config\default replacement variables.txt'

.PARAMETER DomainsToCheckForGroups
List of domains/forests to check for group membership across trusts.
If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered.
If a string starts with a minus or dash ("-domain-a.local"), the domain after the dash or minus is removed from the list.
Default value: '*'

.PARAMETER DeleteUserCreatedSignatures
Shall the script delete signatures which were created by the user itself?
The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository.
Default value: $false

.PARAMETER SetCurrentUserOutlookWebSignature
Shall the script set the Outlook Web signature of the currently logged on user?
If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used.
Default value: $true

.PARAMETER SetCurrentUserOOFMessage
Shall the script set the Out of Office (OOF) auto reply message of the currently logged on user?
If the parameter is set to `$true` and the current user's mailbox is not configured in any Outlook profile, the current user's mailbox is considered nevertheless. This way, the script can be used in environments where only Outlook Web is used.
Default value: $true

.PARAMETER OOFTemplatePath
Path to centrally managed signature templates.
Local and remote paths are supported.
Local paths can be absolute ('C:\OOF templates') or relative to the script path ('.\templates\Out of Office').
WebDAV paths are supported (https only): 'https://server.domain/SignatureSite/OOFTemplates' or '\\server.domain@SSL\SignatureSite\OOFTemplates'
The currently logged on user needs at least read access to the path.
Default value: '.\templates\Out of Office DOCX'

.PARAMETER AdditionalSignaturePath
An additional path that the signatures shall be copied to.
Ideally, this path is available on all devices of the user, for example via Microsoft OneDrive or Nextcloud.
This way, the user can easily copy-paste the preferred preconfigured signature for use in an e-mail app not supported by this script, such as Microsoft Outlook Mobile, Apple Mail, Google Gmail or Samsung Email.
Local and remote paths are supported.
Local paths can be absolute ('C:\Outlook signatures') or relative to the script path ('.\Outlook signatures').
WebDAV paths are supported (https only): 'https://server.domain/User' or '\\server.domain@SSL\User'
The currently logged on user needs at least write access to the path.
Default value: "$([environment]::GetFolderPath('MyDocuments'))"

.PARAMETER AdditionalSignaturePathFolder
A folder or folder structure below AdditionalSignaturePath.
If the folder or folder structure does not exist, it is created.
Default value: 'Outlook signatures'

.PARAMETER UseHtmTemplates
With this parameter, the script searches for templates with the extension .htm instead of .docx.
Each format has advantages and disadvantages, please see "Should I use .docx or .htm as file format for templates? Signatures in Outlook sometimes look different than my templates." for a quick overview.
Default value: \$false

.PARAMETER SimulateUser
SimulateUser is a mandatory parameter for simulation mode. This value replaces the currently logged on user.
Use a logon name in the format 'Domain\User' or a Universal Principal Name (UPN, looks like an e-mail-address, but is not neecessarily one).

.PARAMETER SimulateMailboxes
SimulateMailboxes is optional for simulation mode, although highly recommended.
It is a comma separated list of e-mail addresses replacing the list of mailboxes otherwise gathered from the registry.

.INPUTS
None. You cannot pipe objects to Set-OutlookSignatures.ps1.

.OUTPUTS
Set-OutlookSignatures.ps1 writes the current activities, warnings and error messages to the standard output stream.

.EXAMPLE
PS> .\Set-OutlookSignatures.ps1

.EXAMPLE
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates'

.EXAMPLE
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -DomainsToCheckForGroups '*', '-internal-test.example.com'

.EXAMPLE
PS> .\Set-OutlookSignatures.ps1 -SignatureTemplatePath '\\internal.example.com\share\Signature Templates' -DomainsToCheckForGroups 'internal-test.example.com', 'company.b.com'

.EXAMPLE
PowerShell.exe -Command "& '\\server\share\directory\Set-OutlookSignatures.ps1' -SignatureTemplatePath '\\server\share\directory\templates\Signatures DOCX' -OOFTemplatePath '\\server\share\directory\templates\Out of Office DOCX' -ReplacementVariableConfigFile '\\server\share\directory\config\default replacement variables.ps1'"
Passing arguments to PowerShell.exe from the command line or task scheduler can be very tricky when spaces are involved. See readme for details.

.EXAMPLE
Please see '.\docs\README.html' and https://github.com/GruberMarkus/Set-OutlookSignatures for more details.

.NOTES
Script : Set-OutlookSignatures
Version: v2.2.0-alpha.19
Web    : https://github.com/GruberMarkus/Set-OutlookSignatures
License: MIT license (see '.\docs\LICENSE.txt' for details and copyright)
#>


[CmdletBinding(PositionalBinding = $false)]

Param(
    # Path to centrally managed signature templates
    #   Local and remote paths are supported
    #     Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signatures')
    #   WebDAV paths are supported (https only)
    #     'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
    #   The currently logged on user needs at least read access to the path
    [ValidateNotNullOrEmpty()][string]$SignatureTemplatePath = '.\templates\Signatures DOCX',

    # Path to a replacement variable config file.
    #   Local and remote paths are supported
    #     Local paths can be absolute ('C:\Signature templates') or relative to the script path ('.\templates\Signature')
    #   WebDAV paths are supported (https only)
    #     'https://server.domain/SignatureSite/SignatureTemplates' or '\\server.domain@SSL\SignatureSite\SignatureTemplates'
    #   The currently logged on user needs at least read access to the path
    [ValidateNotNullOrEmpty()][string]$ReplacementVariableConfigFile = '.\config\default replacement variables.ps1',

    # List of domains/forests to check for group membership across trusts
    #   If the first entry in the list is '*', all outgoing and bidirectional trusts in the current user's forest are considered
    #   If a string starts with a minus or dash ("-domain-a.local"), the domain after the dash or minus is removed from the list
    [string[]]$DomainsToCheckForGroups = ('*'),

    # Shall the script delete signatures which were created by the user itself?
    #   The script always deletes signatures which were deployed by the script earlier, but are no longer available in the central repository.
    [bool]$DeleteUserCreatedSignatures = $false,

    # Shall the script set the Outlook Web signature of the currently logged on user?
    [bool]$SetCurrentUserOutlookWebSignature = $true,

    # Shall the script set the Out of Office (OOF) auto reply message(s) of the currently logged on user?
    [bool]$SetCurrentUserOOFMessage = $true,

    # Path to centrally managed Out of Office (OOF, automatic reply) templates
    #   Local and remote paths are supported
    #     Local paths can be absolute ('C:\OOF templates') or relative to the script path ('.\templates\Out of Office')
    #   WebDAV paths are supported (https only)
    #     'https://server.domain/SignatureSite/OOFTemplates' or '\\server.domain@SSL\SignatureSite\OOFTemplates'
    #   The currently logged on user needs at least read access to the path
    [ValidateNotNullOrEmpty()][string]$OOFTemplatePath = '.\templates\Out of Office DOCX',

    # An additional path that the signatures shall be copied to
    [string]$AdditionalSignaturePath = $(try { $([environment]::GetFolderPath('MyDocuments')) }catch {}),

    # Subfolder to create in $AdditionalSignaturePath
    [string]$AdditionalSignaturePathFolder = 'Outlook Signatures',

    # Use templates in .HTM file format instead of .DOCX
    [switch]$UseHtmTemplates = $false,

    # Simulate another user as currently logged on user
    [Alias('SimulationUser')]
    [string]$SimulateUser = $null,

    # Simulate list of mailboxes instead of mailboxes configured in Outlook
    # Works only together with SimulateUser
    [Alias('SimulationMailboxes')]
    [string[]]$SimulateMailboxes = ('')
)


function main {
    Set-Location $PSScriptRoot | Out-Null

    $Search = New-Object DirectoryServices.DirectorySearcher
    $Search.PageSize = 1000

    $script:jobs = New-Object System.Collections.ArrayList

    Add-Type -AssemblyName System.DirectoryServices.AccountManagement

    $HTMLMarkerTag = '<meta name=data-SignatureFileInfo content="Set-OutlookSignatures.ps1">'

    if (-not (Test-Path 'variable:IsWindows')) {
        # Automatic variable $IsWindows not set, must be Powershell version lower than 6 running on Windows
        $IsWindows = $true
        $IsLinux = $IsMacOS = $false
    }

    $script:tempDir = [System.IO.Path]::GetTempPath()

    Write-Host
    Write-Host "Script notes @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    (((Get-Help -Full $PSCommandPath).alertSet.alert.Text) -split "`r?`n").Trim() | ForEach-Object {
        $x = ($_.split(':', 2)).trim()
        Write-Host "  $($x[0].trim()): $($x[1].trim())"
    }


    Write-Host
    Write-Host "Check parameters and script environment @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    if ($IsWindows -eq $false) {
        Write-Host '  This script is supported on Windows, but not on Linux or macOS. Exiting.' -ForegroundColor Red
        exit 1
    }

    if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
        Write-Host "  This PowerShell session is running in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode." -ForegroundColor Red
        Write-Host '  Required features are only available in FullLanguage mode. Exiting.' -ForegroundColor Red
        exit 1
    }

    Write-Host "  Script name: '$PSCommandPath'"
    Write-Host "  Script path: '$PSScriptRoot'"
    Write-Host "  ReplacementVariableConfigFile: '$ReplacementVariableConfigFile'" -NoNewline
    CheckPath $ReplacementVariableConfigFile
    Write-Host "  SignatureTemplatePath: '$SignatureTemplatePath'" -NoNewline
    CheckPath $SignatureTemplatePath
    Write-Host ('  DomainsToCheckForGroups: ' + ('''' + $($DomainsToCheckForGroups -join ''', ''') + ''''))
    Write-Host "  DeleteUserCreatedSignatures: '$DeleteUserCreatedSignatures'"
    Write-Host "  SetCurrentUserOutlookWebSignature: '$SetCurrentUserOutlookWebSignature'"
    Write-Host "  SetCurrentUserOOFMessage: '$SetCurrentUserOOFMessage'"
    if ($SetCurrentUserOOFMessage) {
        Write-Host "  OOFTemplatePath: '$OOFTemplatePath'" -NoNewline
        CheckPath $OOFTemplatePath
    }
    Write-Host "  AdditionalSignaturePath: '$AdditionalSignaturePath'" -NoNewline
    CheckPath $AdditionalSignaturePath
    Write-Host "  AdditionalSignaturePathFolder: '$AdditionalSignaturePathFolder'"
    Write-Host "  UseHtmTemplates: '$UseHtmTemplates'"
    Write-Host "  SimulateUser: '$SimulateUser'"
    Write-Host ('  SimulateMailboxes: ' + ('''' + $($SimulateMailboxes -join ''', ''') + ''''))

    if ($AdditionalSignaturePathFolder -and $AdditionalSignaturePath) {
        $AdditionalSignaturePath = ((Join-Path -Path ($AdditionalSignaturePath) -ChildPath $AdditionalSignaturePathFolder))
        try {
            if (-not (Test-Path -LiteralPath $AdditionalSignaturePath -PathType Container)) {
                New-Item -Path $AdditionalSignaturePath -ItemType directory -Force | Out-Null
                if (-not (Test-Path -LiteralPath $AdditionalSignaturePath -PathType Container)) {
                    throw
                }
            }
            if ($SimulateUser) {
                New-Item -Path ((Join-Path -Path ($AdditionalSignaturePath) -ChildPath 'OOF')) -ItemType directory -Force | Out-Null
                if (-not (Test-Path -LiteralPath ((Join-Path -Path ($AdditionalSignaturePath) -ChildPath 'OOF')) -PathType Container)) {
                    throw
                }
                Get-ChildItem ((Join-Path -Path ($AdditionalSignaturePath) -ChildPath 'OOF\*')) -Recurse | Remove-Item -Force -Recurse -Confirm:$false
            }
        } catch {
            Write-Host "    Problem connecting to, creating or reading from folder '$AdditionalSignaturePath'. Deactivating feature." -ForegroundColor Yellow
            $AdditionalSignaturePath = ''
        }
    }

    ('ReplacementVariableConfigFile', 'SignatureTemplatePath', 'OOFTemplatePath', 'AdditionalSignaturePath') | ForEach-Object {
        $path = (Get-Variable -Name $_).Value
        if ($path.StartsWith('https://', 'CurrentCultureIgnoreCase')) {
            $path = ((([uri]::UnescapeDataString($path) -ireplace ('https://', '\\')) -replace ('(.*?)/(.*)', '${1}@SSL\$2')) -replace ('/', '\'))
        } else {
            $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
            $path = $path
        }
        Set-Variable -Name $_ -Value $path
    }

    if ($SimulateUser) {
        Write-Host
        Write-Host 'Simulation mode enabled' -ForegroundColor Yellow
    }


    Write-Host
    Write-Host "Get Outlook version and profile @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled, skipping task' -ForegroundColor Yellow
    } else {
        $OutlookRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace 'Outlook.Application.', '') + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))

        if ($OutlookRegistryVersion.major -eq 0) {
            $OutlookRegistryVersion = $null
        } elseif ($OutlookRegistryVersion.major -gt 16) {
            Write-Host "Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exiting." -ForegroundColor Red
            exit 1
        } elseif ($OutlookRegistryVersion.major -eq 16) {
            $OutlookRegistryVersion = '16.0'
        } elseif ($OutlookRegistryVersion.major -eq 15) {
            $OutlookRegistryVersion = '15.0'
        } elseif ($OutlookRegistryVersion.major -eq 14) {
            $OutlookRegistryVersion = '14.0'
        } elseif ($OutlookRegistryVersion.major -lt 14) {
            Write-Host "Outlook version $OutlookRegistryVersion is older than Outlook 2010 and not supported. Please inform your administrator. Exiting." -ForegroundColor Red
            exit 1
        }

        if ($null -ne $OutlookRegistryVersion) {
            $OutlookDefaultProfile = (Get-ItemProperty "hkcu:\software\microsoft\office\$OutlookRegistryVersion\Outlook" -ErrorAction SilentlyContinue).DefaultProfile
        } else {
            $OutlookDefaultProfile = $null
        }

        Write-Host "  Outlook registry version: $OutlookRegistryVersion"
        Write-Host "  Outlook default profile: $OutlookDefaultProfile"
    }


    Write-Host
    Write-Host "Get Outlook signature file path(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $SignaturePaths = @()
    if ($SimulateUser) {
        $SignaturePaths = $AdditionalSignaturePath
        Write-Host '  Simulation mode enabled, skipping task, using AdditionalSignaturePath instead' -ForegroundColor Yellow
    } else {
        Get-ItemProperty 'hkcu:\software\microsoft\office\*\common\general' -ErrorAction SilentlyContinue | Where-Object { $_.'Signatures' -ne '' } | ForEach-Object {
            Push-Location ((Join-Path -Path ($env:AppData) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_.Signatures))
            if (Test-Path $x -IsValid) {
                if (-not (Test-Path $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force
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
        Write-Host '  Simulation mode enabled, skipping task, using SimulateMailboxes instead' -ForegroundColor Yellow
        for ($i = 0; $i -lt $SimulateMailboxes.count; $i++) {
            $MailAddresses += $SimulateMailboxes[$i]
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''
        }
    } else {
        Get-ItemProperty "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*" -ErrorAction SilentlyContinue | Where-Object { (($_.'Account Name' -like '*@*.*') -and ($_.'Identity Eid' -ne '')) } | ForEach-Object {
            $MailAddresses += ($_.'Account Name').tolower()
            $RegistryPaths += $_.PSPath
            $LegacyExchangeDN = ('/O=' + (((($_.'Identity Eid' | ForEach-Object { [char]$_ }) -join '' -replace [char]0) -split '/O=')[-1]).ToString().trim())
            if ($LegacyExchangeDN.length -le 3) {
                $LegacyExchangeDN = ''
            }
            $LegacyExchangeDNs += $LegacyExchangeDN
            Write-Host "  $($_.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $_.PSDrive)"
            Write-Host "    $($_.'Account Name')"
        }
    }


    Write-Host
    Write-Host "Enumerate domains @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $x = $DomainsToCheckForGroups
    [System.Collections.ArrayList]$DomainsToCheckForGroups = @()

    # Users own domain/forest is always included
    $y = ([ADSI]"LDAP://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')
    if ($y -ne '') {
        Write-Host "  Current user forest: $y"
        $DomainsToCheckForGroups += $y
    } else {
        Write-Host '  Problem connecting to Active Directory, or user is a local user. Exiting.' -ForegroundColor Red
        exit 1
    }

    # Other domains - either the list provided, or all outgoing and bidirectional trusts
    if ($x[0] -eq '*') {
        $Search.SearchRoot = "GC://$($DomainsToCheckForGroups[0])"
        $Search.Filter = '(ObjectClass=trustedDomain)'

        $Search.FindAll() | ForEach-Object {
            # DNS name of this side of the trust (could be the root domain or any subdomain)
            # $TrustOrigin = ($_.properties.distinguishedname -split ',DC=')[1..999] -join '.'

            # DNS name of the other side of the trust (could be the root domain or any subdomain)
            # $TrustName = $_.properties.name

            # Domain SID of the other side of the trust
            # $TrustNameSID = (New-Object system.security.principal.securityidentifier($($_.properties.securityidentifier), 0)).tostring()

            # Trust direction
            # https://docs.microsoft.com/en-us/dotnet/api/system.directoryservices.activedirectory.trustdirection?view=net-5.0
            # $TrustDirectionNumber = $_.properties.trustdirection

            # Trust type
            # https://docs.microsoft.com/en-us/dotnet/api/system.directoryservices.activedirectory.trusttype?view=net-5.0
            # $TrustTypeNumber = $_.properties.trusttype

            # Trust attributes
            # https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-adts/e9a2d23c-c31e-4a6f-88a0-6646fdb51a3c
            # $TrustAttributesNumber = $_.properties.trustattributes

            # Which domains does the current user have access to?
            # No intra-forest trusts, only bidirectional trusts and outbound trusts

            if (($($_.properties.trustattributes) -ne 32) -and (($($_.properties.trustdirection) -eq 2) -or ($($_.properties.trustdirection) -eq 3)) ) {
                Write-Host "  Trusted domain: $($_.properties.name)"
                $DomainsToCheckForGroups += $_.properties.name
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
            Write-Host '    Skipping domain. Entry * is only allowed at first position in list.' -ForegroundColor Red
            continue
        }

        if ($y -match '[^a-zA-Z0-9.-]') {
            Write-Host '    Skipping domain. Allowed characters are a-z, A-Z, ., -.' -ForegroundColor Red
            continue
        }

        if (-not ($y.StartsWith('-'))) {
            if ($DomainsToCheckForGroups -icontains $y) {
                Write-Host '    Domain already in list.' -ForegroundColor Yellow
            } else {
                $DomainsToCheckForGroups += $y
            }
        } else {
            Write-Host '    Removing domain.'
            for ($z = 0; $z -lt $DomainsToCheckForGroups.Count; $z++) {
                if ($DomainsToCheckForGroups[$z] -ilike $y.substring(1)) {
                    $DomainsToCheckForGroups[$z] = ''
                }
            }
        }
    }


    Write-Host
    Write-Host "Check for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    CheckADConnectivity $DomainsToCheckForGroups 'LDAP' '  ' | Out-Null


    Write-Host
    Write-Host "Check for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    CheckADConnectivity $DomainsToCheckForGroups 'GC' '  ' | Out-Null


    Write-Host
    Write-Host "Get AD properties of currently logged on user and assigned manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (-not $SimulateUser) {
        Write-Host '  Currently logged on user'
    } else {
        Write-Host "  Simulating '$SimulateUser' as currently logged on user" -ForegroundColor Yellow
    }
    try {
        if (-not $SimulateUser) {
            $Search.SearchRoot = "GC://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
            $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"
            $ADPropsCurrentUser = $Search.FindOne().Properties
            if ((($SetCurrentUserOutlookWebSignature -eq $true) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses -notcontains $ADPropsCurrentUser.mail)) {
                # OOF and/or Outlook web signature must be set, but user does not seem to have a mailbox in Outlook
                # Maybe this is a pure Outlook Web user, so we will add a helper entry
                # This entry fakes the users mailbox in his default Outlook profile, so it gets the highest priority later
                Write-Host "    User's mailbox not found in Outlook profiles, but Outlook Web signature and/or OOF message should be set. Adding Mailbox dummy entry." -ForegroundColor Yellow
                $script:CurrentUserDummyMailbox = $true
                $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname)) + $SignaturePaths
                $MailAddresses = @($ADPropsCurrentUser.mail) + $MailAddresses
                $RegistryPaths = @("hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\") + $RegistryPaths
                $LegacyExchangeDNs = @('') + $LegacyExchangeDNs
            } else {
                $script:CurrentUserDummyMailbox = $false
            }
        } else {
            try {
                $SimulateUserSID = (New-Object System.Security.Principal.NTAccount($SimulateUser)).Translate([System.Security.Principal.SecurityIdentifier]).value
                $SimulateUserDN = ([adsi]"LDAP://<SID=$SimulateUserSID>").distinguishedname
                $Search.SearchRoot = "GC://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                $Search.Filter = "((distinguishedname=$SimulateUserDN))"
                $ADPropsCurrentUser = $Search.FindOne().Properties
            } catch {
                Write-Host "    Simulation user '$($SimulateUser)' not found. Exiting." -ForegroundColor REd
                $error[0]
                exit 1
            }
        }
    } catch {
        $ADPropsCurrentUser = $null
        Write-Host '    Problem connecting to Active Directory, or user is a local user. Exiting.' -ForegroundColor Red
        $error[0]
        exit 1
    }

    Write-Host "    $($ADPropsCurrentUser.distinguishedname)"

    if (-not $SimulateUser) {
        Write-Host '  Manager of currently logged on user'
    } else {
        Write-Host '  Manager of simulated currently logged on user' -ForegroundColor Yellow
    }
    try {
        $Search.SearchRoot = "GC://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
        $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"
        $ADPropsCurrentUserManager = $Search.FindOne().Properties
    } catch {
        $ADPropsCurrentUserManager = $null
    }
    if ($ADPropsCurrentUserManager) { Write-Host "    $($ADPropsCurrentUserManager.distinguishedname)" }


    Write-Host
    Write-Host "Get AD properties of each mailbox @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $ADPropsMailboxes = @()
    $ADPropsMailboxesUserDomain = @()

    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        Write-Host "  Mailbox $($MailAddresses[$AccountNumberRunning])"

        $UserDomain = ''
        $ADPropsMailboxes += $null
        $ADPropsMailboxesUserDomain += $null

        if ((($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '') -or ($($MailAddresses[$AccountNumberRunning]) -ne ''))) {
            # Loop through domains until the first one knows the legacyExchangeDN or the proxy address
            for ($DomainNumber = 0; (($DomainNumber -lt $DomainsToCheckForGroups.count) -and ($UserDomain -eq '')); $DomainNumber++) {
                if (($DomainsToCheckForGroups[$DomainNumber] -ne '')) {
                    Write-Host "    $($DomainsToCheckForGroups[$DomainNumber]) (searching for mailbox user object) ... " -NoNewline
                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($DomainsToCheckForGroups[$DomainNumber])")
                    if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                        $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(msExchMailboxGuid=*)(legacyExchangeDN=$($LegacyExchangeDNs[$AccountNumberRunning])))"
                    } elseif (($($MailAddresses[$AccountNumberRunning]) -ne '')) {
                        $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(msExchMailboxGuid=*)(legacyExchangeDN=*)(proxyaddresses=smtp:$($MailAddresses[$AccountNumberRunning])))"
                    }
                    $u = $Search.FindAll()
                    if ($u.count -eq 0) {
                        Write-Host
                        Write-Host "      '$($MailAddresses[$AccountNumberRunning])' matches no Exchange mailbox." -ForegroundColor Yellow
                    } elseif ($u.count -gt 1) {
                        Write-Host
                        Write-Host "      '$($MailAddresses[$AccountNumberRunning])' matches multiple Exchange mailboxes, ignoring." -ForegroundColor Yellow
                        $u | ForEach-Object { Write-Host "          $($_.path)" -ForegroundColor Yellow }
                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $MailAddresses[$AccountNumberRunning] = ''
                        $UserDomain = $null
                    } else {
                        # Connect to Domain Controller (LDAP), as Global Catalog (GC) does not have all attributes,
                        # for example tokenGroups including domain local groups
                        $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"
                        $ADPropsMailboxes[$AccountNumberRunning] = $Search.FindOne().Properties
                        $UserDomain = $DomainsToCheckForGroups[$DomainNumber]
                        $ADPropsMailboxesUserDomain[$AccountNumberRunning] = $DomainsToCheckForGroups[$DomainNumber]
                        $LegacyExchangeDNs[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].legacyexchangedn
                        $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()
                        Write-Host 'found'
                        Write-Host "      $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                    }
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
        Write-Host "  AD mail attribute of currently logged on user: $($ADPropsCurrentUser.mail)"
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
        Write-Host '  AD mail attribute of currently logged on user is empty, searching msExchMasterAccountSid'
        # No mail attribute set, check for match(es) of user's objectSID and mailbox's msExchMasterAccountSid
        for ($i = 0; $i -lt $MailAddresses.count; $i++) {
            if ($ADPropsMailboxes[$i].msexchmasteraccountsid) {
                if (((New-Object System.Security.Principal.SecurityIdentifier $ADPropsMailboxes[$i].msexchmasteraccountsid[0], 0).value -ieq (New-Object System.Security.Principal.SecurityIdentifier $ADPropsCurrentUser.objectsid[0], 0).Value)) {
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

    }

    $MailboxNewOrder = @()
    $PrimaryMailboxAddress = $null

    if ($p -ge 0) {
        $MailboxNewOrder += $p
        $PrimaryMailboxAddress = $MailAddresses[$p]
    }

    for ($i = 0; $i -le $RegistryPaths.length - 1; $i++) {
        if (($RegistryPaths[$i] -ilike "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\*") -and ($i -ne $p) -and ($LegacyExchangeDNs[$i])) {
            $MailboxNewOrder += $i
        }
    }

    for ($i = 0; $i -le $RegistryPaths.length - 1; $i++) {
        if (($RegistryPaths[$i] -notlike "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\*") -and ($i -ne $p) -and ($LegacyExchangeDNs[$i])) {
            $MailboxNewOrder += $i
        }
    }

    ('RegistryPaths', 'MailAddresses', 'LegacyExchangeDNs', 'ADPropsMailboxesUserDomain', 'ADPropsMailboxes') | ForEach-Object {
        (Get-Variable -Name $_).value = (Get-Variable -Name $_).value[$MailboxNewOrder]
    }
    Write-Host '  Mailbox priority (highest to lowest):'
    $MailAddresses | ForEach-Object {
        Write-Host "    $_"
    }


    Write-Host
    Write-Host "Get all signature template files and categorize them @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $SignatureFilesCommon = @{}
    $SignatureFilesGroup = @{}
    $SignatureFilesGroupFilePart = @{}
    $SignatureFilesMailbox = @{}
    $SignatureFilesMailboxFilePart = @{}
    $SignatureFilesDefaultNew = @{}
    $SignatureFilesDefaultReplyFwd = @{}
    $script:SignatureFilesDone = @()
    $SignatureFilesGroupSIDs = @{}

    foreach ($SignatureFile in ((Get-ChildItem -LiteralPath $SignatureTemplatePath -File -Filter $(if ($UseHtmTemplates) { '*.htm' } else { '*.docx' })) | Sort-Object)) {
        Write-Host ("  '$($SignatureFile.Name)'")
        $x = $SignatureFile.name -split '\.(?![\w\s\d]*\[*(\]|@))'
        if ($x.count -ge 3) {
            $SignatureFilePart = $x[-2]
            $SignatureFileTargetName = ($x[($x.count * -1)..-3] -join '.') + '.' + $x[-1]
        } else {
            $SignatureFilePart = ''
            $SignatureFileTargetName = $SignatureFile.Name
        }

        $SignatureFileTimeActive = $true
        if ($SignatureFilePart -match '\[\d{12}-\d{12}\]') {
            $SignatureFileTimeActive = $false
            Write-Host '    Time based signature'
            foreach ($SignatureFilePartTag in ([regex]::Matches((($SignatureFilePart -replace '(?i)\[DefaultNew\]', '') -replace '(?i)\[DefaultReplyFwd\]', ''), '\[\d{12}-\d{12}\]').captures.value)) {
                foreach ($DateTimeTag in ([regex]::Matches($SignatureFilePartTag, '\[\d{12}-\d{12}\]').captures.value)) {
                    Write-Host "      $($DateTimeTag): " -NoNewline
                    try {
                        $DateTimeTagStart = [System.DateTime]::ParseExact(($DateTimeTag.tostring().Substring(1, 12)), 'yyyyMMddHHmm', $null)
                        $DateTimeTagEnd = [System.DateTime]::ParseExact(($DateTimeTag.tostring().Substring(14, 12)), 'yyyyMMddHHmm', $null)

                        if (((Get-Date) -ge $DateTimeTagStart) -and ((Get-Date) -le $DateTimeTagEnd)) {
                            Write-Host 'Current DateTime in range'
                            $SignatureFileTimeActive = $true
                        } else {
                            Write-Host 'Current DateTime out of range'
                        }
                    } catch {
                        Write-Host 'Invalid DateTime, ignoring tag' -ForegroundColor Red
                    }
                }
            }
            if ($SignatureFileTimeActive -eq $true) {
                Write-Host '      Current DateTime is in range of at least one DateTime tag, using signature'
            } else {
                Write-Host '      Current DateTime is not in range of any DateTime tag, ignoring signature' -ForegroundColor Yellow
            }
        }

        if ($SignatureFileTimeActive -ne $true) {
            continue
        }

        [regex]::Matches((($SignatureFilePart -replace '(?i)\[DefaultNew\]', '') -replace '(?i)\[DefaultReplyFwd\]', ''), '\[(.*?)\]').captures.value | ForEach-Object {
            $SignatureFilePartTag = $_
            if ($SignatureFilePartTag -match '\[(.*?)@(.*?)\.(.*?)\]') {
                if (-not $SignatureFilesMailbox.ContainsKey($SignatureFile.FullName)) {
                    Write-Host '    Mailbox specific signature'
                    $SignatureFilesMailbox.add($SignatureFile.FullName, $SignatureFileTargetName)
                }
                Write-Host "      $($SignatureFilePartTag -replace '\[' -replace '\]')"
                $SignatureFilesMailboxFilePart[$SignatureFile.FullName] = ($SignatureFilesMailboxFilePart[$SignatureFile.FullName] + $SignatureFilePartTag)
            } elseif ($SignatureFilePartTag -match '\[.*? .*?\]') {
                if (-not $SignatureFilesGroup.ContainsKey($SignatureFile.FullName)) {
                    Write-Host '    Group specific signature'
                    $SignatureFilesGroup.add($SignatureFile.FullName, $SignatureFileTargetName)
                }
                $NTName = ((($SignatureFilePartTag -replace '\[', '') -replace '\]', '') -replace '(.*?) (.*)', '$1\$2')
                if (-not $SignatureFilesGroupSIDs.ContainsKey($SignatureFilePartTag)) {
                    try {
                        $SignatureFilesGroupSIDs.add($SignatureFilePartTag, (New-Object System.Security.Principal.NTAccount($NTName)).Translate([System.Security.Principal.SecurityIdentifier]))
                    } catch {
                        # No group with this sAMAccountName found. Maybe it's a display name?
                        try {
                            $objTrans = New-Object -ComObject 'NameTranslate'
                            $objNT = $objTrans.GetType()
                            $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                            $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (4, ($NTName -split '\\')[1]))
                            $SignatureFilesGroupSIDs.add($SignatureFilePartTag, ((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                        } catch {
                        }
                    }
                }

                if ($SignatureFilesGroupSIDs.containskey($SignatureFilePartTag)) {
                    Write-Host "      $SignatureFilePartTag = $NTName = $($SignatureFilesGroupSIDs[$SignatureFilePartTag])"
                    $SignatureFilesGroupFilePart[$SignatureFile.FullName] = ($SignatureFilesGroupFilePart[$SignatureFile.FullName] + '[' + $SignatureFilesGroupSIDs[$SignatureFilePartTag] + ']')
                } else {
                    Write-Host "      $SignatureFilePartTag = $($NTName): Not found in Active Directory, please check" -ForegroundColor Yellow
                }
            } else {
                Write-Host '    Common signature'
                if (-not $SignatureFilesCommon.containskey($SignatureFile.FullName)) {
                    $SignatureFilesCommon.add($SignatureFile.FullName, $SignatureFileTargetName)
                }
            }
        }

        if ($SignatureFilePart -match '(?i)\[DefaultNew\]') {
            $SignatureFilesDefaultNew.add($SignatureFile.FullName, $SignatureFileTargetName)
            Write-Host '    Default signature for new mails'
        }

        if ($SignatureFilePart -match '(?i)\[DefaultReplyFwd\]') {
            $SignatureFilesDefaultReplyFwd.add($SignatureFile.FullName, $SignatureFileTargetName)
            Write-Host '    Default signature for replies and forwards'
        }
    }


    if ($SetCurrentUserOOFMessage) {
        Write-Host
        Write-Host "Get all Out of Office (OOF) auto reply template files and categorize them @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $OOFFilesCommon = @{}
        $OOFFilesGroup = @{}
        $OOFFilesGroupFilePart = @{}
        $OOFFilesMailbox = @{}
        $OOFFilesMailboxFilePart = @{}
        $OOFFilesInternal = @{}
        $OOFFilesExternal = @{}
        $script:OOFFilesDone = @()
        $OOFFilesGroupSIDs = @{}

        foreach ($OOFFile in ((Get-ChildItem -LiteralPath $OOFTemplatePath -File -Filter $(if ($UseHtmTemplates) { '*.htm' } else { '*.docx' })) | Sort-Object)) {
            Write-Host ("  '$($OOFFile.Name)'")
            $x = $OOFFile.name -split '\.(?![\w\s\d]*\[*(\]|@))'
            if ($x.count -ge 3) {
                $OOFFilePart = $x[-2]
                $OOFFileTargetName = ($x[($x.count * -1)..-3] -join '.') + '.' + $x[-1]
            } else {
                $OOFFilePart = ''
                $OOFFileTargetName = $OOFFile.Name
            }

            $OOFFileTimeActive = $true
            if ($OOFFilePart -match '\[\d{12}-\d{12}\]') {
                $OOFFileTimeActive = $false
                Write-Host '    Time based OOF message'
                foreach ($OOFFilePartTag in ([regex]::Matches((($OOFFilePart -replace '(?i)\[Internal\]', '') -replace '(?i)\[External\]', ''), '\[\d{12}-\d{12}\]').captures.value)) {
                    foreach ($DateTimeTag in ([regex]::Matches($OOFFilePartTag, '\[\d{12}-\d{12}\]').captures.value)) {
                        Write-Host "      $($DateTimeTag): " -NoNewline
                        try {
                            $DateTimeTagStart = [System.DateTime]::ParseExact(($DateTimeTag.tostring().Substring(1, 12)), 'yyyyMMddHHmm', $null)
                            $DateTimeTagEnd = [System.DateTime]::ParseExact(($DateTimeTag.tostring().Substring(14, 12)), 'yyyyMMddHHmm', $null)

                            if (((Get-Date) -ge $DateTimeTagStart) -and ((Get-Date) -le $DateTimeTagEnd)) {
                                Write-Host 'Current DateTime in range'
                                $OOFFileTimeActive = $true
                            } else {
                                Write-Host 'Current DateTime out of range'
                            }
                        } catch {
                            Write-Host 'Invalid DateTime, ignoring tag' -ForegroundColor Red
                        }
                    }
                }
                if ($OOFFileTimeActive -eq $true) {
                    Write-Host '      Current DateTime is in range of at least one DateTime tag, using OOF message'
                } else {
                    Write-Host '      Current DateTime is not in range of any DateTime tag, ignoring OOF message' -ForegroundColor Yellow
                }
            }

            if ($OOFFileTimeActive -ne $true) {
                continue
            }

            [regex]::Matches((($OOFFilePart -replace '(?i)\[External\]', '') -replace '(?i)\[Internal\]', ''), '\[(.*?)\]').captures.value | ForEach-Object {
                $OOFFilePartTag = $_
                if ($OOFFilePartTag -match '\[(.*?)@(.*?)\.(.*?)\]') {
                    if (-not $OOFFilesMailbox.ContainsKey($OOFFile.FullName)) {
                        Write-Host '    Mailbox specific OOF message'
                        $OOFFilesMailbox.add($OOFFile.FullName, $OOFFileTargetName)
                    }
                    Write-Host "      $($OOFFilePartTag -replace '\[' -replace '\]')"
                    $OOFFilesMailboxFilePart[$OOFFile.FullName] = ($OOFFilesMailboxFilePart[$OOFFile.FullName] + $OOFFilePartTag)
                } elseif ($OOFFilePartTag -match '\[.*? .*?\]') {
                    if (-not $OOFFilesGroup.ContainsKey($OOFFile.FullName)) {
                        Write-Host '    Group specific OOF message'
                        $OOFFilesGroup.add($OOFFile.FullName, $OOFFileTargetName)
                    }
                    $NTName = ((($OOFFilePartTag -replace '\[', '') -replace '\]', '') -replace '(.*?) (.*)', '$1\$2')
                    if (-not $OOFFilesGroupSIDs.ContainsKey($OOFFilePartTag)) {
                        try {
                            $OOFFilesGroupSIDs.add($OOFFilePartTag, (New-Object System.Security.Principal.NTAccount($NTName)).Translate([System.Security.Principal.SecurityIdentifier]))
                        } catch {
                            # No group with this sAMAccountName found. Maybe it's a display name?
                            try {
                                $objTrans = New-Object -ComObject 'NameTranslate'
                                $objNT = $objTrans.GetType()
                                $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                                $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (4, ($NTName -split '\\')[1]))
                                $OOFFilesGroupSIDs.add($OOFFilePartTag, ((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                            } catch {
                            }
                        }
                    }

                    if ($OOFFilesGroupSIDs.containskey($OOFFilePartTag)) {
                        Write-Host "      $OOFFilePartTag = $NTName = $($OOFFilesGroupSIDs[$OOFFilePartTag])"
                        $OOFFilesGroupFilePart[$OOFFile.FullName] = ($OOFFilesGroupFilePart[$OOFFile.FullName] + '[' + $OOFFilesGroupSIDs[$OOFFilePartTag] + ']')
                    } else {
                        Write-Host "      $OOFFilePartTag = $($NTName): Not found in Active Directory, please check" -ForegroundColor Yellow
                    }
                } else {
                    Write-Host '    Common OOF message'
                    if (-not $OOFFilesCommon.containskey($OOFFile.FullName)) {
                        $OOFFilesCommon.add($OOFFile.FullName, $OOFFileTargetName)
                    }
                }
            }

            if ($OOFFilePart -match '(?i)\[Internal\]') {
                $OOFFilesInternal.add($OOFFile.FullName, $OOFFileTargetName)
                Write-Host '    Default OOF message for internal recipients'
            }

            if ($OOFFilePart -match '(?i)\[External\]') {
                $OOFFilesExternal.add($OOFFile.FullName, $OOFFileTargetName)
                Write-Host '    Default OOF message for external recipients'
            }
        }
    }


    Write-Host
    Write-Host "Start Word background process for template editing @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    try {
        $script:COMWord = New-Object -ComObject word.application
        if ($($PSVersionTable.PSEdition) -ieq 'Core') {
            Add-Type -Path (Get-ChildItem -LiteralPath ((Join-Path -Path ($env:SystemRoot) -ChildPath 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word')) -Filter 'Microsoft.Office.Interop.Word.dll' -Recurse | Select-Object -ExpandProperty FullName -Last 1)
        }
    } catch {
        Write-Host 'Word not installed or not working correctly. Exiting.' -ForegroundColor Red
        $error[0]
        exit 1
    }


    # Process each e-mail address only once
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
                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")
                try {
                    $Search.filter = "(distinguishedname=$($ADPropsCurrentMailbox.manager))"
                    $ADPropsCurrentMailboxManager = ([ADSI]"$(($Search.FindOne()).path)").Properties
                } catch {
                    $ADPropsCurrentMailboxManager = @()
                }

                $UserDomain = $ADPropsMailboxesUserDomain[$AccountNumberRunning]
                $SIDsToCheckInTrusts = @()
                $SIDsToCheckInTrusts += $ADPropsCurrentMailbox.objectsid
                try {
                    $UserAccount = [ADSI]"LDAP://$($ADPropsCurrentMailbox.distinguishedname)"
                    $UserAccount.GetInfoEx(@('tokengroups'), 0)
                    foreach ($sidBytes in $UserAccount.Properties.tokengroups) {
                        $sid = New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)
                        $GroupsSIDs += $sid.tostring()
                        Write-Host "      $sid"
                    }
                    $UserAccount.GetInfoEx(@('tokengroupsglobalanduniversal'), 0)
                    $SIDsToCheckInTrusts += $UserAccount.properties.tokengroupsglobalanduniversal
                } catch {
                    Write-Host "      Error getting group information from $((($ADPropsCurrentMailbox.distinguishedname) -split ',DC=')[1..999] -join '.'), check firewalls and AD trust" -ForegroundColor Red
                }
                # Loop through all domains to check if the mailbox account has a group membership there
                # Across a trust, a user can only be added to a domain local group.
                # Domain local groups can not be used outside their own domain, so we don't need to query recursively
                if ($SIDsToCheckInTrusts.count -gt 0) {
                    $LdapFilterSIDs = '(|'
                    $SIDsToCheckInTrusts | ForEach-Object {
                        try {
                            $SidHex = @()
                            $ot = New-Object System.Security.Principal.SecurityIdentifier($_, 0)
                            $c = New-Object 'byte[]' $ot.BinaryLength
                            $ot.GetBinaryForm($c, 0)
                            $c | ForEach-Object {
                                $SidHex += $('\{0:x2}' -f $_)
                            }
                            $LdapFilterSIDs += ('(objectsid=' + $($SidHex -join '') + ')')
                        } catch {
                        }
                    }
                    $LdapFilterSIDs += ')'
                } else {
                    $LdapFilterSIDs = ''
                }

                for ($DomainNumber = 0; $DomainNumber -lt $DomainsToCheckForGroups.count; $DomainNumber++) {
                    if (($DomainsToCheckForGroups[$DomainNumber] -ne '') -and ($DomainsToCheckForGroups[$DomainNumber] -ine $UserDomain) -and ($UserDomain -ne '')) {
                        Write-Host "    $($DomainsToCheckForGroups[$DomainNumber]) (mailbox group membership across trusts, takes some time) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($DomainsToCheckForGroups[$DomainNumber])")
                        $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                        foreach ($fsp in $Search.FindAll()) {
                            if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                                # Foreign Security Principals do not have the tokengroups attribute
                                # We need to switch to another, slower search method
                                # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member
                                # A Foreign Security Principal ist created in each (sub)domain, in which it is granted permissions,
                                # and it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal.
                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                $Search.filter = "(&(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                foreach ($group in $Search.findall()) {
                                    $sid = New-Object System.Security.Principal.SecurityIdentifier($group.properties.objectsid[0], 0)
                                    $GroupsSIDs += $sid.tostring()
                                    Write-Host "      $sid"
                                }
                            }
                        }
                    }
                }
            } else {
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor yellow
            }

            Write-Host "  SMTP addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $CurrentMailboxSMTPAddresses = @()
            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                $ADPropsCurrentMailbox.proxyaddresses | ForEach-Object {
                    if ([string]$_ -ilike 'smtp:*') {
                        $CurrentMailboxSMTPAddresses += [string]$_ -ireplace 'smtp:', ''
                        Write-Host ('    ' + ([string]$_ -ireplace 'smtp:', ''))
                    }
                }
            } else {
                $CurrentMailboxSMTPAddresses += $($MailAddresses[$AccountNumberRunning])
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor Yellow
                Write-Host '    Using mailbox name as single known SMTP address' -ForegroundColor Yellow
            }

            Write-Host "  Data for replacement variables @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $ReplaceHash = @{}
            if (Test-Path -Path $ReplacementVariableConfigFile -PathType Leaf) {
                try {
                    Write-Host "    Executing content of config file '$ReplacementVariableConfigFile'"
                    . ([System.Management.Automation.ScriptBlock]::Create((Get-Content -LiteralPath $ReplacementVariableConfigFile -Raw)))
                } catch {
                    Write-Host "    Problem executing content of '$ReplacementVariableConfigFile'. Exiting." -ForegroundColor Red
                    Write-Host "    Error: $_" -ForegroundColor Red
                    $error[0]
                    exit 1
                }
            } else {
                Write-Host "    Problem connecting to or reading from file '$ReplacementVariableConfigFile'. Exiting." -ForegroundColor Red
                exit 1
            }
            foreach ($replaceKey in ($replaceHash.Keys | Sort-Object)) {
                if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                    if ($($replaceHash[$replaceKey])) {
                        Write-Host "    $($replaceKey): $($replaceHash[$replaceKey])"
                    }
                } else {
                    if ($null -ne $($replaceHash[$replaceKey])) {
                        Write-Host "    $($replaceKey): Photo available"
                    }
                }
            }

            # Export pictures if available
            ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$') | ForEach-Object {
                if ($null -ne $ReplaceHash[$_]) {
                    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        $ReplaceHash[$_] | Set-Content -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg')))) -AsByteStream -Force
                    } else {
                        $ReplaceHash[$_] | Set-Content -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg')))) -Encoding Byte -Force
                    }
                }
            }

            Write-Host "  Process common signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            if ($SignatureFilesCommon.count -gt 0) {
                foreach ($Signature in ($SignatureFilesCommon.GetEnumerator() | Sort-Object -Property Name)) {
                    Set-Signatures
                }
            } else {
                Write-Host '    Found no common signatures for this mailbox.'
            }


            Write-Host "  Process group specific signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $SignatureHash = @{}
            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                foreach ($x in ($SignatureFilesGroupFilePart.GetEnumerator() | Sort-Object -Property Name)) {
                    $GroupsSIDs | ForEach-Object {
                        if ($x.Value.tolower().Contains('[' + $_.tolower() + ']')) {
                            $SignatureHash.add($x.Name, $SignatureFilesGroup[$x.Name])
                        }
                    }
                }
                if ($SignatureHash.count -gt 0) {
                    foreach ($Signature in ($SignatureHash.GetEnumerator() | Sort-Object -Property Name)) {
                        Set-Signatures
                    }
                } else {
                    Write-Host '    Found no group specific signatures for this mailbox.'
                }
            } else {
                $CurrentMailboxSMTPAddresses += $($MailAddresses[$AccountNumberRunning])
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor Yellow
            }

            Write-Host "  Process e-mail address specific signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
            $SignatureHash = @{}
            foreach ($x in ($SignatureFilesMailboxFilePart.GetEnumerator() | Sort-Object -Property Name)) {
                foreach ($y in $CurrentMailboxSMTPAddresses) {
                    if ($x.Value.tolower().contains('[' + $y.tolower() + ']')) {
                        $SignatureHash.add($x.Name, $SignatureFilesMailbox[$x.Name])
                    }
                }
            }
            if ($SignatureHash.count -gt 0) {
                foreach ($Signature in ($SignatureHash.GetEnumerator() | Sort-Object -Property Name)) {
                    Set-Signatures
                }
            } else {
                Write-Host '    Found no e-mail address specific signatures for this mailbox.'
            }

            # Delete photos from file system
            ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$') | ForEach-Object {
                Remove-Item -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg')))) -Force -ErrorAction SilentlyContinue
                $ReplaceHash.Remove($_)
                $ReplaceHash.Remove(($_[-999..-2] -join '') + 'DELETEEMPTY$')
            }

        }

        # Set OOF message and Outlook Web signature
        if ((($SetCurrentUserOutlookWebSignature -eq $true) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress)) {
            if (-not $SimulateUser) {
                try {
                    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\Microsoft.Exchange.WebServices.NETStandard.dll')) -Destination ((Join-Path -Path $script:tempDir -ChildPath 'Microsoft.Exchange.WebServices.NETStandard.dll')) -Force -ErrorAction SilentlyContinue
                    } else {
                        Copy-Item -Path ((Join-Path -Path '.' -ChildPath 'bin\Microsoft.Exchange.WebServices.dll')) -Destination ((Join-Path -Path $script:tempDir -ChildPath 'Microsoft.Exchange.WebServices.dll')) -Force
                    }
                } catch {
                }

                $error.clear()

                try {
                    if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                        Import-Module -Name ((Join-Path -Path $script:tempDir -ChildPath 'Microsoft.Exchange.WebServices.NETStandard.dll')) -Force
                    } else {
                        Import-Module -Name ((Join-Path -Path $script:tempDir -ChildPath 'Microsoft.Exchange.WebServices.dll')) -Force
                    }

                    $exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
                    $exchService.UseDefaultCredentials = $true
                    $exchService.AutodiscoverUrl($PrimaryMailboxAddress) | Out-Null
                } catch {
                    Write-Host "  Error connecting to Outlook Web: $_" -ForegroundColor Red

                    if ($SetCurrentUserOutlookWebSignature) {
                        Write-Host '  Outlook Web signature can not be set' -ForegroundColor Red
                    }

                    if ($SetCurrentUserOOFMessage) {
                        Write-Host '  Out of Office (OOF) auto reply message(s) can not be set' -ForegroundColor Red
                    }
                }
            }
            if ((!$error -and (-not $SimulateUser)) -or ($SimulateUser)) {
                if ($SetCurrentUserOutlookWebSignature) {
                    Write-Host "  Set Outlook Web signature @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    if ($SimulateUser) {
                        Write-Host '    Simulation mode enabled, skipping task' -ForegroundColor Yellow
                    } else {
                        # If this is the primary mailbox, set OWA signature
                        for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                            if ($MailAddresses[$j] -ieq $PrimaryMailboxAddress) {
                                if ($RegistryPaths[$j] -like ('*\Outlook\Profiles\' + $OutlookDefaultProfile + '\9375CFF0413111d3B88A00104B2A6676\*')) {
                                    try {
                                        if ($script:CurrentUserDummyMailbox -ne $true) {
                                            $TempNewSig = Get-ItemPropertyValue -LiteralPath $RegistryPaths[$j] -Name 'New Signature'
                                        } else {
                                            $TempNewSig = $script:CurrentUserDummyMailboxDefaultSigNew
                                        }
                                    } catch {
                                        $TempNewSig = ''
                                    }
                                    try {
                                        if ($script:CurrentUserDummyMailbox -ne $true) {
                                            $TempReplySig = Get-ItemPropertyValue -LiteralPath $RegistryPaths[$j] -Name 'Reply-Forward Signature'
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
                                        Write-Host "    Only default signature for new mails is set: '$TempNewSig'"
                                        $TempOWASigFile = $TempNewSig
                                        $TempOWASigSetNew = 'True'
                                        $TempOWASigSetReply = 'False'
                                    }

                                    if (($TempNewSig -eq '') -and ($TempReplySig -ne '')) {
                                        Write-Host "    Only default signature for reply/forward is set: '$TempReplySig'"
                                        $TempOWASigFile = $TempReplySig
                                        $TempOWASigSetNew = 'False'
                                        $TempOWASigSetReply = 'True'
                                    }


                                    if ((($TempNewSig -ne '') -and ($TempReplySig -ne '')) -and ($TempNewSig -ine $TempReplySig)) {
                                        Write-Host "    Different default signatures for new and reply/forward set, using new one: '$TempNewSig'"
                                        $TempOWASigFile = $TempNewSig
                                        $TempOWASigSetNew = 'True'
                                        $TempOWASigSetReply = 'False'
                                    }

                                    if ((($TempNewSig -ne '') -and ($TempReplySig -ne '')) -and ($TempNewSig -ieq $TempReplySig)) {
                                        Write-Host "    Same default signature for new and reply/forward: '$TempNewSig'"
                                        $TempOWASigFile = $TempNewSig
                                        $TempOWASigSetNew = 'True'
                                        $TempOWASigSetReply = 'True'
                                    }
                                    if (($null -ne $TempOWASigFile) -and ($TempOWASigFile -ne '')) {

                                        try {
                                            if (Test-Path -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.htm'))) -PathType Leaf) {
                                                $hsHtmlSignature = (Get-Content -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.htm'))) -Raw).ToString()
                                            } else {
                                                $hsHtmlSignature = ''
                                                Write-Host "      Signature file '$($TempOWASigFile + '.htm')' not found. Outlook Web HTML signature will be blank." -ForegroundColor Yellow
                                            }
                                            if (Test-Path -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.txt'))) -PathType Leaf) {
                                                $stTextSig = (Get-Content -LiteralPath ((Join-Path -Path ($SignaturePaths[0]) -ChildPath ($TempOWASigFile + '.txt'))) -Raw).ToString()
                                            } else {
                                                $hsHtmlSignature = ''
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
                                            if ($($PSVersionTable.PSEdition) -ieq 'Core') { $UsrConfig = $UsrConfig.result }

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
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if ($SetCurrentUserOOFMessage) {
                    $OOFCommonGUID = (New-Guid).guid
                    $OOFInternalGUID = (New-Guid).guid
                    $OOFExternalGUID = (New-Guid).guid

                    Write-Host "  Process Out of Office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
                    $OOFDisabled = $null
                    if ($SimulateUser) {
                        Write-Host '    Simulation mode enabled, processing OOF templates without changing OOF settings' -ForegroundColor Yellow
                    } else {
                        $OOFSettings = $exchService.GetUserOOFSettings($PrimaryMailboxAddress)
                        if ($($PSVersionTable.PSEdition) -ieq 'Core') { $OOFSettings = $OOFSettings.result }
                        if ($OOFSettings.STATE -eq [Microsoft.Exchange.WebServices.Data.OOFState]::Disabled) { $OOFDisabled = $true }
                    }

                    if (($OOFDisabled -and (-not $SimulateUser)) -or ($SimulateUser)) {
                        # First, loop through common OOF files
                        foreach ($OOF in ($OOFFilesCommon.GetEnumerator() | Sort-Object -Property Name)) {
                            if (($OOFFilesInternal.contains('' + $OOF.name + '')) -or (-not ($OOFFilesExternal.contains('' + $OOF.name + '')))) {
                                $OOFInternal = $OOF.name
                            }
                            if (($OOFFilesExternal.contains('' + $OOF.name + '')) -or (-not ($OOFFilesInternal.contains('' + $OOF.name + '')))) {
                                $OOFExternal = $OOF.name
                            }
                        }
                        # Second, loop through group OOF files
                        if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                            foreach ($x in ($OOFFilesGroupFilePart.GetEnumerator() | Sort-Object -Property Name)) {
                                $GroupsSIDs | ForEach-Object {
                                    if ($x.Value.tolower().Contains('[' + $_.tolower() + ']')) {
                                        if (($OOFFilesInternal.contains('' + $x.name + '')) -or (-not ($OOFFilesExternal.contains('' + $x.name + '')))) {
                                            $OOFInternal = $x.name
                                        }
                                        if (($OOFFilesExternal.contains('' + $x.name + '')) -or (-not ($OOFFilesInternal.contains('' + $x.name + '')))) {
                                            $OOFExternal = $x.name
                                        }
                                    }
                                }
                            }
                        } else {
                            $CurrentMailboxSMTPAddresses += $($MailAddresses[$AccountNumberRunning])
                            Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox' -ForegroundColor Yellow
                        }
                        # Third, loop through e-mail address specific OOF files
                        foreach ($x in ($OOFFilesMailboxFilePart.GetEnumerator() | Sort-Object -Property Name)) {
                            foreach ($y in ($CurrentMailboxSMTPAddresses | Sort-Object -Property Name)) {
                                if ($x.Value.tolower().contains('[' + $y.tolower() + ']')) {
                                    if (($OOFFilesInternal.contains('' + $x.name + '')) -or (-not ($OOFFilesExternal.contains('' + $x.name + '')))) {
                                        $OOFInternal = $x.name
                                    }
                                    if (($OOFFilesExternal.contains('' + $x.name + '')) -or (-not ($OOFFilesInternal.contains('' + $x.name + '')))) {
                                        $OOFExternal = $x.name
                                    }
                                }
                            }
                        }

                        $SignatureHash = @{}
                        if ($OOFInternal -ine $OOFExternal) {
                            Write-Host "    Message template for internal recpients: '$OOFInternal'"
                            if ($UseHtmTemplates) {
                                $SignatureHash.add($OOFInternal, "$OOFInternalGUID OOFInternal.htm")
                            } else {
                                $SignatureHash.add($OOFInternal, "$OOFInternalGUID OOFInternal.docx")
                            }
                            Write-Host "    Message template for external recpients: '$OOFExternal'"
                            if ($UseHtmTemplates) {
                                $SignatureHash.add($OOFExternal, "$OOFExternalGUID OOFExternal.htm")
                            } else {
                                $SignatureHash.add($OOFExternal, "$OOFExternalGUID OOFExternal.docx")
                            }
                        } else {
                            Write-Host "    Common template for internal and external recpients: '$OOFInternal'"
                            if (($null -ne $OOFInternal) -and ($OOFInternal -ne '')) {
                                if ($UseHtmTemplates) {
                                    $SignatureHash.add($OOFInternal, "$OOFCommonGUID OOFCommon.htm")
                                } else {
                                    $SignatureHash.add($OOFInternal, "$OOFCommonGUID OOFCommon.docx")
                                }
                            }
                        }
                        foreach ($Signature in ($SignatureHash.GetEnumerator() | Sort-Object -Property Name)) {
                            Set-Signatures -ProcessOOF
                        }

                        if (Test-Path -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFCommonGUID OOFCommon.htm"))) {
                            if (-not $SimulateUser) {
                                $OOFSettings.InternalReply = New-Object Microsoft.Exchange.WebServices.Data.OOFReply((Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFCommonGUID OOFCommon.htm")) -Raw).ToString())
                                $OOFSettings.ExternalReply = New-Object Microsoft.Exchange.WebServices.Data.OOFReply((Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFCommonGUID OOFCommon.htm")) -Raw).ToString())
                            } else {
                                $SignaturePaths | ForEach-Object {
                                    Copy-Item -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFCommonGUID OOFCommon.htm")) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($_) -ChildPath 'OOF\') -Force).fullname) -ChildPath 'OOFInternal.htm')) -Force
                                    Copy-Item -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFCommonGUID OOFCommon.htm")) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($_) -ChildPath 'OOF\') -Force).fullname) -ChildPath 'OOFExternal.htm')) -Force
                                }
                            }
                        } else {
                            if (-not $SimulateUser) {
                                $OOFSettings.InternalReply = New-Object Microsoft.Exchange.WebServices.Data.OOFReply((Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) -Raw).ToString())
                                $OOFSettings.ExternalReply = New-Object Microsoft.Exchange.WebServices.Data.OOFReply((Get-Content -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) -Raw).ToString())
                            } else {
                                $SignaturePaths | ForEach-Object {
                                    Copy-Item -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm")) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($_) -ChildPath 'OOF\') -Force).fullname) -ChildPath 'OOFInternal.htm')) -Force
                                    Copy-Item -LiteralPath ((Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")) -Destination ((Join-Path -Path ((New-Item -ItemType Directory (Join-Path -Path ($_) -ChildPath 'OOF\') -Force).fullname) -ChildPath 'OOFExternal.htm')) -Force
                                }
                            }
                        }
                        if (-not $SimulateUser) {
                            try {
                                $exchService.SetUserOOFSettings($PrimaryMailboxAddress, $OOFSettings) | Out-Null
                            } catch {
                                Write-Host '    Error setting Outlook Web Out of Office (OOF) auto reply message(s)' -ForegroundColor Red
                            }
                        }
                    } else {
                        Write-Host '    Out of Office (OOF) auto reply currently active or scheduled, not changing settings' -ForegroundColor Yellow
                    }

                    # Delete temporary OOF files from file system
                    ("$OOFCommonGUID OOFCommon", "$OOFInternalGUID OOFInternal", "$OOFExternalGUID OOFExternal") | ForEach-Object {
                        Remove-Item ((Join-Path -Path $script:tempDir -ChildPath ($_ + '.*'))) -Force -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }


    # Delete old signatures created by this script, which are no longer available in $SignatureTemplatePath
    # We check all local signatures for a specific marker in HTML code, so we don't touch user created signatures
    Write-Host
    Write-Host "Remove old signatures created by this script, which are no longer centrally available @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $SignaturePaths | ForEach-Object {
        Get-ChildItem -LiteralPath $_ -Filter '*.htm' -File | ForEach-Object {
            if ((Get-Content -LiteralPath $_.fullname -Raw) -like ('*' + $HTMLMarkerTag + '*')) {
                if ($_.name -notin $script:SignatureFilesDone) {
                    Write-Host ("  '" + $([System.IO.Path]::ChangeExtension($_.fullname, '')) + "*'")
                    Remove-Item -LiteralPath $_.fullname -Force -ErrorAction silentlycontinue
                    Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($_.fullname, '.rtf'))) -Force -ErrorAction silentlycontinue
                    Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($_.fullname, '.txt'))) -Force -ErrorAction silentlycontinue
                }
            }
        }
    }


    # Delete user created signatures if $DeleteUserCreatedSignatures -eq $true
    if ($DeleteUserCreatedSignatures -eq $true) {
        Write-Host
        Write-Host "Remove user created signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
        $SignaturePaths | ForEach-Object {
            Get-ChildItem -LiteralPath $_ -Filter '*.htm' -File | ForEach-Object {
                if ((Get-Content -LiteralPath $_.fullname -Raw) -notlike ('*' + $HTMLMarkerTag + '*')) {
                    Write-Host ("  '" + $([System.IO.Path]::ChangeExtension($_.fullname, '')) + "*'")
                    Remove-Item -LiteralPath $_.fullname -Force -ErrorAction silentlycontinue
                    Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($_.fullname, '.rtf'))) -Force -ErrorAction silentlycontinue
                    Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($_.fullname, '.txt'))) -Force -ErrorAction silentlycontinue
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
            Write-Host '    Simulation mode enabled, skipping task' -ForegroundColor Yellow
        } else {
            if (-not (Test-Path $AdditionalSignaturePath -PathType Container -ErrorAction SilentlyContinue)) {
                New-Item -Path $AdditionalSignaturePath -ItemType Directory -Force | Out-Null
                if (-not (Test-Path $AdditionalSignaturePath -PathType Container -ErrorAction SilentlyContinue)) {
                    Write-Host '  Path could not be accessed or created, ignoring path.' -ForegroundColor Yellow
                } else {
                    Copy-Item -Path (Join-Path -Path ($(if ($script:CurrentUserDummyMailbox -eq $true) { $SignaturePaths[1] } else { $SignaturePaths[0] })) -ChildPath '*') -Destination $AdditionalSignaturePath -Recurse -Force -ErrorAction SilentlyContinue
                }
            } else {
                (Get-ChildItem -Path $AdditionalSignaturePath -Recurse -Force).fullname | Remove-Item -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                Copy-Item -Path (Join-Path -Path ($(if ($script:CurrentUserDummyMailbox -eq $true) { $SignaturePaths[1] } else { $SignaturePaths[0] })) -ChildPath '*') -Destination $AdditionalSignaturePath -Recurse -Force
            }
        }
    }

    if ($script:CurrentUserDummyMailbox -eq $true) {
        Remove-Item $SignaturePaths[0] -Recurse -Force
    }
}


Function ConvertTo-SingleFileHTML([string]$inputfile, [string]$outputfile) {
    $tempFileContent = Get-Content -LiteralPath $inputfile -Raw -Encoding UTF8

    $src = @()
    ([regex]'(?i)src="(.*?)"').Matches($tempFileContent) | ForEach-Object {
        $src += $_.Groups[0].Value
        if ($_.Groups[0].Value.StartsWith('src="data:')) {
            $src += ''
        } else {
            $src += (Join-Path -Path (Split-Path -Path ($inputfile) -Parent) -ChildPath ([uri]::UnEscapeDataString($_.Groups[1].Value)))
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

    $tempFileContent | Out-File -LiteralPath $outputfile -Encoding UTF8 -Force
}


function Set-Signatures {
    Param(
        [switch]$ProcessOOF = $false
    )

    Write-Host "    '$($Signature.Name)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if (-not $ProcessOOF) {
        Write-Host "      Outlook signature name: '$([System.IO.Path]::ChangeExtension($($Signature.value), $null) -replace '.$')'"
    }

    if (-not $ProcessOOF) {
        $SignatureFileAlreadyDone = ($script:SignatureFilesDone -contains $($Signature.Name))
        if ($SignatureFileAlreadyDone) {
            Write-Host '      Template already processed before (mailbox or signature group with higher priority), skipping' -ForegroundColor Yellow
        } else {
            if ($script:CurrentUserDummyMailbox -ne $true) {
                $script:SignatureFilesDone += $($Signature.Name)
            }
        }
    }
    if (($SignatureFileAlreadyDone -eq $false) -or $ProcessOOF) {
        Write-Host '      Create temporary file copy'

        if ($UseHtmTemplates) {
            # use .html for temporary file, .htm for final file
            $path = ($(Join-Path -Path $script:tempDir -ChildPath (New-Guid).guid).tostring() + '.htm')
            #try {
            ConvertTo-SingleFileHTML $Signature.Name $path
            #} catch {
            #Write-Host '        Error copying file. Skipping signature.' -ForegroundColor Red
            #continue
            #}
        } else {
            $path = ($(Join-Path -Path $script:tempDir -ChildPath (New-Guid).guid).tostring() + '.docx')
            try {
                Copy-Item -LiteralPath $Signature.Name -Destination $path -Force
            } catch {
                Write-Host '        Error copying file. Skipping signature.' -ForegroundColor Red
                continue
            }
        }

        $Signature.value = $([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))
        if (-not $ProcessOOF) {
            if ($script:CurrentUserDummyMailbox -ne $true) {
                $script:SignatureFilesDone += $Signature.Value
            }
        }

        if ($UseHtmTemplates) {
            Write-Host '      Replace picture variables'
            $html = New-Object -ComObject 'HTMLFile'
            $HTML.IHTMLDocument2_write((Get-Content -LiteralPath $path -Raw -Encoding UTF8))

            foreach ($image in ($html.images)) {
                ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$') | ForEach-Object {
                    if (($image.src -clike "*$_*") -or ($image.alt -clike "*$_*")) {
                        if ($null -ne $ReplaceHash[$_]) {
                            $ImageAlternativeTextOriginal = $image.alt
                            $image.src = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg'))))))
                            $image.alt = $ImageAlternativeTextOriginal.replace($_, '')
                        }
                    } elseif (($image.src -clike "*$(($_[-999..-2] -join '') + 'DELETEEMPTY$')*") -or ($image.alt -clike "*$(($_[-999..-2] -join '') + 'DELETEEMPTY$')*")) {
                        if ($null -ne $ReplaceHash[$_]) {
                            $ImageAlternativeTextOriginal = $image.alt
                            $image.src = ('data:image/jpeg;base64,' + [Convert]::ToBase64String([IO.File]::ReadAllBytes(((Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg'))))))
                            $image.alt = $ImageAlternativeTextOriginal.replace((($_[-999..-2] -join '') + 'DELETEEMPTY$'), '')
                        } else {
                            $image.removenode() | Out-Null
                        }
                    }
                }
            }

            Write-Host '      Replace non-picture variables'
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

        $script:COMWord.Documents.Open($path, $false) | Out-Null

        if (-not $UseHtmTemplates) {
            Write-Host '      Replace picture variables'
            foreach ($image in ($script:COMWord.ActiveDocument.Shapes + $script:COMWord.ActiveDocument.InlineShapes)) {
                try {
                    if (($null -ne $($image.linkformat.sourcefullname)) -and ($($image.linkformat.sourcefullname) -ne '')) {
                        ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$') | ForEach-Object {
                            if ((((Split-Path -Path ($($image.linkformat.sourcefullname)) -Leaf).contains($_)) -or (($image.alternativetext).contains($_)))) {
                                if ($null -ne $ReplaceHash[$_]) {
                                    $ImageAlternativeTextOriginal = $image.AlternativeText
                                    $image.linkformat.sourcefullname = (Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg'))
                                    $image.alternativetext = $ImageAlternativeTextOriginal.replace($_, '')
                                }
                            } elseif (((Split-Path -Path ($($image.linkformat.sourcefullname)) -Leaf).contains(($_[-999..-2] -join '') + 'DELETEEMPTY$')) -or ($image.alternativetext.contains(($_[-999..-2] -join '') + 'DELETEEMPTY$'))) {
                                if ($null -ne $ReplaceHash[$_]) {
                                    $ImageAlternativeTextOriginal = $image.AlternativeText
                                    $image.linkformat.sourcefullname = (Join-Path -Path $script:tempDir -ChildPath ($_ + '.jpeg'))
                                    $image.alternativetext = $ImageAlternativeTextOriginal.replace((($_[-999..-2] -join '') + 'DELETEEMPTY$'), '')
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
                $tempImageHyperlinkName = $image.hyperlink.Name
                $tempImageHyperlinkAddress = $image.hyperlink.Address
                $tempImageHyperlinkAddressOld = $image.hyperlink.AddressOld
                $tempImageHyperlinkSubAddress = $image.hyperlink.SubAddress
                $tempImageHyperlinkSubaddressOld = $image.hyperlink.SubAddressOld
                $tempImageHyperlinkEmailSubject = $image.hyperlink.EmailSubject
                $tempImageHyperlinkScreenTip = $image.hyperlink.ScreenTip

                foreach ($replaceKey in $replaceHash.Keys) {
                    if ($replaceKey -notin ('$CURRENTMAILBOXMANAGERPHOTO$', '$CURRENTMAILBOXPHOTO$', '$CURRENTUSERMANAGERPHOTO$', '$CURRENTUSERPHOTO$', '$CURRENTMAILBOXMANAGERPHOTODELETEEMPTY$', '$CURRENTMAILBOXPHOTODELETEEMPTY$', '$CURRENTUSERMANAGERPHOTODELETEEMPTY$', '$CURRENTUSERPHOTODELETEEMPTY$')) {
                        if ($null -ne $tempimagealternativetext) {
                            $tempimagealternativetext = $tempimagealternativetext.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkName) {
                            $tempimagehyperlinkname = $tempimagehyperlinkname.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkAddress) {
                            $tempimagehyperlinkAddress = $tempimagehyperlinkAddress.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkAddressOld) {
                            $tempimagehyperlinkAddressOld = $tempimagehyperlinkAddressOld.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkSubAddress) {
                            $tempimagehyperlinkSubAddress = $tempimagehyperlinkSubAddress.replace($replaceKey, $replaceHash.replaceKey)
                        }
                        if ($null -ne $tempimagehyperlinkSubAddressOld) {
                            $tempimagehyperlinkSubAddressOld = $tempimagehyperlinkSubAddressOld.replace($replaceKey, $replaceHash.replaceKey)
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
                if ($null -ne $tempimagehyperlinkName) {
                    $image.hyperlink.Name = $tempImageHyperlinkName
                }
                if ($null -ne $tempimagehyperlinkAddress) {
                    $image.hyperlink.Address = $tempImageHyperlinkAddress
                }
                if ($null -ne $tempimagehyperlinkAddressOld) {
                    $image.hyperlink.AddressOld = $tempImageHyperlinkAddressOld
                }
                if ($null -ne $tempimagehyperlinkSubAddress) {
                    $image.hyperlink.SubAddress = $tempImageHyperlinkSubAddress
                }
                if ($null -ne $tempimagehyperlinkSubAddressOld) {
                    $image.hyperlink.SubAddressOld = $tempImageHyperlinkSubaddressOld
                }
                if ($null -ne $tempimagehyperlinkEmailSubject) {
                    $image.hyperlink.EmailSubject = $tempImageHyperlinkEmailSubject
                }
                if ($null -ne $tempimagehyperlinkScreenTip) {
                    $image.hyperlink.ScreenTip = $tempImageHyperlinkScreenTip
                }
            }

            Write-Host '      Replace non-picture variables'
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
                    $ReplaceWith = $replaceHash.$replaceKey
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
                    $ReplaceWith = $replaceHash.$replaceKey
                    $script:COMWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, `
                            $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                            $Wrap, $Format, $ReplaceWith, $ReplaceAll) | Out-Null
                }
            }

            # Restore original view
            $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = (-not $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes)

            # Exports
            Write-Host '      Export to HTM format'
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatFilteredHTML')
            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
            $script:COMWord.ActiveDocument.Weboptions.encoding = 65001
            $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
        }

        if (-not $ProcessOOF) {
            Write-Host '      Export to RTF format'
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatRTF')
            $path = $([System.IO.Path]::ChangeExtension($path, '.rtf'))
            $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
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
            $script:COMWord.ActiveDocument.Save()
            $script:COMWord.ActiveDocument.Close($false)


            Write-Host '      Export to TXT format'
            # We work with the .htm file to avoid problems with empty lines at the end of exported .txt files. Details: https://eileenslounge.com/viewtopic.php?t=16703
            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
            $script:COMWord.Documents.Open($path, $false) | Out-Null
            $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatUnicodeText')
            $script:COMWord.ActiveDocument.TextEncoding = 1200
            $path = $([System.IO.Path]::ChangeExtension($path, '.txt'))
            $script:COMWord.ActiveDocument.SaveAs($path, $saveFormat)
            $script:COMWord.ActiveDocument.Close($false)
        } else {
            $script:COMWord.ActiveDocument.Close($false)
        }

        Write-Host '      Embed local files in HTM format and add marker'
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

        $tempFileContent = Get-Content -LiteralPath $path -Raw -Encoding UTF8

        if ($tempFileContent -notlike "*$HTMLMarkerTag*") {
            if ($tempFileContent -like '*<head>*') {
                $tempFileContent = $tempFileContent -ireplace ('<HEAD>', ('<HEAD>' + $HTMLMarkerTag))
            } else {
                $tempFileContent = $tempFileContent -ireplace ('<HTML>', ('<HTML><HEAD>' + $HTMLMarkerTag + '</HEAD>'))
            }
        }

        $tempFileContent | Out-File -LiteralPath $path -Encoding UTF8 -Force

        if (-not $ProcessOOF) {
            ConvertTo-SingleFileHTML $path $path
        } else {
            ConvertTo-SingleFileHTML $path ((Join-Path -Path $script:tempDir -ChildPath $Signature.Value)) -Encoding UTF8 -Force
        }


        if (-not $ProcessOOF) {
            $(if ($script:CurrentUserDummyMailbox -eq $true) { $SignaturePaths[0] } else { $SignaturePaths }) | ForEach-Object {
                Write-Host "      Copy signature files to '$_'"
                Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.htm')) -Destination ((Join-Path -Path ($_) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))) -Force
                Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.rtf')) -Destination ((Join-Path -Path ($_) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))) -Force
                Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.txt')) -Destination ((Join-Path -Path ($_) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))) -Force
            }
        }

        Write-Host '      Remove temporary files'
        Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.docx')) -Force -Recurse -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.htm')) -Force -Recurse -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.rtf')) -Force -Recurse -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.txt')) -Force -Recurse -ErrorAction SilentlyContinue
        Foreach ($x in (Get-ChildItem -Path ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($path) + '*') -Directory).FullName) {
            Remove-Item -LiteralPath $x -Force -Recurse -ErrorAction SilentlyContinue
        }
    }

    if ((-not $ProcessOOF) -and (-not $SimulateUser)) {
        # Set default signature for new mails
        if ($SignatureFilesDefaultNew.contains('' + $Signature.name + '')) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    Write-Host '      Set signature as default for new messages'
                    if ($script:CurrentUserDummyMailbox -ne $true) {
                        Set-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -Type String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force
                    } else {
                        $script:CurrentUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }
                }
            }
        }

        # Set default signature for replies and forwarded mails
        if ($SignatureFilesDefaultReplyFwd.contains($Signature.name)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    Write-Host '      Set signature as default for reply/forward messages'
                    if ($script:CurrentUserDummyMailbox -ne $true) {
                        Set-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -Type String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force
                    } else {
                        $script:CurrentUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
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
    $PowerShell = [powershell]::Create()
    $PowerShell.RunspacePool = $RunspacePool
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
        $script:jobs | ForEach-Object {
            if (($null -eq $_.StartTime) -and ($_.Powershell.Streams.Debug[0].Message -match 'Start')) {
                $StartTicks = $_.powershell.Streams.Debug[0].Message -replace '[^0-9]'
                $_.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $_.StartTime) {
                if ((($_.handle.IsCompleted -eq $true) -and ($_.Done -eq $false)) -or (($_.Done -eq $false) -and ((New-TimeSpan -Start $_.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $_.Object[0..$(($_.object).count - 1)]
                    Write-Host "$Indent$($data[0])"
                    if ($data -icontains 'QueryPassed') {
                        Write-Host "$Indent  $CheckProtocolText query successful."
                        $returnvalue = $true
                    } else {
                        Write-Host "$Indent  $CheckProtocolText query failed, removing domain from list." -ForegroundColor Red
                        Write-Host "$Indent  If this error is permanent, check firewalls and AD trust. Consider using parameter DomainsToCheckForGroups." -ForegroundColor Red
                        $DomainsToCheckForGroups.remove($data[0])
                        $returnvalue = $false
                    }
                    $_.Done = $true
                }
            }
        }
    }
    return $returnvalue
}


function CheckPath([string]$path) {
    if ($path.StartsWith('https://', 'CurrentCultureIgnoreCase')) {
        $path = ((([uri]::UnescapeDataString($path) -ireplace ('https://', '\\')) -replace ('(.*?)/(.*)', '${1}@SSL\$2')) -replace ('/', '\'))
    } else {
        $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
    }

    if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) {
        # Reconnect already connected network drives at the OS level
        # New-PSDrive is not enough for this
        Get-CimInstance Win32_NetworkConnection | ForEach-Object {
            & net use $_.LocalName $_.RemoteName 2>&1 | Out-Null
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
                    $app.windows() | Where-Object { $_.LocationURL -like ('*' + ([uri]::EscapeUriString(((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') } | ForEach-Object {
                        $i = $i + 1
                    }
                    Start-Sleep -Milliseconds 50
                }

                # Wait until the corresponding URL is fully loaded, then close the tab
                $app.windows() | Where-Object { $_.LocationURL -like ('*' + ([uri]::EscapeUriString(((($path -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') } | ForEach-Object {
                    while ($_.busy) {
                        Start-Sleep -Milliseconds 50
                    }
                    $_.quit()
                }

                $app = $null
            } catch {
            }
        }
    }

    if ((Test-Path -LiteralPath $path) -eq $false) {
        Write-Host ": Problem connecting to or reading from folder '$path'. Exiting." -ForegroundColor Red
        exit 1
    } else {
        Write-Host
    }
}


try {
    Clear-Host

    Write-Host "Script started @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    main
} catch {
    Write-Host
    Write-Host 'Unexpected error. Exiting.' -ForegroundColor red
    $Error[0]
    exit 1
} finally {
    Write-Host
    Write-Host "Cleanup tasks @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    if ($script:COMWord) {
        $script:COMWord.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null
        Remove-Variable COMWord
    }

    Remove-Module -Name Microsoft.Exchange.WebServices -Force -ErrorAction SilentlyContinue
    Remove-Item ((Join-Path -Path $script:tempDir -ChildPath 'Microsoft.Exchange.WebServices.dll')) -Force -ErrorAction SilentlyContinue
    Remove-Item ((Join-Path -Path $script:tempDir -ChildPath 'Microsoft.Exchange.WebServices.NETStandard.dll')) -Force -ErrorAction SilentlyContinue


    Write-Host
    Write-Host "Script ended @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
}
