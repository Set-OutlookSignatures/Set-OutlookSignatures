<#
.SYNOPSIS
Set-OutlookSignatures XXXVersionStringXXX
Email signatures and out-of-office replies for Exchange and all of Outlook.
Full-featured, cost-effective, unsurpassed data privacy.

.DESCRIPTION
Find the full documentation at https://set-outlooksignatures.com.

.LINK
Web: https://set-outlooksignatures.com
Benefactor Circle add-on: https://set-outlooksignatures.com/benefactorcircle
Support: https://set-outlooksignatures.com/support

.EXAMPLE
Find the full documentation at https://set-outlooksignatures.com.

.NOTES
Software: Set-OutlookSignatures
Version : XXXVersionStringXXX
Web     : https://set-outlooksignatures.com
License : See '.\LICENSE.txt' for details and copyright
#>


# Suppress specific PSScriptAnalyzerRules for specific variables
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'SimulateAndDeployGraphCredentialFile')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ADPropsCurrentMailboxManager')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentAutodiscoverSecureName')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentAzureADEndpoint')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentEnvironmentName')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentGraphApiEndpoint')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironmentSharePointOnlineDomains')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ConnectedFilesFolderNames')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CurrentTemplateisForAliasSmtp')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'data')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'GraphClientID')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'GraphClientIDOriginal')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFExternalValueBasename')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFFilesExternal')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFFilesInternal')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'OOFInternalValueBasename')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'pathHtmlFolderSuffix')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'PrimaryMailboxAddress')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ScriptInvocation')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'ScriptVersion')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'SignatureFilesDefaultNew')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'SignatureFilesDefaultReplyFwd')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'SignatureFilesWriteProtect')]


[CmdletBinding(PositionalBinding = $false, DefaultParameterSetName = 'Z: All parameters')]


param(
    # Path to a Benefactor Circle license file
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$BenefactorCircleLicenseFile = '',

    # The Benefactor Circle Member ID matching the Benefactor Circle license file
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$BenefactorCircleID = '',

    # Use templates in .HTM file format instead of .DOCX
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $UseHtmTemplates = $(if ($IsWindows -or (-not (Test-Path -LiteralPath 'variable:IsWindows'))) { $false } else { $true }),

    # Path to centrally managed signature templates
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$SignatureTemplatePath = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Signatures DOCX' } else { '.\sample templates\Signatures HTML' }),

    # Path to INI file containing signature template tags
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [Alias('SignatureIniPath')]
    [string]$SignatureIniFile = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Signatures DOCX\_Signatures.ini' } else { '.\sample templates\Signatures HTML\_Signatures.ini' }),

    # Deploy signatures for automapped mailboxes and additional mailboxes
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SignaturesForAutomappedAndAdditionalMailboxes = $true,

    # Shall the software delete signatures which were created by the user itself?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $DeleteUserCreatedSignatures = $false,

    # Shall the software delete signatures which were created by the software before but are no longer available as template?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $DeleteScriptCreatedSignaturesWithoutTemplate = $true,

    # Shall the software set the Outlook Web signature of the currently logged-in user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SetCurrentUserOutlookWebSignature = $true,

    # An additional path that the signatures shall be copied to
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [string]$AdditionalSignaturePath = $(try { $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Outlook Signatures')) } catch {}),

    # Should HTML signatures contain high resolution images?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $DocxHighResImageConversion = $true,

    # Create RTF signatures
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $CreateRtfSignatures = $false,

    # Create TXT signatures
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $CreateTxtSignatures = $true,

    # Move CSS to inline style attributes
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $MoveCSSInline = $true,

    # Embed images in HTML
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $EmbedImagesInHtml = $false,

    # Embed images in HTML for AdditionalSignaturePath
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $EmbedImagesInHtmlAdditionalSignaturePath = $true,

    # Should signature names be mailbox specific by adding the email address?
    [Parameter(Mandatory = $false, ParameterSetName = 'B: Signatures')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    $MailboxSpecificSignatureNames = $false,

    # Shall the software set the out-of-office (OOF) message(s) of the currently logged-in user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SetCurrentUserOOFMessage = $true,

    # Path to centrally managed out-of-office (OOF, automatic reply) templates
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$OOFTemplatePath = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Out-of-Office DOCX' } else { '.\sample templates\Out-of-Office HTML' }),

    # Path to INI file containing OOF template tags
    [Parameter(Mandatory = $false, ParameterSetName = 'C: OOF messages')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [Alias('OOFIniPath')]
    [string]$OOFIniFile = $(if (($UseHtmTemplates -inotin @(1, 'true', '$true', 'yes')) -or (-not $UseHtmTemplates)) { '.\sample templates\Out-of-Office DOCX\_OOF.ini' } else { '.\sample templates\Out-of-Office HTML\_OOF.ini' }),

    # Path to a replacement variable config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'D: Replacement variables')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$ReplacementVariableConfigFile = '.\config\default replacement variables.ps1',

    # Path to a virtual mailbox config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'D: Replacement variables')]
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$VirtualMailboxConfigFile = '',

    # Try to connect to Microsoft Graph only, ignoring any local Active Directory.
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $GraphOnly = $(if ($IsWindows -or (-not (Test-Path -LiteralPath 'variable:IsWindows'))) { $false } else { $true }),

    # GraphClientID, later overwritten by $GraphConfigFile
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    $GraphClientID = $null,

    # Cloud environment to use
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC', 'AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4', 'AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5', 'China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud')]
    [string]$CloudEnvironment = 'Public',

    # Path to a Graph variable config file.
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$GraphConfigFile = '.\config\default graph config.ps1',

    # List of domains/forests to check for group membership across trusts
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('DomainsToCheckForGroups')]
    [string[]]$TrustsToCheckForGroups = @('*'),

    # Shall the software consider group membership in domain local groups in the mailbox's AD forest?
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $IncludeMailboxForestDomainLocalGroups = $false,

    # Deploy while simulating
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SimulateAndDeploy = $false,

    # Path to file containing Graph credential which should be used as alternative to other token acquisition methods
    [Parameter(Mandatory = $false, ParameterSetName = 'E: Graph and Active Directory')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateNotNullOrEmpty()]
    [string]$SimulateAndDeployGraphCredentialFile = '',

    # Simulate another user as currently logged-in user
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('SimulationUser', 'WhatIf')]
    [validatescript({
            $tempSimulateUser = $_
            if ($tempSimulateUser -imatch '^\S+@\S+$|^\S+\\\S+$') {
                $true
            } else {
                throw "'$tempSimulateUser' does not match the required format 'User@Domain' (UPN) or 'Domain\User'."
            }
        }
    )]
    [string]$SimulateUser = $null,

    # Simulate list of mailboxes instead of mailboxes configured in Outlook
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('SimulationMailboxes')]
    [mailaddress[]]$SimulateMailboxes = $null,

    # Use a specific time for simulation mode
    [Parameter(Mandatory = $false, ParameterSetName = 'F: Simulation mode')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [Alias('SimulationTime')]
    [validatescript({
            $tempSimulateTime = $_
            if ($tempSimulateTime -imatch '\d{12}') {
                [DateTime]::ParseExact($tempSimulateTime, 'yyyyMMddHHmm', $null)
                $true
            } else {
                throw "'$tempSimulateTime' does not match the required format 'yyyyMMddHHmm'."
            }
        }
    )]
    [string]$SimulateTime = $null,

    # Should roaming signatures be disabled in Outlook?
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    $DisableRoamingSignatures = $true,

    # Should local signatures be uploaded as roaming signature for the current user?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no', 'CurrentUserOnly')]
    [Alias('MirrorLocalSignaturesToCloud')]
    $MirrorCloudSignatures = $true,

    # Word process priority
    [Parameter(Mandatory = $false, ParameterSetName = 'H: Word')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('Idle', 64, 'BelowNormal', 16384, 'Normal', 32, 'AboveNormal', 32768, 'High', 128, 'RealTime', 256)]
    $WordProcessPriority = 'Normal',

    # Script process priority
    [Parameter(Mandatory = $false, ParameterSetName = 'I: Script')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet('', 'Idle', 64, 'BelowNormal', 16384, 'Normal', 32, 'AboveNormal', 32768, 'High', 128, 'RealTime', 256)]
    $ScriptProcessPriority = '',

    # Should the 'SignatureCollectionInDrafts' email draft be created and updated?
    [Parameter(Mandatory = $false, ParameterSetName = 'A: Benefactor Circle')]
    [Parameter(Mandatory = $false, ParameterSetName = 'G: Outlook')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Z: All parameters')]
    [ValidateSet(1, 'true', '$true', 'yes', 0, 'false', '$false', 'no')]
    $SignatureCollectionInDrafts = $true
)


function ConvertToPSCustomObject {
    param($item)

    if ($item.PSObject.Methods.Name -contains 'GetEnumerator') {
        $tempHashtable = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($prop in $item.getenumerator()) {
            $tempHashtable.add($prop.name, $($prop.value))
        }

        return [PSCustomObject]$tempHashtable
    } else {
        return $item
    }
}


function ToSemVer($version) {
    $major = 0
    $minor = 0
    $patch = 0
    $pre = @()

    if (($version -ireplace '^v') -match '^(?<major>\d+)(\.(?<minor>\d+))?(\.(?<patch>\d+))?(\-(?<pre>[0-9A-Za-z\-\.]+))?(\+(?<build>[0-9A-Za-z\-\.]+))?$') {
        $major = [int]$matches['major']
        $minor = [int]$matches['minor']
        $patch = [int]$matches['patch']

        if ($null -eq $matches['pre']) {
            $pre = @()
        } else {
            $pre = $matches['pre'].Split('.')
        }
    }

    New-Object PSObject -Property @{
        Major         = $major
        Minor         = $minor
        Patch         = $patch
        Pre           = $pre
        VersionString = $version
    } | Select-Object -Property Major, Minor, Patch, Pre, VersionString
}


function CompareSemVer($a, $b) {
    $result = 0
    $result = $a.Major.CompareTo($b.Major)
    if ($result -ne 0) { return $result }

    $result = $a.Minor.CompareTo($b.Minor)
    if ($result -ne 0) { return $result }

    $result = $a.Patch.CompareTo($b.Patch)
    if ($result -ne 0) { return $result }

    $ap = $a.Pre
    $bp = $b.Pre

    if ($ap.Length -eq 0 -and $bp.Length -eq 0) { return 0 }
    if ($ap.Length -eq 0) { return 1 }
    if ($bp.Length -eq 0) { return -1 }

    $minLength = [Math]::Min($ap.Length, $bp.Length)

    for ($i = 0; $i -lt $minLength; $i++) {
        $ac = $ap[$i]
        $bc = $bp[$i]

        $anum = 0
        $bnum = 0
        $aIsNum = [Int]::TryParse($ac, [ref] $anum)
        $bIsNum = [Int]::TryParse($bc, [ref] $bnum)

        if ($aIsNum -and $bIsNum) {
            $result = $anum.CompareTo($bnum)
            if ($result -ne 0) {
                return $result
            }
        }
        if ($aIsNum) {
            return -1
        }
        if ($bIsNum) {
            return 1
        }

        $result = [string]::CompareOrdinal($ac, $bc)

        if ($result -ne 0) { return $result }
    }

    return $ap.Length.CompareTo($bp.Length)
}


function rankedSemVer($versions) {
    for ($i = 0; $i -lt $versions.Length; $i++) {
        $rank = 0

        for ($j = 0; $j -lt $versions.Length; $j++) {
            $diff = 0
            $diff = compareSemVer $versions[$i] $versions[$j]

            if ($diff -gt 0) {
                $rank++
            }
        }

        $current = [PsObject]$versions[$i]
        Add-Member -InputObject $current -MemberType NoteProperty -Name Rank -Value $rank -Force
    }

    return $versions
}


function CheckFilenamePossiblyInvalid ([string] $Filename = '', [bool] $CheckOutlook = $true, [bool] $CheckDeviceNames = $false) {
    $InvalidCharacters = @()

    # [System.Io.Path]::GetInvalidFileNameChars()
    $InvalidCharacters += @(($Filename | Select-String -Pattern "[$([regex]::escape(([System.Io.Path]::GetInvalidFileNameChars() -join '')))]" -AllMatches).Matches.Value) | Where-Object { $_ }

    # Outlook GUI
    if ($CheckOutlook) {
        $InvalidCharacters += @(($Filename | Select-String -Pattern "[$([regex]::escape('\/:"*?><,|@'))]" -AllMatches).Matches.Value) | Where-Object { $_ }
    }

    # Windows reserved file names and device names (CON, PRN, AUX, COMx, LPTx, …)
    if ($CheckDeviceNames) {
        if (([System.Io.Path]::GetFullPath($Filename)).StartsWith('\\.\')) {
            $InvalidCharacters += $Filename
        }
    }

    $InvalidCharacters = @(@($InvalidCharacters | Select-Object -Unique | Where-Object { $_ } | Sort-Object -Culture 127))

    if ($InvalidCharacters) {
        return $InvalidCharacters -join ', '
    }
}


### ▼▼▼ BlockSleep initiation code below ▼▼▼
##
#
# Place this code in your main script, as early in the code as possible
#
# Call BlockSleep wherever you want the current process to block sleep
#   BlockSleep
#
# On Windows, you can set three parameters:
#   -RequireAwaymode: Allows Away mode (defaults to true when not set)
#   -RequireDisplay: Requires the display to be on (defaults to false when not set)
#   -RequireSystem: Requires the system to be on (default to true when not set)
# On Linux, systemd-inhibit is required (should be available on most distributions)
# On macOS, caffeinate is required (should be available built-in)
#
# To allow sleep again, call BlockSleep with the AllowSleep parameter:
#   BlockSleep -AllowSleep
#
function BlockSleep {
    param (
        [switch]$AllowSleep,
        [switch]$RequireAwayMode,
        [switch]$RequireDisplay,
        [switch]$RequireSystem
    )

    if ($AllowSleep) {
        $RequireAwayMode = $false
        $RequireDisplay = $false
        $RequireSystem = $false
    } else {
        if (-not $PSBoundParameters.ContainsKey('RequireAwayMode')) {
            $RequireAwayMode = $true
        }

        if (-not $PSBoundParameters.ContainsKey('RequireDisplay')) {
            $RequireDisplay = $false
        }

        if (-not $PSBoundParameters.ContainsKey('RequireSystem')) {
            $RequireSystem = $true
        }

        if (
            ($RequireAwayMode -eq $false) -and
            ($RequireDisplay -eq $false) -and
            ($RequireSystem -eq $false)
        ) {
            $AllowSleep = $true
        }
    }

    if ($isWindows -or (-not (Test-Path -LiteralPath 'variable:IsWindows'))) {
        $code = @'
[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
public static extern void SetThreadExecutionState(uint esFlags);
'@

        $ste = Add-Type -MemberDefinition $code -Name System -Namespace Win32 -PassThru
        $ES_CONTINUOUS = [uint32]'0x80000000'
        $ES_AWAYMODE_REQUIRED = [uint32]'0x00000040'
        $ES_DISPLAY_REQUIRED = [uint32]'0x00000002'
        $ES_SYSTEM_REQUIRED = [uint32]'0x00000001'

        $flags = $ES_CONTINUOUS

        if (-not $AllowSleep) {
            if ($RequireAwayMode) { $flags = $flags -bor $ES_AWAYMODE_REQUIRED }
            if ($RequireDisplay) { $flags = $flags -bor $ES_DISPLAY_REQUIRED }
            if ($RequireSystem) { $flags = $flags -bor $ES_SYSTEM_REQUIRED }
        }

        $ste::SetThreadExecutionState($flags)
    } elseif ($isLinux) {
        if (Get-Command systemd-inhibit -ErrorAction SilentlyContinue) {
            if ($script:BlockSleepInhibitPID) {
                Stop-Process -Id $script:BlockSleepInhibitPID -Force
                Remove-Variable -Name BlockSleepInhibitPID -Scope script
            }

            if (-not $AllowSleep) {
                $script:BlockSleepInhibitPID = Start-Process systemd-inhibit -ArgumentList "--what=idle --why=""Set-OutlookSignatures"" --who=""Set-OutlookSignatures"" tail --pid=$($PID) --follow /dev/null" -PassThru | Select-Object -ExpandProperty Id
            }
        } else {
            Write-Host "  'systemd-inhibit' is not available."
        }
    } elseif ($isMacOS) {
        if (Get-Command caffeinate -ErrorAction SilentlyContinue) {
            if ($script:BlockSleepInhibitPID) {
                Stop-Process -Id $script:BlockSleepInhibitPID -Force
                Remove-Variable -Name BlockSleepInhibitPID -Scope script
            }

            if (-not $AllowSleep) {
                $script:BlockSleepInhibitPID = Start-Process caffeinate -ArgumentList "-ims -w $($PID)" -PassThru | Select-Object -ExpandProperty Id
            }
        } else {
            Write-Host "  'caffeinate' is not available."
        }
    }
}
#
##
### ▲▲▲ BlockSleep initiation code above ▲▲▲


function main {
    $ScriptVersion = 'XXXVersionStringXXX'

    try { WatchCatchableExitSignal } catch { }

    # Init default values
    if ($null -ne [SetOutlookSignatures.Common].GetMethod('Init')) {
        [SetOutlookSignatures.Common]::Init()

        if (-not $SetOutlookSignaturesCommonInitDone) {
            $script:ExitCode = 5
            $script:ExitCodeDescription = 'Common initialization routine failed.'
            exit
        }
    } else {
        Write-Host 'Error initializing Set-OutlookSignatures. Exiting.' -ForegroundColor Red
        $script:ExitCode = 6
        $script:ExitCodeDescription = 'Common initialization routine not available.'
        exit
    }


    try { WatchCatchableExitSignal } catch { }


    # Import QRCoder
    $script:QRCoderModulePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))

    Copy-Item -LiteralPath ((Join-Path -Path '.' -ChildPath 'bin\QRCoder\netstandard2.0')) -Destination $script:QRCoderModulePath -Recurse
    if (-not $IsLinux) { Get-ChildItem -LiteralPath $script:QRCoderModulePath -Recurse | Unblock-File }
    Import-Module (Join-Path -Path $script:QRCoderModulePath -ChildPath 'QRCoder.dll')


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Get basic Outlook and Word information @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $OutlookProfiles = @()
    $OutlookUseNewOutlook = $null

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled, skip Outlook checks'
    } else {
        if ($IsWindows) {
            Write-Host '  Outlook'

            if ($(Get-Command -Name 'Get-AppPackage' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                $NewOutlook = Get-AppPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            } else {
                $NewOutlook = $null
            }

            $OutlookRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Outlook.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))

            if ($OutlookRegistryVersion -eq [System.Version]::Parse('0.0.0.0')) {
                $OutlookRegistryVersion = $null
            }

            try {
                # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
                # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                # Covers:
                #   Office x86 on Windows x86
                #   Office x86 on Windows x64
                #   Any PowerShell process bitness
                $OutlookFilePath = Get-ChildItem -LiteralPath (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
            } catch {
                try {
                    # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
                    # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                    # Covers:
                    #   Office x64 on Windows x64
                    #   Any PowerShell process bitness
                    $OutlookFilePath = Get-ChildItem -LiteralPath (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
                } catch {
                    $OutlookFilePath = $null
                }
            }

            if ($OutlookFilePath) {
                try {
                    $OutlookBitnessInfo = GetBitness -fullname $OutlookFilePath
                    $OutlookFileVersion = [System.Version]::Parse((((($OutlookBitnessInfo.'File Version'.ToString() + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
                    $OutlookBitness = $OutlookBitnessInfo.Architecture
                    Remove-Variable -Name 'OutlookBitnessInfo'
                } catch {
                    $OutlookBitness = $null
                    $OutlookFilePath = $null
                    $OutlookFileVersion = $null
                }
            } else {
                $OutlookBitness = $null
                $OutlookFilePath = $null
                $OutlookFileVersion = $null
            }

            if ($OutlookRegistryVersion.Major -ne $OutlookFileVersion.Major) {
                Write-Host "    Major parts of Outlook version from registry ('$OutlookRegistryVersion') and from outlook.exe ('$OutlookFileVersion') do not match." -ForegroundColor Yellow
                Write-Host '    Assuming that Outlook is not installed.' -ForegroundColor Yellow
                Write-Host '    To resolve this, repair the Outlook installation and/or the registry information about Outlook.' -ForegroundColor Yellow

                $OutlookRegistryVersion = $null
                $OutlookFilePath = $null
                $OutlookFileVersion = $null
                $OutlookBitness = $null
            }

            if ($null -ne $OutlookRegistryVersion) {
                if ($OutlookRegistryVersion.major -gt 16) {
                    Write-Host "    Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
                    $script:ExitCode = 7
                    $script:ExitCodeDescription = 'Outlook version newer than 16 is not yet known.'
                    exit
                } elseif ($OutlookRegistryVersion.major -eq 16) {
                    $OutlookRegistryVersion = '16.0'
                } elseif ($OutlookRegistryVersion.major -eq 15) {
                    $OutlookRegistryVersion = '15.0'
                } elseif ($OutlookRegistryVersion.major -eq 14) {
                    $OutlookRegistryVersion = '14.0'
                } elseif ($OutlookRegistryVersion.major -lt 14) {
                    Write-Host "    Outlook version $OutlookRegistryVersion is older than Outlook 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
                    $script:ExitCode = 8
                    $script:ExitCodeDescription = 'Outlook version older than 2010 is not supported.'
                    exit
                }
            }

            if ($null -ne $OutlookRegistryVersion) {
                Write-Host "    Set 'Send Pictures With Document' registry value to '1'"
                $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Options\Mail" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'Send Pictures With Document' -Type DWORD -Value 1 -Force
            }

            if (($DisableRoamingSignatures -in @($true, $false)) -and $OutlookRegistryVersion -and ($OutlookFileVersion -ge '16.0.0.0')) {
                Write-Host "    Set 'DisableRoamingSignatures' registry value to '$([int]$DisableRoamingSignatures)'"
                $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableRoamingSignaturesTemporaryToggle' -Type DWORD -Value $([int]$DisableRoamingSignatures) -Force
                $null = "HKCU:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableRoamingSignatures' -Type DWORD -Value $([int]$DisableRoamingSignatures) -Force
            }

            if ($null -ne $OutlookRegistryVersion) {
                try {
                    $OutlookDefaultProfile = (Get-ItemProperty -LiteralPath "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\Outlook" -ErrorAction Stop -WarningAction SilentlyContinue).DefaultProfile

                    $OutlookProfiles = @(@((Get-ChildItem -LiteralPath "hkcu:\SOFTWARE\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles" -ErrorAction Stop -WarningAction SilentlyContinue).PSChildName) | Where-Object { $_ })

                    if ($OutlookDefaultProfile -and ($OutlookDefaultProfile -iin $OutlookProfiles)) {
                        $OutlookProfiles = @(@($OutlookDefaultProfile) + @($OutlookProfiles | Where-Object { $_ -ine $OutlookDefaultProfile }))
                    }
                } catch {
                    $OutlookDefaultProfile = $null
                    $OutlookProfiles = @()
                }

                $OutlookIsBetaversion = $false

                if (
                    ((Get-Item -LiteralPath 'registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Property -contains 'UpdateChannel') -and
                    ($OutlookFileVersion -ge '16.0.0.0')
                ) {
                    $x = (Get-ItemProperty -LiteralPath 'registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop -WarningAction SilentlyContinue).'UpdateChannel'

                    if ($x -ieq 'http://officecdn.microsoft.com/pr/5440FD1F-7ECB-4221-8110-145EFAA6372F') {
                        $OutlookIsBetaversion = $true
                    }

                    if ((Get-Item -LiteralPath "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Common\OfficeUpdate" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Property -contains 'UpdateBranch') {
                        $x = (Get-ItemProperty -LiteralPath "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Common\OfficeUpdate" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).'UpdateBranch'

                        if ($x -ieq 'InsiderFast') {
                            $OutlookIsBetaversion = $true
                        }
                    }
                }

                $OutlookDisableRoamingSignatures = 0

                foreach ($RegistryFolder in @(
                        "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                        "registry::HKEY_LOCAL_MACHINE\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                        "registry::HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup",
                        "registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Setup"
                    )
                ) {
                    try { WatchCatchableExitSignal } catch { }

                    $x = (Get-ItemProperty -LiteralPath $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignaturesTemporaryToggle'

                    if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                        $OutlookDisableRoamingSignatures = $x
                    }

                    $x = (Get-ItemProperty -LiteralPath $RegistryFolder -ErrorAction SilentlyContinue).'DisableRoamingSignatures'

                    if (($x -in (0, 1)) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                        $OutlookDisableRoamingSignatures = $x
                    }
                }

                if ($NewOutlook -and ($((Get-ItemProperty -LiteralPath "registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Preferences" -ErrorAction SilentlyContinue).'UseNewOutlook') -eq 1)) {
                    $OutlookUseNewOutlook = $true
                    $OutlookDisableRoamingSignatures = 1
                } else {
                    $OutlookUseNewOutlook = $false
                }
            } else {
                $OutlookDefaultProfile = $null
                $OutlookDisableRoamingSignatures = 1
                $OutlookIsBetaVersion = $null

                if ($NewOutlook) {
                    $OutlookUseNewOutlook = $true
                } else {
                    $OutlookUseNewOutlook = $false
                }
            }

            Write-Host "    Registry version: $OutlookRegistryVersion"
            Write-Host "    File version: $OutlookFileVersion"
            if (($OutlookFileVersion -lt '16.0.0.0') -and ($EmbedImagesInHtml -eq $true)) {
                Write-Host '      Outlook 2013 or earlier detected.' -ForegroundColor Yellow
                Write-Host '      Consider parameter ''-EmbedImagesInHtml false'' to avoid problems with images in templates.' -ForegroundColor Yellow
                Write-Host '      Microsoft supports Outlook 2013 until April 2023, older versions are already out of support.' -ForegroundColor Yellow
            }
            Write-Host "    Bitness: $OutlookBitness"
            Write-Host "    Default profile: $OutlookDefaultProfile"
            Write-Host "    Is C2R Beta: $OutlookIsBetaversion"
            Write-Host "    DisableRoamingSignatures: $OutlookDisableRoamingSignatures"
            if (($OutlookDisableRoamingSignatures -eq 0) -and ($OutlookFileVersion -ge '16.0.0.0')) {
                Write-Host '      Outlook syncs signatures itself, so it may overwrite signatures created by this software.' -ForegroundColor Yellow
                Write-Host '      Consider setting parameters DisableRoamingSignatures and MirrorCloudSignatures to true instead.' -ForegroundColor Yellow
                Write-Host '      Also consider using the MailboxSpecificSignaturesNames parameter.' -ForegroundColor Yellow
            }

            Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"
            Write-Host '  New Outlook'
            Write-Host "    Version: $($NewOutlook.Version)"
            Write-Host "    Status: $($NewOutlook.Status)"
            Write-Host "    UseNewOutlook: $OutlookUseNewOutlook"
        } elseif ($IsMacOS) {
            Write-Host '  Outlook'

            $macOsIsRunningNewOutlook = ($(defaults read com.microsoft.Outlook IsRunningNewOutlook *>&1).ToString() -eq 1)

            $OutlookFileVersion = @(@($(
                        @'
tell application "Microsoft Outlook"
            get version
end tell
'@ | osascript *>&1)) | ForEach-Object { $_.tostring() })[0]

            Write-Host "    Version: $($OutlookFileVersion)"

            try { WatchCatchableExitSignal } catch { }

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

            try { WatchCatchableExitSignal } catch { }

            $macOSOutlookMailboxes = @(@($(
                        @'
tell application "Microsoft Outlook"
    try
        set exchangeAccounts to get exchange accounts
    on error
        set exchangeAccounts to {}
    end try

    try
        set popAccounts to get pop accounts
    on error
        set popAccounts to {}
    end try

    try
        set imapAccounts to get imap accounts
    on error
        set imapAccounts to {}
    end try

    try
        set ldapAccounts to get ldap accounts
    on error
        set ldapAccounts to {}
    end try

    try
        set delegatedAccounts to get delegated accounts
    on error
        set delegatedAccounts to {}
    end try

    try
        set otherAccounts to get other users folder accounts
    on error
        set otherAccounts to {}
    end try

    set allAccounts to exchangeAccounts & popAccounts & imapAccounts & ldapAccounts & delegatedAccounts & otherAccounts

    repeat with singleAccount in allAccounts
        set x to email address of singleAccount
        log x
    end repeat
end tell
'@ | osascript *>&1)) | ForEach-Object { $_.tostring() })

            try { WatchCatchableExitSignal } catch { }

            $OutlookFilePath = $null
            $OutlookRegistryVersion = $null
            $OutlookDefaultProfile = $null
            $OutlookProfiles = @()
            $OutlookIsBetaversion = $false
            $OutlookDisableRoamingSignatures = 1
            $OutlookUseNewOutlook = $false
            $script:WordRegistryVersion = $null
            $WordFilePath = $null

            if ($macOSSignaturesScriptable) {
                Write-Host '    Outlook for Mac with scriptable signatures detected.'

                $EmbedImagesInHtml = $true

                if ($macOSOutlookMailboxes.count -gt 0) {
                    Write-Host '    Outlook has accounts configured.'
                } else {
                    if ($macOsIsRunningNewOutlook) {
                        Write-Host '    No accounts detected via AppleScript, but New Outlook is enabled. Trying alternate detection method.'

                        if (Test-Path -LiteralPath '~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist') {
                            $macOSOutlookMailboxes = @($(@((ConvertEncoding -InFile '~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist' -InIsHtml $false) -split '\r?\n') | Where-Object { $_ -match '.*actionsEndPointURLFor.*' } | ForEach-Object { $_.trim() -ireplace '<key>ActionsEndPointURLFor', '' -ireplace '</key>', '' }))

                            if ($macOSOutlookMailboxes.count -gt 0) {
                                Write-Host '      Accounts found. If too many accounts are found:'
                                Write-Host '        1. Quit Outlook'
                                Write-Host '        2. Delete ''~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist'''
                                Write-Host '        3. Start Outlook and run Set-OutlookSignatures'
                            }
                        } else {
                            Write-Host "      Failed. '~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles/Main Profile/ProfilePreferences.plist' not found." -ForegroundColor Yellow
                        }
                    }

                    if (-not ($macOSOutlookMailboxes.count -gt 0)) {
                        Write-Host '    Outlook does not have accounts configured, or accounts cannot be scripted. Continuing with Outlook Web only.' -ForegroundColor Yellow
                        Write-Host "      Consider using 'sample code/SwitchTo-ClassicOutlookForMac.ps1' to temporarily switch from New Outlook to Classic Outlook." -ForegroundColor Yellow

                        $OutlookUseNewOutlook = $true
                        $macOSOutlookMailboxes = @()
                    }
                }
            } else {
                Write-Host '    Outlook for Mac not installed, or signatures cannot be scripted. Continuing with Outlook Web only.' -ForegroundColor Yellow

                $OutlookUseNewOutlook = $true
                $macOSOutlookMailboxes = @()
            }
        } else {
            $OutlookFilePath = $null
            $OutlookRegistryVersion = $null
            $OutlookDefaultProfile = $null
            $OutlookProfiles = @()
            $OutlookIsBetaversion = $false
            $OutlookDisableRoamingSignatures = 1
            $OutlookUseNewOutlook = $true
            $script:WordRegistryVersion = $null
            $WordFilePath = $null
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ((($UseHtmTemplates -eq $true) -and (-not $CreateRtfSignatures)) -or (-not $IsWindows)) {
        Write-Host '  UseHtmTemplates set to true or not running on Windows, skip Word checks'
    } else {
        Write-Host '  Word'

        $script:WordRegistryVersion = $null

        $script:WordAlertIfNotDefaultOriginal = (Get-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -ErrorAction SilentlyContinue).AlertIfNotDefault

        $script:WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Word.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
        if ($script:WordRegistryVersion.major -gt 16) {
            Write-Host "    Word version $($script:WordRegistryVersion) is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
            $script:ExitCode = 9
            $script:ExitCodeDescription = 'Word version newer than 16 is not yet known.'
            exit
        } elseif ($script:WordRegistryVersion.major -eq 16) {
            $script:WordRegistryVersion = '16.0'
        } elseif ($script:WordRegistryVersion.major -eq 15) {
            $script:WordRegistryVersion = '15.0'
        } elseif ($script:WordRegistryVersion.major -eq 14) {
            $script:WordRegistryVersion = '14.0'
        } elseif ($script:WordRegistryVersion.major -lt 14) {
            Write-Host "    Word version $($script:WordRegistryVersion) is older than Word 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
            $script:ExitCode = 10
            $script:ExitCodeDescription = 'Word version older than 2010 is not supported.'
            exit
        }

        try {
            # [Microsoft.Win32.RegistryView]::Registry32 makes sure view the registry as a 32 bit application would
            # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
            # Covers:
            #   Office x86 on Windows x86
            #   Office x86 on Windows x64
            #   Any PowerShell process bitness
            $WordFilePath = Get-ChildItem -LiteralPath (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry32)).OpenSubKey("CLSID\$((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
        } catch {
            try {
                # [Microsoft.Win32.RegistryView]::Registry64 makes sure we view the registry as a 64 bit application would
                # This is independent from the bitness of the PowerShell process, while Get-ItemProperty always uses the bitness of the PowerShell process
                # Covers:
                #   Office x64 on Windows x64
                #   Any PowerShell process bitness
                $WordFilePath = Get-ChildItem -LiteralPath (((([Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Registry64)).OpenSubKey("CLSID\$((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Word.Application\CLSID' -ErrorAction Stop).'(default)')\LocalServer32")).GetValue('') -split ' \/')[0].Split([IO.Path]::GetInvalidPathChars()) -join '').trim('"').trim('''') -ErrorAction Stop
            } catch {
                $WordFilePath = $null
            }
        }

        if ($WordFilePath) {
            Write-Host "    Set 'DontUseScreenDpiOnOpen' registry value to '1'"
            $null = "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DontUseScreenDpiOnOpen' -Type DWORD -Value 1 -Force

            try {
                $WordBitnessInfo = GetBitness -fullname $WordFilePath
                $WordFileVersion = [System.Version]::Parse((((($WordBitnessInfo.'File Version'.ToString() + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
                $WordBitness = $WordBitnessInfo.Architecture
                Remove-Variable -Name 'WordBitnessInfo'
            } catch {
                $WordBitness = 'Error'
                $WordFileVersion = $null
            }
        } else {
            $WordBitness = $null
            $WordFileVersion = $null
        }

        Write-Host "    Registry version: $script:WordRegistryVersion"
        Write-Host "    File version: $WordFileVersion"
        Write-Host "    Bitness: $WordBitness"
    }

    try { WatchCatchableExitSignal } catch { }

    Write-Host
    Write-Host "Get Outlook signature file path(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $SignaturePaths = @()

    if ($SimulateUser) {
        Write-Host '  Simulation mode enabled. Skip task, use AdditionalSignaturePath instead'
        if ($AdditionalSignaturePath) {
            $SignaturePaths += $AdditionalSignaturePath
        }
    } elseif ($OutlookProfiles -and ($OutlookUseNewOutlook -ne $true)) {
        $x = (Get-ItemProperty -LiteralPath "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'

        if ($x) {
            Push-Location -LiteralPath ((Join-Path -Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))

            if (Test-Path -LiteralPath $x -IsValid) {
                if (-not (Test-Path -LiteralPath $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }

                if ($x -inotin $SignaturePaths) {
                    $SignaturePaths += $x
                    Write-Host "  '$x'"
                }
            }

            Pop-Location
        }
    } else {
        $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname))

        if ($Iswindows) {
            Write-Host "  '$($SignaturePaths[-1])' (Outlook Web/New Outlook)"
        } elseif ($IsMacOS) {
            if ($macOSSignaturesScriptable) {
                Write-Host "  '$($SignaturePaths[-1])' (Outlook for Mac with scriptable signatures)"
            } else {
                Write-Host "  '$($SignaturePaths[-1])' (Outlook Web, because no Outlook, no accounts configured or signatures not scriptable)"
            }
        } elseif ($IsLinux) {
            Write-Host "  '$($SignaturePaths[-1])' (Outlook Web)"
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # If Outlook is installed, synch profile folders anyway
    # Also makes sure that signatures are already there when starting Outlook for the first time
    if ((-not $SimulateUser) -and $OutlookFileVersion) {
        $x = (Get-ItemProperty -LiteralPath "hkcu:\software\microsoft\office\$($OutlookRegistryVersion)\common\general" -ErrorAction SilentlyContinue).'Signatures'

        if ($x) {
            Push-Location -LiteralPath ((Join-Path -Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) -ChildPath 'Microsoft'))
            $x = ($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($x))

            if (Test-Path -LiteralPath $x -IsValid) {
                if (-not (Test-Path -LiteralPath $x -type container)) {
                    New-Item -Path $x -ItemType directory -Force | Out-Null
                }

                if ($x -inotin $SignaturePaths) {
                    $SignaturePaths += $x
                    Write-Host "  '$x'"
                }
            }

            Pop-Location
        }

        $SignaturePaths = @($SignaturePaths | Select-Object -Unique)
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Enumerate domains @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $x = $TrustsToCheckForGroups
    [System.Collections.ArrayList]$TrustsToCheckForGroups = @()
    $LookupDomainsToTrusts = @{}

    if ($GraphOnly -eq $false) {
        # Users own domain/forest is always included
        try {
            $objTrans = New-Object -ComObject 'NameTranslate'
            $objNT = $objTrans.GetType()
            $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
            $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $(([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value))) # 12 = ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME
            try { WatchCatchableExitSignal } catch { }
            $UserForest = (([ADSI]"LDAP://$(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1) -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -ireplace [Regex]::Escape('DC='), '' -ireplace [Regex]::Escape(','), '.').tolower()
            try { WatchCatchableExitSignal } catch { }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
            Remove-Variable -Name 'objTrans'
            Remove-Variable -Name 'objNT'

            if ($UserForest -ne '') {
                Write-Host "  User forest: $UserForest"

                if ($TrustsToCheckForGroups -inotcontains $UserForest) {
                    $TrustsToCheckForGroups += $UserForest.tolower()
                }

                if (-not $LookupDomainsToTrusts.ContainsKey($UserForest.tolower())) {
                    $LookupDomainsToTrusts.add($UserForest.tolower(), $UserForest.tolower())
                }

                $Search.SearchRoot = "GC://$($UserForest)"
                $Search.Filter = '(ObjectClass=trustedDomain)'

                try { WatchCatchableExitSignal } catch { }
                $TrustedDomains = @($Search.FindAll())
                try { WatchCatchableExitSignal } catch { }

                if ($TrustedDomains) {
                    $TrustedDomains = @(
                        @($TrustedDomains) | Where-Object { $_ -ine $UserForest } | Sort-Object -Culture 127 -Property @{Expression = {
                                $TemporaryArray = @($_.properties.name.Split('.'))
                                [Array]::Reverse($TemporaryArray)
                                $TemporaryArray
                            }
                        }
                    )
                }

                try { WatchCatchableExitSignal } catch { }

                # Internal trusts
                foreach ($TrustedDomain in $TrustedDomains) {
                    if (($TrustedDomain.properties.trustattributes -eq 32) -and ($TrustedDomain.properties.name -ine $UserForest) -and (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower()))) {
                        Write-Host "    Child domain: $($TrustedDomain.properties.name.tolower())"

                        if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                            $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $UserForest.tolower())
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # Other trusts
                if ($x[0] -eq '*') {
                    foreach ($TrustedDomain in $TrustedDomains) {
                        # No intra-forest trusts, only bidirectional trusts and outbound trusts
                        if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3))) {
                            if ($TrustedDomain.properties.trustattributes -eq 8) {
                                # Cross-forest trust
                                Write-Host "  Trusted forest: $($TrustedDomain.properties.name.tolower())"
                                if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                    Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name.tolower())'"
                                } else {
                                    if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                        $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                    }

                                    if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                        $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }

                                try { WatchCatchableExitSignal } catch { }

                                $temp = @(
                                    @(@(Resolve-DnsName -Name "_gc._tcp.$($TrustedDomain.properties.name)" -Type srv).nametarget) | ForEach-Object { ($_ -split '\.')[1..999] -join '.' } | Where-Object { $_ -ine $TrustedDomain.properties.name } | Select-Object -Unique | Sort-Object -Culture 127 -Property @{Expression = {
                                            $TemporaryArray = @($_.Split('.'))
                                            [Array]::Reverse($TemporaryArray)
                                            $TemporaryArray
                                        }
                                    }
                                )

                                $temp | ForEach-Object {
                                    Write-Host "    Child domain: $($_.tolower())"

                                    if (-not $LookupDomainsToTrusts.ContainsKey($_.tolower())) {
                                        $LookupDomainsToTrusts.add($_.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }
                            } else {
                                # No cross-forest trust
                                Write-Host "  Trusted domain: $($TrustedDomain.properties.name)"
                                if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                    Write-Host "    Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                } else {
                                    if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                        $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                    }

                                    if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                        $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                    }
                                }
                            }
                        }
                    }
                }

                for ($a = 0; $a -lt $x.Count; $a++) {
                    if (($a -eq 0) -and ($x[$a] -ieq '*')) {
                        continue
                    }

                    $y = ($x[$a] -ireplace [Regex]::Escape('DC='), '' -ireplace ',', '.').tolower()

                    if ($y -eq $x[$a]) {
                        Write-Host "  User provided trusted domain/forest: $y"
                    } else {
                        Write-Host "  User provided trusted domain/forest: $($x[$a]) -> $y"
                    }

                    if (($a -ne 0) -and ($x[$a] -ieq '*')) {
                        Write-Host '    Entry * is only allowed at first position in list. Skip entry.' -ForegroundColor Red
                        continue
                    }

                    if ($y -imatch '[^a-zA-Z0-9.-]') {
                        Write-Host '    Allowed characters are a-z, A-Z, ., -. Skip entry.' -ForegroundColor Red
                        continue
                    }

                    if (-not ($y.StartsWith('-'))) {
                        if ($TrustsToCheckForGroups -icontains $y) {
                            Write-Host '    Trusted domain/forest already in list.'
                        } else {
                            if ($TrustedDomains.properties.name -icontains $y) {
                                foreach ($TrustedDomain in @($TrustedDomains | Where-Object { $_.properties.name -ieq $y })) {
                                    # No intra-forest trusts, only bidirectional trusts and outbound trusts
                                    if (($($TrustedDomain.properties.trustattributes) -ne 32) -and (($($TrustedDomain.properties.trustdirection) -eq 2) -or ($($TrustedDomain.properties.trustdirection) -eq 3))) {
                                        if ($TrustedDomain.properties.trustattributes -eq 8) {
                                            # Cross-forest trust
                                            Write-Host "    Trusted forest: $($TrustedDomain.properties.name)"
                                            if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                                Write-Host "      Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                            } else {
                                                if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                                    $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                                }

                                                if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                                    $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                                }
                                            }

                                            try { WatchCatchableExitSignal } catch { }

                                            $temp = @(
                                                @(@(Resolve-DnsName -Name "_gc._tcp.$($TrustedDomain.properties.name)" -Type srv).nametarget) | ForEach-Object { ($_ -split '\.')[1..999] -join '.' } | Where-Object { $_ -ine $TrustedDomain.properties.name } | Select-Object -Unique | Sort-Object -Culture 127 -Property @{Expression = {
                                                        $TemporaryArray = @($_.Split('.'))
                                                        [Array]::Reverse($TemporaryArray)
                                                        $TemporaryArray
                                                    }
                                                }
                                            )

                                            $temp | ForEach-Object {
                                                Write-Host "      Child domain: $($_.tolower())"

                                                if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                                    $LookupDomainsToTrusts.add($_.tolower(), $TrustedDomain.properties.name.tolower())
                                                }
                                            }
                                        } else {
                                            # No cross-forest trust
                                            Write-Host "    Trusted domain: $($TrustedDomain.properties.name)"
                                            if ("-$($TrustedDomain.properties.name)" -iin $x) {
                                                Write-Host "      Ignoring because of TrustsToCheckForGroups entry '-$($TrustedDomain.properties.name)'"
                                            } else {
                                                if ($TrustsToCheckForGroups -inotcontains $TrustedDomain.properties.name) {
                                                    $TrustsToCheckForGroups += $TrustedDomain.properties.name.tolower()
                                                }

                                                if (-not $LookupDomainsToTrusts.ContainsKey($TrustedDomain.properties.name.tolower())) {
                                                    $LookupDomainsToTrusts.add($TrustedDomain.properties.name.tolower(), $TrustedDomain.properties.name.tolower())
                                                }
                                            }
                                        }
                                    }
                                }
                            } else {
                                Write-Host '    No trust to this domain/forest found.' -ForegroundColor Yellow
                            }
                        }
                    } else {
                        Write-Host '    Remove trusted domain/forest.'
                        for ($z = 0; $z -lt $TrustsToCheckForGroups.Count; $z++) {
                            if ($TrustsToCheckForGroups[$z] -ieq $y.substring(1)) {
                                $TrustsToCheckForGroups.RemoveAt($z)
                                $LookupDomainsToTrusts = $LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $y.substring(1) }
                            }
                        }
                    }
                }

                $TrustsToCheckForGroups = @($TrustsToCheckForGroups | Where-Object { $_ })


                try { WatchCatchableExitSignal } catch { }


                Write-Host
                Write-Host "Check trusts for open LDAP port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                CheckADConnectivity @(@(@($TrustsToCheckForGroups) + @($LookupDomainsToTrusts.GetEnumerator() | ForEach-Object { $_.Name })) | Select-Object -Unique) 'LDAP' '  ' | Out-Null


                try { WatchCatchableExitSignal } catch { }


                Write-Host
                Write-Host "Check trusts for open Global Catalog port and connectivity @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
                CheckADConnectivity $TrustsToCheckForGroups 'GC' '  ' | Out-Null
            } else {
                Write-Host '  Problem connecting to logged-in user''s Active Directory (no error message, but forest root domain name is empty).' -ForegroundColor Yellow
                Write-Host '  Assuming Graph/Entra ID from now on.' -ForegroundColor Yellow
                $GraphOnly = $true
            }
        } catch {
            Write-Verbose "  $($error[0])"
            $y = ''
            Write-Host "  Problem connecting to logged-in user's Active Directory, see verbose output for details." -ForegroundColor Yellow
            Write-Host '  Assuming Graph/Entra ID from now on.' -ForegroundColor Yellow
            $GraphOnly = $true
        }
    } else {
        Write-Host "  Parameter GraphOnly set to '$GraphOnly', ignore user's Active Directory in favor of Graph/Entra ID."
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Get properties of currently logged-in user and assigned manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (-not $SimulateUser) {
        Write-Host '  Currently logged-in user'
    } else {
        Write-Host "  Simulate '$SimulateUser' as currently logged-in user"
    }

    if ($GraphOnly -eq $false) {
        if ($null -ne $TrustsToCheckForGroups[0]) {
            try {
                if (-not $SimulateUser) {
                    $Search.SearchRoot = "GC://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUser = $Search.FindOne().Properties

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUser = [hashtable]::new($ADPropsCurrentUser, [StringComparer]::OrdinalIgnoreCase)

                    $Search.SearchRoot = "LDAP://$((([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$(([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName)))"

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUserLdap = $Search.FindOne().Properties

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUserLdap = [hashtable]::new($ADPropsCurrentUserLdap, [StringComparer]::OrdinalIgnoreCase)

                    foreach ($keyName in @($ADPropsCurrentUserLdap.Keys)) {
                        if (
                            $($keyName -inotin $ADPropsCurrentUser.Keys) -or
                            $(-not ($ADPropsCurrentUser[$keyName]) -and ($ADPropsCurrentUserLdap[$keyName]))
                        ) {
                            $ADPropsCurrentUser[$keyName] = $ADPropsCurrentUserLdap[$keyName]
                        }
                    }
                } else {
                    try {
                        $objTrans = New-Object -ComObject 'NameTranslate'
                        $objNT = $objTrans.GetType()
                        $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $null))
                        $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, $SimulateUser))

                        try { WatchCatchableExitSignal } catch { }

                        $SimulateUserDN = $objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 1)

                        try { WatchCatchableExitSignal } catch { }

                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                        Remove-Variable -Name 'objTrans'
                        Remove-Variable -Name 'objNT'

                        $Search.SearchRoot = "GC://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsCurrentUser = $Search.FindOne().Properties

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsCurrentUser = [hashtable]::new($ADPropsCurrentUser, [StringComparer]::OrdinalIgnoreCase)

                        $Search.SearchRoot = "LDAP://$(($SimulateUserDN -split ',DC=')[1..999] -join '.')"
                        $Search.Filter = "((distinguishedname=$SimulateUserDN))"

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsCurrentUserLdap = $Search.FindOne().Properties

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsCurrentUserLdap = [hashtable]::new($ADPropsCurrentUserLdap, [StringComparer]::OrdinalIgnoreCase)

                        foreach ($keyName in @($ADPropsCurrentUserLdap.Keys)) {
                            if (
                                $($keyName -inotin $ADPropsCurrentUser.Keys) -or
                                $(-not ($ADPropsCurrentUser[$keyName]) -and ($ADPropsCurrentUserLdap[$keyName]))
                            ) {
                                $ADPropsCurrentUser[$keyName] = $ADPropsCurrentUserLdap[$keyName]
                            }
                        }
                    } catch {
                        Write-Host "    $($error[0])"
                        Write-Host "    Simulation user '$($SimulateUser)' not found. Exit." -ForegroundColor REd
                        $script:ExitCode = 11
                        $script:ExitCodeDescription = 'Simulation user not found.'
                        exit
                    }
                }

                $ADPropsCurrentUser = ConvertToPSCustomObject -item $ADPropsCurrentUser
            } catch {
                Write-Host $error[0]
                $ADPropsCurrentUser = $null
                Write-Host '    Problem connecting to Active Directory, or user is a local user. Exit.' -ForegroundColor Red
                $script:ExitCode = 12
                $script:ExitCodeDescription = 'Problem connecting to Active Directory, or user is a local user.'
                exit
            }
        }
    }

    if (
        ($GraphOnly -eq $true) -or
        (($GraphOnly -eq $false) -and ($ADPropsCurrentUser.msexchrecipienttypedetails -ge 2147483648) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true))) -or
        (($GraphOnly -eq $false) -and ($null -eq $ADPropsCurrentUser)) -or
        ($OutlookUseNewOutlook -eq $true) -or
        $(
            if (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('LicenseGroupRequiresGraph'))) {
                $result = [SetOutlookSignatures.BenefactorCircle]::LicenseGroupRequiresGraph()

                if ($result -ine 'false') {
                    $true
                } else {
                    $false
                }
            } else {
                $false
            }
        )
    ) {
        Write-Host "    Enforcing Graph$(if ($null -ne $TrustsToCheckForGroups[0]) { ' instead of Active Directory' }) because at least one condition is true:"
        Write-Host "      GraphOnly is true: $($GraphOnly -eq $true)"
        Write-Host "      GraphOnly is false, mailbox is in cloud, SetCurrentUserOOFMessage and/or SetCurrentUserOutlookWebSignature is true: $(($GraphOnly -eq $false) -and ($ADPropsCurrentUser.msexchrecipienttypedetails -ge 2147483648) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true)))"
        Write-Host "      GraphOnly is false and on-prem AD properties of current user are empty: $(($GraphOnly -eq $false) -and ($null -eq $ADPropsCurrentUser))"
        Write-Host "      New Outlook is used: $($OutlookUseNewOutlook -eq $true)"
        Write-Host "      The only Benefactor Circle license group is in Entra ID: $(
            if (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('LicenseGroupRequiresGraph'))) {
                $result = [SetOutlookSignatures.BenefactorCircle]::LicenseGroupRequiresGraph()

                if ($result -ine 'false') {
                    $true
                } else {
                    $false
                }
            } else {
                $false
            }
        )"

        $GraphOnly = $true
        [System.Collections.ArrayList]$TrustsToCheckForGroups = @()

        if (-not $script:GraphToken) {
            GraphGetTokenWrapper -indent '    '
        }

        if ($script:GraphToken -and (-not $SimulateAndDeployGraphCredentialFile)) {
            Write-Host "      Graph token cache: $($script:msalClientApp.cacheInfo)"
        }

        if ($script:GraphToken.error -eq $false) {
            Write-Verbose "      Graph Token metadata: $((ParseJwtToken $script:GraphToken.AccessToken) | ConvertTo-Json)"

            Write-Verbose "      Graph Token EXO metadata: $((ParseJwtToken $script:GraphToken.AccessTokenExo) | ConvertTo-Json)"

            if ($SimulateAndDeployGraphCredentialFile) {
                Write-Verbose "      App Graph Token metadata: $((ParseJwtToken $script:GraphToken.AppAccessToken) | ConvertTo-Json)"

                Write-Verbose "      App Graph Token EXO metadata: $((ParseJwtToken $script:GraphToken.AppAccessTokenExo) | ConvertTo-Json)"
            }
        } else {
            Write-Host '      Problem connecting to Microsoft Graph. Exit.' -ForegroundColor Red
            Write-Host $script:GraphToken.error -ForegroundColor Red
            $script:ExitCode = 14
            $script:ExitCodeDescription = 'Problem connecting to Microsoft Graph.'
            exit
        }

        if ($SimulateUser) {
            $script:GraphUser = $SimulateUser
        }

        GraphSwitchContext -TenantID $script:GraphUser
        $x = (GraphGetUserProperties $script:GraphUser)

        if (($x.error -eq $false) -and ($x.properties.id)) {
            $AADProps = $x.properties
            $ADPropsCurrentUser = [PSCustomObject]@{}

            foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                $z = $AADProps

                foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                    $z = $z.$y
                }

                $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
            }

            $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $script:GraphUser).photo -Force
            $ADPropsCurrentUser | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $script:GraphUser).properties.userprincipalname -Force
        } else {
            Write-Host "      Problem getting data for '$($script:GraphUser)' from Microsoft Graph. Exit." -ForegroundColor Red
            Write-Host $x.error -ForegroundColor Red
            $script:ExitCode = 15
            $script:ExitCodeDescription = "Problem getting data for '$($script:GraphUser)' from Microsoft Graph."
            exit
        }
    }

    if ($ADPropsCurrentUser) {
        Write-Host "    DistinguishedName: $($ADPropsCurrentUser.distinguishedname)"
        Write-Host "    UserPrincipalName: $($ADPropsCurrentUser.userprincipalname)"
        Write-Host "    Mail: $($ADPropsCurrentUser.mail)"
    } else {
        Write-Host '    User not found'
    }


    try { WatchCatchableExitSignal } catch { }

    $CurrentUserSIDs = @()

    if (($ADPropsCurrentUser.objectsid -ne '') -and ($null -ne $ADPropsCurrentUser.objectsid)) {
        if ($ADPropsCurrentUser.objectsid.tostring().startswith('S-', 'CurrentCultureIgnorecase')) {
            $CurrentUserSids += $ADPropsCurrentUser.objectsid.tostring()
        } else {
            $CurrentUserSids += (New-Object system.security.principal.securityidentifier($ADPropsCurrentUser.objectsid, 0)).value
        }
    }

    if (($ADPropsCurrentUser.onpremisessecurityidentifier -ne '') -and ($null -ne $ADPropsCurrentUser.onpremisessecurityidentifier)) {
        $CurrentUserSids += $ADPropsCurrentUser.onpremisessecurityidentifier.tostring()
    }

    foreach ($SidHistorySid in @($ADPropsCurrentUser.sidhistory | Where-Object { $_ })) {
        try { WatchCatchableExitSignal } catch { }

        if ($SidHistorySid.tostring().startswith('S-', 'CurrentCultureIgnorecase')) {
            $CurrentUserSids += $SidHistorySid.tostring()
        } else {
            $CurrentUserSids += (New-Object system.security.principal.securityidentifier($SidHistorySid, 0)).value
        }
    }

    if (-not $SimulateUser) {
        Write-Host '  Manager of currently logged-in user'
    } else {
        Write-Host '  Manager of simulated currently logged-in user'
    }

    $ADPropsCurrentUserManager = $null

    if ($ADPropsCurrentUser -and ($ADPropsCurrentUser.manager)) {
        if ($ADPropsCurrentUser.manager -imatch '(\S+?)@(\S+?)\.(\S+?)') {
            # Manager is in UPN format, search via Graph
            # Graph connection must already be established, else the manager would not be in UPN format

            Write-Verbose "    Search manager '$($ADPropsCurrentUser.manager)' via Graph"

            try {
                $AADProps = (GraphGetUserProperties $ADPropsCurrentUser.manager).properties
                $ADPropsCurrentUserManager = [PSCustomObject]@{}

                foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                    $z = $AADProps

                    foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                        $z = $z.$y
                    }

                    $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
                }

                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsCurrentUserManager.userprincipalname).photo -Force
                $ADPropsCurrentUserManager | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null -Force
            } catch {
                $ADPropsCurrentUserManager = $null
            }
        } else {
            # Manager is not in UPN format, try search on-prem
            # But only if ($GraphOnly -ne $true)

            Write-Verbose "    Search manager '$($ADPropsCurrentUser.manager)' on-prem"

            if ($GraphOnly -ne $true) {
                try {
                    $Search.SearchRoot = "GC://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUserManager = $Search.FindOne().Properties

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUserManager = [hashtable]::new($ADPropsCurrentUserManager, [StringComparer]::OrdinalIgnoreCase)


                    $Search.SearchRoot = "LDAP://$(($ADPropsCurrentUser.manager -split ',DC=')[1..999] -join '.')"
                    $Search.Filter = "((distinguishedname=$($ADPropsCurrentUser.manager)))"

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUserManagerLdap = $Search.FindOne().Properties

                    try { WatchCatchableExitSignal } catch { }

                    $ADPropsCurrentUserManagerLdap = [hashtable]::new($ADPropsCurrentUserManagerLdap, [StringComparer]::OrdinalIgnoreCase)

                    foreach ($keyName in @($ADPropsCurrentUserManagerLdap.Keys)) {
                        if (
                            $($keyName -inotin $ADPropsCurrentUserManager.Keys) -or
                            $(-not ($ADPropsCurrentUserManager[$keyName]) -and ($ADPropsCurrentUserManagerLdap[$keyName]))
                        ) {
                            $ADPropsCurrentUserManager[$keyName] = $ADPropsCurrentUserManagerLdap[$keyName]
                        }
                    }
                } catch {
                    $ADPropsCurrentUserManager = $null
                }

                $ADPropsCurrentUserManager = ConvertToPSCustomObject -item $ADPropsCurrentUserManager
            } else {
                $ADPropsCurrentUserManager = $null

                Write-Verbose "    Undefined combination: GraphOnly is set to true, but manager '$($ADPropsCurrentUser.manager)' is not in UPN format."
            }
        }
    }

    if ($ADPropsCurrentUserManager) {
        Write-Host "    DistinguishedName: $($ADPropsCurrentUserManager.distinguishedname)"
        Write-Host "    UserPrincipalName: $($ADPropsCurrentUserManager.userprincipalname)"
        Write-Host "    Mail: $($ADPropsCurrentUserManager.mail)"
    } else {
        Write-Host '    No manager found'
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Get email addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    $MailAddresses = @()
    $RegistryPaths = @()
    $LegacyExchangeDNs = @()

    if ($SimulateUser -and $SimulateMailboxes) {
        Write-Host '  Simulation mode enabled and SimulateMailboxes defined, use SimulateMailboxes as mailbox list'
        for ($i = 0; $i -lt $SimulateMailboxes.count; $i++) {
            $MailAddresses += $SimulateMailboxes[$i].ToLower()
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''
        }
    } elseif ($IsWindows -and $OutlookProfiles -and ($OutlookUseNewOutlook -ne $true)) {
        Write-Host '  Get email addresses from Outlook'

        foreach ($OutlookProfile in $OutlookProfiles) {
            try { WatchCatchableExitSignal } catch { }

            Write-Host "    Profile '$($OutlookProfile)'"

            foreach ($RegistryFolder in @(Get-ChildItem -LiteralPath "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookProfile\9375CFF0413111d3B88A00104B2A6676" -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { if ($OutlookFileVersion -ge '16.0.0.0') { ($_.'Account Name' -like '*@*.*') } else { (($_.'Account Name' -join ',') -like '*,64,*,46,*') } })) {
                try { WatchCatchableExitSignal } catch { }

                if ($OutlookFileVersion -ge '16.0.0.0') {
                    $MailAddresses += ($RegistryFolder.'Account Name').ToLower()
                } else {
                    $MailAddresses += (@(foreach ($char in @(($RegistryFolder.'Account Name' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join '').ToLower()
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

                Write-Host "      $($MailAddresses[-1])"
                Write-Verbose "        Registry: $($RegistryFolder.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $RegistryFolder.PSDrive)"
                Write-Verbose "        LegacyExchangeDN: $($LegacyExchangeDNs[-1])"
            }

            if ($SignaturesForAutomappedAndAdditionalMailboxes) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SignaturesForAutomappedAndAdditionalMailboxes')))) {
                    Write-Host '    Automapped and additional mailboxes will not be found.' -ForegroundColor Green
                    Write-Host "    The 'SignaturesForAutomappedAndAdditionalMailboxes' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                    Write-Host '    Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignaturesForAutomappedAndAdditionalMailboxes()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '      Error finding automapped and additional mailboxes.' -ForegroundColor Yellow
                        Write-Host "      $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "    Parameter 'SignaturesForAutomappedAndAdditionalMailboxes' is not enabled, skipping task."
            }
        }
    } elseif ($IsMacOS -and $macOSSignaturesScriptable -and ($macOSOutlookMailboxes.count -gt 0)) {
        Write-Host '  Get email addresses from Outlook'

        $macOSOutlookMailboxes | ForEach-Object {
            $MailAddresses += $_
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''

            Write-Host "    $($MailAddresses[-1])"
            Write-Verbose "      Registry: $($RegistryPaths[-1])"
            Write-Verbose "      LegacyExchangeDN: $($LegacyExchangeDNs[-1])"
        }
    } else {
        if ($IsWindows -and $OutlookUseNewOutlook) {
            Write-Host '  Get email addresses from New Outlook and Outlook Web, as New Outlook is set as default'
        } else {
            Write-Host '  Get email addresses from Outlook Web'
        }

        $OutlookProfiles = @()
        $OutlookDefaultProfile = $null

        $script:GraphUserDummyMailbox = $true

        if ($IsWindows -and $OutlookUseNewOutlook -eq $true) {
            $x = @(
                @((ConvertEncoding -InFile $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Microsoft\Olk\UserSettings.json') | ConvertFrom-Json).Identities.IdentityMap.PSObject.Properties | Select-Object -Unique | Where-Object { $_.name -match '(\S+?)@(\S+?)\.(\S+?)' }) | ForEach-Object {
                    if ((ConvertEncoding -InFile $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath "\Microsoft\OneAuth\accounts\$($_.Value)") | ConvertFrom-Json).association_status -ilike '*"com.microsoft.Olk":"associated"*') {
                        $_.name
                    }
                }
            )
        } else {
            $x = @()
        }

        if ($ADPropsCurrentUser.mail) {
            if ($x -icontains $ADPropsCurrentUser.mail) {
                $x = @($ADPropsCurrentUser.mail.tolower()) + @($x | Where-Object { $_ -ine $ADPropsCurrentUser.mail })
            } else {
                $x = @($ADPropsCurrentUser.mail.tolower()) + $x
            }
        } else {
            Write-Host '    User does not have mail attribute configured' -ForegroundColor Yellow
            $script:GraphUserDummyMailbox = $false
        }

        $x | ForEach-Object {
            $MailAddresses += $_.ToLower()
            $RegistryPaths += ''
            $LegacyExchangeDNs += ''

            Write-Host "    $($MailAddresses[-1])"
            Write-Verbose "      Registry: $($RegistryFolder.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $RegistryFolder.PSDrive)"
            Write-Verbose "      LegacyExchangeDN: $($LegacyExchangeDNs[-1])"

            if ($ADPropsCurrentUser.mail -and ($_ -ieq $ADPropsCurrentUser.mail)) {
                $PrimaryMailboxAddress = $ADPropsCurrentUser.mail

                if ($SignaturesForAutomappedAndAdditionalMailboxes) {
                    if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SignaturesForAutomappedAndAdditionalMailboxes')))) {
                        Write-Host '    Automapped and additional mailboxes will not be found.' -ForegroundColor Green
                        Write-Host "    The 'SignaturesForAutomappedAndAdditionalMailboxes' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                        Write-Host '    Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                    } else {
                        try { WatchCatchableExitSignal } catch { }

                        $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignaturesForAutomappedAndAdditionalMailboxes()

                        if ($FeatureResult -ne 'true') {
                            Write-Host '    Error finding automapped and additional mailboxes.' -ForegroundColor Yellow
                            Write-Host "    $FeatureResult" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "    Parameter 'SignaturesForAutomappedAndAdditionalMailboxes' is not enabled, skipping task."
                }
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ((($SetCurrentUserOutlookWebSignature -eq $true) -or ($SetCurrentUserOOFMessage -eq $true)) -and ($MailAddresses -inotcontains $ADPropsCurrentUser.mail)) {
        # OOF and/or Outlook web signature must be set, but user does not seem to have a mailbox in Outlook
        # Maybe this is a pure Outlook Web user, so we will add a helper entry
        # This entry fakes the users mailbox in his default Outlook profile, so it gets the highest priority later
        Write-Host "  User's mailbox not found in email address list, but Outlook Web signature and/or OOF message should be set. Adding dummy mailbox entry." -ForegroundColor Yellow

        if ($ADPropsCurrentUser.mail) {
            $script:GraphUserDummyMailbox = $true

            $SignaturePaths = @(((New-Item -ItemType Directory (Join-Path -Path $script:tempDir -ChildPath ((New-Guid).guid))).fullname)) + $SignaturePaths

            $MailAddresses = @($ADPropsCurrentUser.mail.tolower()) + $MailAddresses
            $RegistryPaths = @('') + $RegistryPaths
            $LegacyExchangeDNs = @('') + $LegacyExchangeDNs
        } else {
            Write-Host '      User does not have mail attribute configured.' -ForegroundColor Yellow
            $script:GraphUserDummyMailbox = $false
        }
    } else {
        $script:GraphUserDummyMailbox = $false
    }

    try { WatchCatchableExitSignal } catch { }

    if ($MailAddresses.count -eq 0) {
        Write-Host
        Write-Host 'No email addresses found, exiting.'
        Write-Host '  In simulation mode, this might be a permission problem.'
        $script:ExitCode = 16
        $script:ExitCodeDescription = 'No email addresses found.'
        exit
    }


    try { WatchCatchableExitSignal } catch { }


    $ADPropsMailboxes = @()
    $ADPropsMailboxesUserDomain = @()
    $ADPropsMailboxManagers = @()

    Write-Host
    Write-Host "Get properties of each mailbox and its manager @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        Write-Host "  $($MailAddresses[$AccountNumberRunning])"

        $UserDomain = ''
        $ADPropsMailboxes += $null
        $ADPropsMailboxesUserDomain += $null
        $ADPropsMailboxManagers += $null
        $GroupsSIDs = @()

        $CurrentMailboxAlreadyFoundFirstIndex = $null

        for ($i = 0; $i -lt $ADPropsMailboxes.Count; $i++) {
            if ($ADPropsMailboxes[$i].proxyaddresses -icontains "smtp:$($MailAddresses[$AccountNumberRunning])") {
                $CurrentMailboxAlreadyFoundFirstIndex = $i
                break
            }
        }

        if (
            $null -eq $CurrentMailboxAlreadyFoundFirstIndex
        ) {
            if (
                (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '') -or ($($MailAddresses[$AccountNumberRunning]) -ne ''))
            ) {
                if ($null -ne $TrustsToCheckForGroups[0]) {
                    # Loop through domains until the first one knows the legacyExchangeDN or the proxy address
                    for ($DomainNumber = 0; (($DomainNumber -lt $TrustsToCheckForGroups.count) -and ($UserDomain -eq '')); $DomainNumber++) {
                        try { WatchCatchableExitSignal } catch { }

                        if (($TrustsToCheckForGroups[$DomainNumber] -ne '')) {
                            Write-Host "    Search for mailbox user object in domain/forest '$($TrustsToCheckForGroups[$DomainNumber])'"

                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustsToCheckForGroups[$DomainNumber])")

                            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                                $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(|(legacyExchangeDN=$($LegacyExchangeDNs[$AccountNumberRunning]))(&(legacyExchangeDN=*)(proxyaddresses=x500:$($LegacyExchangeDNs[$AccountNumberRunning])))))"
                            } elseif (($($MailAddresses[$AccountNumberRunning]) -ne '')) {
                                $Search.filter = "(&(ObjectCategory=person)(objectclass=user)(|(msexchrecipienttypedetails<=32)(msexchrecipienttypedetails>=2147483648))(msExchMailboxGuid=*)(legacyExchangeDN=*)(proxyaddresses=smtp:$($MailAddresses[$AccountNumberRunning])))"
                            }

                            try { WatchCatchableExitSignal } catch { }

                            $u = $Search.FindAll()

                            try { WatchCatchableExitSignal } catch { }

                            if ($u.count -eq 0) {
                                Write-Host '      Not found'
                            } elseif ($u.count -gt 1) {
                                Write-Host '      Multiple matches found' -ForegroundColor Yellow

                                foreach ($SingleU in $u) {
                                    Write-Host "      $($SingleU.path)" -ForegroundColor Yellow
                                }

                                Write-Host '        Check why your Active Directory returns multiple results for the following query:' -ForegroundColor Yellow
                                Write-Host "          $($Search.SearchRoot)" -ForegroundColor Yellow
                                Write-Host "          $($Search.Filter)" -ForegroundColor Yellow

                                $LegacyExchangeDNs[$AccountNumberRunning] = ''
                                $MailAddresses[$AccountNumberRunning] = ''
                                $UserDomain = $null
                            } else {
                                $Search.SearchRoot = "GC://$(($(([adsi]"$($u[0].path)").distinguishedname) -split ',DC=')[1..999] -join '.')"
                                $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"

                                try { WatchCatchableExitSignal } catch { }

                                $ADPropsMailboxes[$AccountNumberRunning] = $Search.FindOne().Properties

                                try { WatchCatchableExitSignal } catch { }

                                $ADPropsMailboxes[$AccountNumberRunning] = [hashtable]::new($ADPropsMailboxes[$AccountNumberRunning], [StringComparer]::OrdinalIgnoreCase)

                                $Search.SearchRoot = "LDAP://$(($(([adsi]"$($u[0].path)").distinguishedname) -split ',DC=')[1..999] -join '.')"
                                $Search.Filter = "((distinguishedname=$(([adsi]"$($u[0].path)").distinguishedname)))"

                                try { WatchCatchableExitSignal } catch { }

                                $tempLdap = $Search.FindOne().Properties

                                try { WatchCatchableExitSignal } catch { }

                                $tempLdap = [hashtable]::new($tempLdap, [StringComparer]::OrdinalIgnoreCase)

                                foreach ($keyName in @($tempLdap.Keys)) {
                                    if (
                                        ($keyName -inotin $ADPropsMailboxes[$AccountNumberRunning].Keys) -or
                                        (-not ($ADPropsMailboxes[$AccountNumberRunning][$keyName]) -and ($tempLdap[$keyName]))
                                    ) {
                                        $ADPropsMailboxes[$AccountNumberRunning][$keyName] = $tempLdap[$keyName]
                                    }
                                }

                                $ADPropsMailboxes[$AccountNumberRunning] = ConvertToPSCustomObject -item $ADPropsMailboxes[$AccountNumberRunning]

                                $UserDomain = $TrustsToCheckForGroups[$DomainNumber]
                                $ADPropsMailboxesUserDomain[$AccountNumberRunning] = $TrustsToCheckForGroups[$DomainNumber]
                                $LegacyExchangeDNs[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].legacyexchangedn
                                $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()
                                Write-Host "      DistinguishedName: $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                                Write-Host "      UserPrincipalName: $($ADPropsMailboxes[$AccountNumberRunning].userprincipalname)"
                                Write-Host "      Mail: $($ADPropsMailboxes[$AccountNumberRunning].mail)"
                                Write-Host "      Manager: $($ADPropsMailboxes[$AccountNumberRunning].manager)"
                            }
                        }
                    }

                    if ($u.count -eq 0) {
                        Write-Host '      No matching mailbox object found in any Active Directory. See verbose output for details.' -ForegroundColor Yellow
                        Write-Host '      This message can be ignored if the mailbox in question is not part of your environment.' -ForegroundColor Yellow
                        Write-Verbose "        You may have restricted the accessible environment with the 'TrustsToCheckForGroups' parameter."
                        Write-Verbose '        Else, check why the following Active Directory query did not return a result:'
                        Write-Verbose "          $($Search.Filter)"
                        Write-Verbose '        Usual root causes: Mailbox added in Outlook no longer exists or is not in your tenant, Exchange data in Active Directory is not complete, firewall rules, DNS.'
                        Write-Verbose "        Check if all required attributes documented in the 'README' file are available in your on-prem Active Directory and have values."
                        Write-Verbose "          Look for 'msExchMailboxGuid' in the 'README' file for details about the required attributes."
                        Write-Verbose '        For hybrid environments:'
                        Write-Verbose '          Add missing msExchMailboxGuid for cloud mailboxes to on-prem AD: https://learn.microsoft.com/en-US/exchange/troubleshoot/move-mailboxes/migrationpermanentexception-when-moving-mailboxes.'
                        Write-Verbose "          Consider using the '-GraphOnly true' parameter to not query on-prem Active Directory at all."
                    }

                    if (-not $ADPropsMailboxes[$AccountNumberRunning]) {
                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $UserDomain = $null
                    } elseif ($ADPropsMailboxManagers[$AccountNumberRunning].manager) {
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")

                        try {
                            $Search.filter = "(distinguishedname=$($ADPropsMailboxes[$AccountNumberRunning].manager))"

                            try { WatchCatchableExitSignal } catch { }

                            $ADPropsMailboxManagers[$AccountNumberRunning] = ([ADSI]"$(($Search.FindOne()).path)").Properties

                            try { WatchCatchableExitSignal } catch { }

                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")
                            $Search.filter = "(distinguishedname=$($ADPropsMailboxes[$AccountNumberRunning].manager))"

                            try { WatchCatchableExitSignal } catch { }

                            $ADPropsCurrentMailboxManagerLdap = ([ADSI]"$(($Search.FindOne()).path)").Properties

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($keyName in @($ADPropsCurrentMailboxManagerLdap.Keys)) {
                                if (
                                    ($keyName -inotin $ADPropsMailboxManagers[$AccountNumberRunning].Keys) -or
                                    (-not ($ADPropsMailboxManagers[$AccountNumberRunning][$keyName]) -and ($ADPropsCurrentMailboxManagerLdap[$keyName]))
                                ) {
                                    $ADPropsMailboxManagers[$AccountNumberRunning][$keyName] = $ADPropsCurrentMailboxManagerLdap[$keyName]
                                }
                            }

                            $ADPropsMailboxManagers[$AccountNumberRunning] = ConvertToPSCustomObject -item $ADPropsMailboxManagers[$AccountNumberRunning]

                            Write-Host "        DistinguishedName: $($ADPropsMailboxManagers[$AccountNumberRunning].distinguishedname)"
                            Write-Host "        UserPrincipalName: $($ADPropsMailboxManagers[$AccountNumberRunning].userprincipalname)"
                            Write-Host "        Mail: $($ADPropsMailboxManagers[$AccountNumberRunning].mail)"
                        } catch {
                            $ADPropsMailboxManagers[$AccountNumberRunning] = @()
                        }
                    }
                } else {
                    Write-Host '    Search for mailbox user object in Graph'

                    $ADPropsMailboxes[$AccountNumberRunning] = [PSCustomObject]@{}

                    try { WatchCatchableExitSignal } catch { }

                    $AADProps = (GraphGetUserProperties $($MailAddresses[$AccountNumberRunning])).properties

                    try { WatchCatchableExitSignal } catch { }

                    if ($AADProps) {
                        foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                            $z = $AADProps

                            foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                $z = $z.$y
                            }

                            $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
                        }

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).photo -Force

                        try { WatchCatchableExitSignal } catch { }

                        $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'manager' -Value (GraphGetUserManager $ADPropsMailboxes[$AccountNumberRunning].userprincipalname).properties.userprincipalname -Force

                        try { WatchCatchableExitSignal } catch { }

                        if (-not $LegacyExchangeDNs[$AccountNumberRunning]) {
                            $LegacyExchangeDNs[$AccountNumberRunning] = 'dummy'
                        }

                        $MailAddresses[$AccountNumberRunning] = $ADPropsMailboxes[$AccountNumberRunning].mail.tolower()

                        Write-Host "      DistinguishedName: $($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                        Write-Host "      UserPrincipalName: $($ADPropsMailboxes[$AccountNumberRunning].userprincipalname)"
                        Write-Host "      Mail: $($ADPropsMailboxes[$AccountNumberRunning].mail)"
                        Write-Host "      Manager: $($ADPropsMailboxes[$AccountNumberRunning].manager)"

                        if ($ADPropsMailboxes[$AccountNumberRunning].manager) {
                            # get properties of mailbox manager here

                            try {
                                $AADProps = $null

                                if ($ADPropsMailboxes[$AccountNumberRunning].manager) {
                                    try { WatchCatchableExitSignal } catch { }

                                    $AADProps = (GraphGetUserProperties $ADPropsMailboxes[$AccountNumberRunning].manager).properties

                                    try { WatchCatchableExitSignal } catch { }

                                    $ADPropsMailboxManagers[$AccountNumberRunning] = [PSCustomObject]@{}

                                    foreach ($GraphUserAttributeMappingName in $GraphUserAttributeMapping.GetEnumerator()) {
                                        $z = $AADProps

                                        foreach ($y in ($GraphUserAttributeMappingName.value -split '\.')) {
                                            $z = $z.$y
                                        }

                                        $ADPropsMailboxManagers[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name ($GraphUserAttributeMappingName.Name) -Value $z -Force
                                    }

                                    try { WatchCatchableExitSignal } catch { }

                                    $ADPropsMailboxManagers[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'thumbnailphoto' -Value (GraphGetUserPhoto $ADPropsMailboxManagers[$AccountNumberRunning].userprincipalname).photo -Force

                                    try { WatchCatchableExitSignal } catch { }

                                    $ADPropsMailboxManagers[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'manager' -Value $null -Force

                                    try { WatchCatchableExitSignal } catch { }

                                    Write-Host "        DistinguishedName: $($ADPropsMailboxManagers[$AccountNumberRunning].distinguishedname)"
                                    Write-Host "        UserPrincipalName: $($ADPropsMailboxManagers[$AccountNumberRunning].userprincipalname)"
                                    Write-Host "        Mail: $($ADPropsMailboxManagers[$AccountNumberRunning].mail)"
                                }

                                try { WatchCatchableExitSignal } catch { }
                            } catch {
                                $ADPropsMailboxManagers[$AccountNumberRunning] = @()
                                Write-Host '        Skipping, mailbox manager not in Microsoft Graph.' -ForegroundColor yellow
                            }
                        }
                    } else {
                        Write-Host '      No matching mailbox object found via Graph/Entra ID. See verbose output for details.' -ForegroundColor Yellow
                        Write-Host '      This message can be ignored if the mailbox in question is not part of your environment.' -ForegroundColor Yellow
                        Write-Verbose '        Check why the following Graph queries return zero or more than 1 results, or do not contain any properties:'
                        Write-Verbose "          UserPrincipalName from: $("$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$filter=proxyAddresses/any(x:x eq 'smtp:$($MailAddresses[$AccountNumberRunning])')")"
                        Write-Verbose "          Replace XXX with UPN from query above: $("$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/XXX?`$select=" + [System.Net.WebUtility]::UrlEncode($(@($GraphUserProperties | Select-Object -Unique) -join ',')))"
                        Write-Verbose '        Usual root causes: Mailbox added in Outlook no longer exists or is not in your tenant, firewall rules, DNS.'

                        $LegacyExchangeDNs[$AccountNumberRunning] = ''
                        $UserDomain = $null
                        $ADPropsMailboxManagers[$AccountNumberRunning] = $null
                    }
                }

                Write-Host '      Get group membership of mailbox'
                if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                    try { WatchCatchableExitSignal } catch { }

                    if ($null -ne $TrustsToCheckForGroups[0]) {
                        Write-Host "        $($ADPropsMailboxesUserDomain[$AccountNumberRunning]) (mailbox home domain/forest)"

                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($ADPropsMailboxesUserDomain[$AccountNumberRunning])")

                        $UserDomain = $ADPropsMailboxesUserDomain[$AccountNumberRunning]
                        $SIDsToCheckInTrusts = @()

                        if ($ADPropsMailboxes[$AccountNumberRunning].objectsid) {
                            $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier($ADPropsMailboxes[$AccountNumberRunning].objectsid, 0)).value
                        }

                        foreach ($SidHistorySid in @($ADPropsMailboxes[$AccountNumberRunning].sidhistory | Where-Object { $_ })) {
                            $SIDsToCheckInTrusts += (New-Object System.Security.Principal.SecurityIdentifier($SidHistorySid, 0)).value
                        }

                        try { WatchCatchableExitSignal } catch { }

                        try {
                            # Security groups, global and universal, forest-wide
                            Write-Host '          LDAP query for security groups (global and universal, forest-wide, via tokengroupsglobalanduniversal)'
                            $UserAccount = [ADSI]"LDAP://$($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)"
                            try { WatchCatchableExitSignal } catch { }
                            $UserAccount.GetInfoEx(@('tokengroupsglobalanduniversal'), 0)
                            try { WatchCatchableExitSignal } catch { }

                            foreach ($sidBytes in $UserAccount.Properties.tokengroupsglobalanduniversal) {
                                $sid = (New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)).value
                                Write-Verbose "            $($sid) (global or universal group, incl. sIDHistory)"
                                $GroupsSIDs += $sid
                                $SIDsToCheckInTrusts += $sid
                            }

                            try { WatchCatchableExitSignal } catch { }

                            # Distribution groups (static only)
                            try { WatchCatchableExitSignal } catch { }
                            Write-Host '          GC query for static distribution groups (global and universal, forest-wide)'
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')")
                            $Search.filter = "(&(objectClass=group)(!(groupType:1.2.840.113556.1.4.803:=2147483648))(member:1.2.840.113556.1.4.1941:=$($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)))"

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($DistributionGroup in $search.findall()) {
                                try { WatchCatchableExitSignal } catch { }

                                if ($DistributionGroup.properties.objectsid) {
                                    $sid = (New-Object System.Security.Principal.SecurityIdentifier($DistributionGroup.properties.objectsid[0], 0)).value
                                    Write-Verbose "            $($sid) (static distribution group)"
                                    $GroupsSIDs += $sid
                                    $SIDsToCheckInTrusts += $sid
                                }

                                foreach ($SidHistorySid in @($DistributionGroup.properties.sidhistory | Where-Object { $_ })) {
                                    $sid = (New-Object System.Security.Principal.SecurityIdentifier($SidHistorySid, 0)).value
                                    Write-Verbose "            $($sid) (static distribution group sIDHistory)"
                                    $GroupsSIDs += $sid
                                    $SIDsToCheckInTrusts += $sid
                                }
                            }

                            try { WatchCatchableExitSignal } catch { }

                            # Domain local groups
                            if ($IncludeMailboxForestDomainLocalGroups -eq $true) {
                                Write-Host '        LDAP query for domain local groups (security and distribution, one query per domain)'

                                foreach ($DomainToCheckForDomainLocalGroups in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ieq $LookupDomainsToTrusts[$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')] }).name)) {
                                    try { WatchCatchableExitSignal } catch { }
                                    Write-Host "          $($DomainToCheckForDomainLocalGroups)"
                                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DomainToCheckForDomainLocalGroups)")
                                    $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($ADPropsMailboxes[$AccountNumberRunning].distinguishedname)))"

                                    try { WatchCatchableExitSignal } catch { }

                                    foreach ($LocalGroup in $search.findall()) {
                                        try { WatchCatchableExitSignal } catch { }

                                        if ($LocalGroup.properties.objectsid) {
                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier($LocalGroup.properties.objectsid[0], 0)).value
                                            Write-Verbose "            $($sid) (domain local group)"
                                            $GroupsSIDs += $sid
                                            $SIDsToCheckInTrusts += $sid
                                        }

                                        foreach ($SidHistorySid in @($LocalGroup.properties.sidhistory | Where-Object { $_ })) {
                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier($SidHistorySid, 0)).value
                                            Write-Verbose "            $($sid) (domain local group sIDHistory)"
                                            $GroupsSIDs += $sid
                                            $SIDsToCheckInTrusts += $sid
                                        }
                                    }
                                }
                            }
                        } catch {
                            Write-Host $error[0]
                            Write-Host "            Error getting group information from $((($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.'), check firewalls, DNS and AD trust" -ForegroundColor Red
                        }

                        try { WatchCatchableExitSignal } catch { }

                        $GroupsSIDs = @($GroupsSIDs | Select-Object -Unique | Sort-Object -Culture 127)

                        # Loop through all domains outside the mailbox account's home forest to check if the mailbox account has a group membership there
                        # Across a trust, a user can only be added to a domain local group.
                        # Domain local groups cannot be used outside their own domain, so we don't need to query recursively
                        # But when it's a cross-forest trust, we need to query every every domain on that other side of the trust
                        #   This is handled before by adding every single domain of a cross-forest trusted forest to $TrustsToCheckForGroups
                        if ($SIDsToCheckInTrusts.count -gt 0) {
                            $SIDsToCheckInTrusts = @($SIDsToCheckInTrusts | Select-Object -Unique)
                            $LdapFilterSIDs = '(|'

                            foreach ($SidToCheckInTrusts in $SIDsToCheckInTrusts) {
                                try { WatchCatchableExitSignal } catch { }

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
                                    Write-Host $error[0]
                                    Write-Host '        Error creating LDAP filter for search across trusts.' -ForegroundColor Red
                                }
                            }
                            $LdapFilterSIDs += ')'
                        } else {
                            $LdapFilterSIDs = ''
                        }

                        if ($LdapFilterSids -ilike '*(objectsid=*') {
                            # Across each trust, search for all Foreign Security Principals matching a SID from our list
                            foreach ($TrustToCheckForFSPs in @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')] }).value | Select-Object -Unique)) {
                                try { WatchCatchableExitSignal } catch { }

                                Write-Host "        $($TrustToCheckForFSPs) (trusted domain/forest of mailbox home forest) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($TrustToCheckForFSPs)")
                                $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"

                                try { WatchCatchableExitSignal } catch { }
                                $fsps = $Search.FindAll()
                                try { WatchCatchableExitSignal } catch { }

                                if ($fsps.count -gt 0) {
                                    foreach ($fsp in $fsps) {
                                        try { WatchCatchableExitSignal } catch { }

                                        if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                                            # A Foreign Security Principal (FSP) is created in each (sub)domain in which it is granted permissions
                                            # A FSP it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal
                                            # FSPs have no tokengroups or tokengroupsglobalanduniversal attribute, which would not contain domain local groups anyhow
                                            # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) returns groups containing a specific DN as member, incl. nesting
                                            Write-Verbose "          Found ForeignSecurityPrincipal $($fsp.properties.cn) in $((($fsp.path -split ',DC=')[1..999] -join '.'))"

                                            if ($((($fsp.path -split ',DC=')[1..999] -join '.')) -iin @(($LookupDomainsToTrusts.GetEnumerator() | Where-Object { $_.Value -ine $LookupDomainsToTrusts[$(($($ADPropsMailboxes[$AccountNumberRunning].distinguishedname) -split ',DC=')[1..999] -join '.')] }).name)) {
                                                try {
                                                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$((($fsp.path -split ',DC=')[1..999] -join '.'))")
                                                    $Search.filter = "(&(objectClass=group)(groupType:1.2.840.113556.1.4.803:=4)(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname)))"

                                                    try { WatchCatchableExitSignal } catch { }
                                                    $fspGroups = $Search.FindAll()
                                                    try { WatchCatchableExitSignal } catch { }

                                                    if ($fspGroups.count -gt 0) {
                                                        foreach ($group in $fspgroups) {
                                                            try { WatchCatchableExitSignal } catch { }

                                                            $sid = (New-Object System.Security.Principal.SecurityIdentifier($group.properties.objectsid[0], 0)).value
                                                            Write-Verbose "          $($sid) (domain local group across trust)"
                                                            $GroupsSIDs += $sid

                                                            foreach ($SidHistorySid in @($group.properties.sidhistory | Where-Object { $_ })) {
                                                                $sid = (New-Object System.Security.Principal.SecurityIdentifier($SidHistorySid, 0)).value
                                                                Write-Verbose "          $($sid) (domain local group sIDHistory across trust)"
                                                                $GroupsSIDs += $sid
                                                            }
                                                        }
                                                    } else {
                                                        Write-Verbose '          FSP is not member of any group'
                                                    }
                                                } catch {
                                                    Write-Host "          Error: $($error[0].exception)" -ForegroundColor red
                                                }
                                            } else {
                                                Write-Verbose "          Ignoring, because '$($fsp.path)' is not part of a trust in TrustsToCheckForGroups."
                                            }
                                        }
                                    }
                                } else {
                                    Write-Verbose '          No ForeignSecurityPrincipal(s) found'
                                }
                            }
                        }
                    } else {
                        try {
                            try { WatchCatchableExitSignal } catch { }

                            $tempX = GraphGetUserTransitiveMemberOf $ADPropsMailboxes[$AccountNumberRunning].userPrincipalName

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($sid in @($tempX.memberof.value.securityidentifier | Where-Object { $_ })) {
                                $GroupsSIDs += $sid
                                Write-Verbose "        $($sid) (Entra ID group)"
                            }

                            try { WatchCatchableExitSignal } catch { }

                            foreach ($sid in @($tempX.memberof.value.onpremisessecurityidentifier | Where-Object { $_ })) {
                                $GroupsSIDs += $sid
                                Write-Verbose "        $($sid) (on-prem group)"
                            }

                            $tempX = $null
                        } catch {
                            Write-Host '        Skipping, mailbox not found in Microsoft Graph.' -ForegroundColor yellow
                        }
                    }
                } else {
                    Write-Host '        Skipping, as mailbox could not be found in your environment in an earlier step.' -ForegroundColor yellow
                }

                $ADPropsMailboxes[$AccountNumberRunning] | Add-Member -MemberType NoteProperty -Name 'GroupsSIDs' -Value $GroupsSIDs -Force

                if ($ADPropsCurrentUser.proxyaddresses -icontains "smtp:$($MailAddresses[$AccountNumberRunning])") {
                    $ADPropsCurrentUser = $ADPropsMailboxes[$AccountNumberRunning]
                }
            } else {
                $ADPropsMailboxes[$AccountNumberRunning] = $null
                $ADPropsMailboxManagers[$AccountNumberRunning] = $null
            }
        } else {
            Write-Host "    Mailbox user object already found before, using cached data of $($MailAddresses[$CurrentMailboxAlreadyFoundFirstIndex])"

            $ADPropsMailboxes[$AccountNumberRunning] = $ADPropsMailboxes[$CurrentMailboxAlreadyFoundFirstIndex]
            $ADPropsMailboxManagers[$AccountNumberRunning] = $ADPropsMailboxManagers[$CurrentMailboxAlreadyFoundFirstIndex]
        }

        if ($AccountNumberRunning -eq ($MailAddresses.count - 1)) {
            if ($VirtualMailboxConfigFile) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DefineAndAddVirtualMailboxes')))) {
                    Write-Host '  Virtual mailboxes and dynamic signature INI entries cannot be defined and added.' -ForegroundColor Green
                    Write-Host "  The 'VirtualMailboxConfigFile' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                    Write-Host '  Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DefineAndAddVirtualMailboxes()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '  Error defining and adding virtual mailboxes.' -ForegroundColor Yellow
                        Write-Host "  $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "  Parameter 'VirtualMailboxConfigFile' is not enabled, skipping task."
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Sort mailbox list: User's primary mailbox, mailboxes in default Outlook profile, others @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    # Get users primary mailbox
    $p = $null

    # First, check if the user has a mail attribute set
    if ($ADPropsCurrentUser.mail) {
        Write-Host "  Mail attribute of currently logged-in or simulated user: '$($ADPropsCurrentUser.mail)'"

        for ($i = 0; $i -lt $LegacyExchangeDNs.count; $i++) {
            # if (($LegacyExchangeDNs[$i]) -and (($ADPropsMailboxes[$i].proxyaddresses) -icontains "smtp:$($ADPropsCurrentUser.mail)")) {
            if ((($ADPropsMailboxes[$i].proxyaddresses) -icontains "smtp:$($ADPropsCurrentUser.mail)")) {
                if (($SimulateUser) -or ((-not $SimulateUser) -and ($LegacyExchangeDNs[$i]))) {
                    $p = $i
                    break
                }
            }
        }

        if ($p -ge 0) {
            Write-Host '    Matching mailbox found'
        } else {
            Write-Host '    No matching mailbox found, see prior warning messages for details' -ForegroundColor Yellow
        }
    } else {
        Write-Host '  AD mail attribute of currently logged-in user is empty'

        if ($null -ne $TrustsToCheckForGroups[0]) {
            Write-Host '    Searching msExchMasterAccountSid'
            # No mail attribute set, check for match(es) of user's objectSID and mailbox's msExchMasterAccountSid
            for ($i = 0; $i -lt $MailAddresses.count; $i++) {
                if ($ADPropsMailboxes[$i].msexchmasteraccountsid) {
                    try { WatchCatchableExitSignal } catch { }

                    if ((New-Object System.Security.Principal.SecurityIdentifier($ADPropsMailboxes[$i].msexchmasteraccountsid[0], 0)).value -iin $CurrentUserSIDs) {
                        if ($p -ge 0) {
                            # $p already set before, there must be at least two matches, so set it to -1
                            $p = -1
                        } elseif ((-not $p) -and ($RegistryPaths[$i] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*')) {
                            $p = $i
                        }
                    }
                }
            }

            if ($p -ge 0) {
                Write-Host "    One matching primary mailbox found: $MailAddresses[$i]"
            } elseif ($null -eq $p) {
                Write-Host '    No matching primary mailbox found' -ForegroundColor Yellow
            } else {
                Write-Host '    Multiple matching primary mailboxes found, no prioritization possible' -ForegroundColor Yellow
            }
        } else {
            Write-Host
        }
    }

    Write-Host '  Mailbox priority (highest to lowest)'
    $MailboxNewOrder = @()
    $PrimaryMailboxAddress = $null

    if ($p -ge 0) {
        $MailboxNewOrder += $p
        $PrimaryMailboxAddress = $MailAddresses[$p]
    }

    if ((-not $SimulateUser) -and ($OutlookProfiles.count -gt 0)) {
        foreach ($OutlookProfile in $OutlookProfiles) {
            $MailAddressesToSearch = @()
            $MailAddressesToSearchLookup = @{}
            for ($count = 0; $count -lt $RegistryPaths.count; $count++) {
                if ($MailAddresses[$count] -and (($RegistryPaths[$count]).StartsWith("Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\"))) {
                    $MailAddressesToSearch += $MailAddresses[$count]
                    $MailAddressesToSearchLookup[$($MailAddresses[$count])] = $MailAddresses[$count]

                    foreach ($ProxyAddress in $ADPropsMailboxes[$count].proxyaddresses) {
                        if ([string]$ProxyAddress -ilike 'smtp:*') {
                            $MailAddressesToSearch += $([string]$ProxyAddress -ireplace 'smtp:', '')
                            $MailAddressesToSearchLookup[$([string]$ProxyAddress -ireplace 'smtp:', '')] = $MailAddresses[$count]
                        }
                    }
                }
            }

            $CurrentOutlookProfileMailboxSortOrder = @()

            foreach ($RegistryFolder in @(Get-ItemProperty -LiteralPath "hkcu:\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$($OutlookProfile)\0a0d020000000000c000000000000046" -ErrorAction SilentlyContinue | Where-Object { ($_.'11020458') })) {
                try { WatchCatchableExitSignal } catch { }

                try {
                    @(@(([regex]::Matches((@(foreach ($char in @(($RegistryFolder.'11020458' -join ',').Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -gt '0' })) { [char][int]"$($char)" }) -join ''), (@(@($MailAddressesToSearch) | ForEach-Object { [Regex]::Escape($_) }) -join '|'), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).captures.value).tolower()) | Select-Object -Unique) | ForEach-Object {
                        $CurrentOutlookProfileMailboxSortOrder += $MailAddressesToSearchLookup[$_]
                    }
                } catch {
                }
            }

            if (($CurrentOutlookProfileMailboxSortOrder.count -gt 0) -and ($CurrentOutlookProfileMailboxSortOrder.count -eq (@($RegistryPaths | Where-Object { $_.startswith("Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\") }).count))) {
                Write-Verbose '  Outlook mailbox display sort order is defined and contains all found mail addresses.'
                foreach ($CurrentOutlookProfileMailboxSortOrderMailbox in $CurrentOutlookProfileMailboxSortOrder) {
                    for ($i = 0; $i -le $RegistryPaths.count - 1; $i++) {
                        try { WatchCatchableExitSignal } catch { }

                        if ((($RegistryPaths[$i]).startswith("Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\")) -and ($i -ne $p)) {
                            if ($MailAddresses[$i] -ieq $CurrentOutlookProfileMailboxSortOrderMailbox) {
                                $MailboxNewOrder += $i
                                break
                            }
                        }
                    }
                }
            } else {
                for ($i = 0; $i -le $RegistryPaths.count - 1; $i++) {
                    try { WatchCatchableExitSignal } catch { }

                    if ((($RegistryPaths[$i]).startswith("Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\$($OutlookRegistryVersion)\Outlook\Profiles\$OutlookProfile\")) -and ($i -ne $p)) {
                        $MailboxNewOrder += $i
                    }
                }
            }

        }
    }

    for ($i = 0; $i -lt $MailAddresses.Count; $i++) {
        if ($MailboxNewOrder -inotcontains $i ) {
            $MailboxNewOrder += $i
        }
    }

    foreach ($VariableName in ('RegistryPaths', 'MailAddresses', 'LegacyExchangeDNs', 'ADPropsMailboxesUserDomain', 'ADPropsMailboxes')) {
        (Get-Variable -Name $VariableName).value = (Get-Variable -Name $VariableName).value[$MailboxNewOrder]
    }

    for ($x = 0; $x -lt $MailAddresses.count; $x++) {
        if ($MailAddresses.IndexOf($MailAddresses[$x]) -eq $x) {
            Write-Host "    $($MailAddresses[$x])"

            $y = 0

            @(
                foreach ($MailAddress in $MailAddresses) {
                    if ($MailAddress -ieq $MailAddresses[$x]) {
                        $y
                    }
                    $y++
                }
            ) | ForEach-Object {
                Write-Verbose "      Outlook profile '$(($RegistryPaths[$_] -split '\\')[8])'"
                Write-Verbose "        Registry: $($RegistryPaths[$_] -ireplace '^Microsoft\.PowerShell\.Core\\Registry::HKEY_CURRENT_USER', 'HKCU')"
                Write-Verbose "        LegacyExchangeDN: $($LegacyExchangeDNs[$_])"
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    $TemplateFilesGroupSIDsOverall = @{}

    foreach ($SigOrOOF in ('signature', 'OOF')) {
        if (($SigOrOOF -eq 'OOF') -and ($SetCurrentUserOOFMessage -eq $false)) {
            break
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host
        Write-Host "Get all $($SigOrOOF) template files and categorize them @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        $TemplateFilesCommon = @{}
        $TemplateFilesGroup = @{}
        $TemplateFilesGroupFilePart = @{}
        $TemplateFilesMailbox = @{}
        $TemplateFilesMailboxFilePart = @{}
        $TemplateFilesReplacementvariable = @{}
        $TemplateFilesReplacementvariableFilePart = @{}
        $TemplateFilesDefaultnewOrInternal = @{}
        $TemplateFilesDefaultreplyfwdOrExternal = @{}
        $TemplateFilesWriteProtect = @{}

        $TemplateTemplatePath = Get-Variable -Name "$($SigOrOOF)TemplatePath" -ValueOnly
        $TemplateIniFile = Get-Variable -Name "$($SigOrOOF)IniFile" -ValueOnly
        $TemplateIniSettings = Get-Variable -Name "$($SigOrOOF)IniSettings" -ValueOnly

        # Remove trailing null character from file names being enumerated in SharePoint folders. .Net or the WebDAV client sometimes add a null character, which is not allowed in file and path names.
        ## Original code:
        ## $TemplateFiles = @((Get-ChildItem $TemplateTemplatePath -File -Filter $(if ($UseHtmTemplates) { '*.htm' } else { '*.docx' })) | Sort-Object -Culture 127)
        $TemplateFiles = @(@(@(@(Get-ChildItem -LiteralPath $TemplateTemplatePath -File) | Where-Object { $_.Extension -iin $(if ($UseHtmTemplates) { @('.htm', ".htm$([char]0)") } else { @('*.docx', ".docx$([char]0)") }) }) | Select-Object -Property @{n = 'FullName'; e = { $_.FullName.ToString() -ireplace '\x00$', '' } }, @{n = 'Name'; Expression = { $_.Name.ToString() -ireplace '\x00$', '' } }) | Sort-Object -Culture 127 -Property FullName, Name)

        if ($TemplateIniFile -ne '') {
            Write-Host "  Compare $($SigOrOOF) INI entries and file system"
            foreach ($Enumerator in $TemplateIniSettings.GetEnumerator().name) {
                try { WatchCatchableExitSignal } catch { }

                if ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>']) {
                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotin $TemplateFiles.name)) {
                        Write-Host "    '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' ($($SigOrOOF) INI index #$($Enumerator)) found in INI but not in signature template path." -ForegroundColor Yellow
                    }

                    if (($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -ine '<Set-OutlookSignatures configuration>') -and ($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'] -inotlike "*.$(if($UseHtmTemplates){'htm'} else {'docx'})")) {
                        Write-Host "    '$($TemplateIniSettings[$Enumerator]['<Set-OutlookSignatures template>'])' ($($SigOrOOF) INI index #$($Enumerator)) has the wrong file extension ('-UseHtmTemplates true' allows .htm, else .docx)" -ForegroundColor Yellow
                    }
                }
            }

            $x = @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })

            foreach ($TemplateFile in $TemplateFiles) {
                if ($TemplateFile.name -inotin $x) {
                    Write-Host "    '$($TemplateFile.name)' found in $($SigOrOOF) template path but not in INI file." -ForegroundColor Yellow
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host '  Sort template files according to configuration'
            $TemplateFilesSortCulture = (@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortCulture']

            # Populate template files in the most complicated way first: SortOrder 'AsInThisFile'
            # This also considers that templates can be referenced multiple times in the INI file
            # If the setting in the INI file is different, we only need to sort $TemplateFiles
            $TemplateFilesExisting = @(foreach ($Enumerator in $TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)]) { $Enumerator['<Set-OutlookSignatures template>'] })
            $TemplateFiles = @($TemplateFiles | Where-Object { $_.name -iin $TemplateFilesExisting })
            $TemplateFiles | Add-Member -MemberType NoteProperty -Name TemplateIniSettingsIndex -Value $null -Force
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

                $TemplateFiles = @($TemplateFiles[$TemplateFilesSortOrder] | Select-Object -Property fullname, name, TemplateIniSettingsIndex)

                if ($TemplateFiles.count -gt 0) {
                    foreach ($index in 0..($TemplateFiles.Count - 1)) {
                        $TemplateFiles[$index].TemplateIniSettingsIndex = $TemplateFilesIniIndex[$index]
                    }
                }

                if (($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' })) {
                    switch ((@($TemplateIniSettings[($TemplateIniSettings.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1)['SortOrder']) {
                        { $_ -iin ('AsInThisFile', 'AsListed') } {
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
                            # same as 'AsInThisFile'
                            # nothing to do, $TemplateFiles is already correctly populated and sorted
                        }
                    }
                } else {
                    $TemplateFiles = @($TemplateFiles | Sort-Object -Culture $TemplateFilesSortCulture -Property Name, @{expression = { [int]$_.TemplateIniSettingsIndex } })
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        foreach ($TemplateFile in $TemplateFiles) {
            try { WatchCatchableExitSignal } catch { }

            $TemplateIniSettingsIndex = $TemplateFile.TemplateIniSettingsIndex
            $TemplateFileGroupSIDs = @{}
            Write-Host ("    '$($TemplateFile.Name)' ($($SigOrOOF) INI index #$($TemplateIniSettingsIndex))")

            if ($TemplateIniSettings[$TemplateIniSettingsIndex]['<Set-OutlookSignatures template>'] -ieq $TemplateFile.name) {
                $TemplateFilePart = (@(@($TemplateIniSettings[$TemplateIniSettingsIndex].GetEnumerator().Name) | Sort-Object -Culture 127) -join '] [')
                if ($TemplateFilePart) {
                    $TemplateFilePart = ($TemplateFilePart -split '\] \[' | Where-Object { $_ -inotin ('OutlookSignatureName', '<Set-OutlookSignatures template>') }) -join '] ['
                    $TemplateFilePart = '[' + $TemplateFilePart + ']'
                    $TemplateFilePart = $TemplateFilePart -ireplace '\[\]', ''
                }

                if ($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']) {
                    Write-Host "      Outlook signature name: '$($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'])'"

                    if ((CheckFilenamePossiblyInvalid -Filename $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'])) {
                        # Write-Host "        Ignore INI entry, signature name is invalid: $((CheckFilenamePossiblyInvalid -Filename $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']))" -ForegroundColor Yellow
                        # Continue

                        Write-Host "        Signature name has invalid characters. Replacing: $((CheckFilenamePossiblyInvalid -Filename $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']))" -ForegroundColor Yellow

                        $tempOutlookSignatureName = $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']

                        @(
                            (CheckFilenamePossiblyInvalid -Filename $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName']) -split [regex]::Escape(', ')
                        ) | ForEach-Object {
                            $tempOutlookSignatureName = $tempOutlookSignatureName -ireplace [regex]::Escape($_), $(if ($_ -eq '@') { '_at_' } else { '_' })
                        }

                        Write-Host "          '$($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'])' -> '$($tempOutlookSignatureName)'" -ForegroundColor Yellow

                        $TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'] = $tempOutlookSignatureName
                    }

                    $TemplateFileTargetName = ($TemplateIniSettings[$TemplateIniSettingsIndex]['OutlookSignatureName'] + $(if ($UseHtmTemplates) { '.htm' } else { '.docx' }))
                } else {
                    if ((CheckFilenamePossiblyInvalid -Filename $TemplateFile.Name)) {
                        # Write-Host "      Ignore INI entry, signature name is invalid: $((CheckFilenamePossiblyInvalid -Filename $TemplateFile.Name))" -ForegroundColor Yellow
                        # continue

                        Write-Host "        Signature name has invalid characters. Replacing: $((CheckFilenamePossiblyInvalid -Filename $TemplateFile.Name))" -ForegroundColor Yellow

                        $tempOutlookSignatureName = $TemplateFile.Name

                        @(
                            (CheckFilenamePossiblyInvalid -Filename $TemplateFile.Name) -split [regex]::Escape(', ')
                        ) | ForEach-Object {
                            $tempOutlookSignatureName = $tempOutlookSignatureName -ireplace [regex]::Escape($_), $(if ($_ -eq '@') { '_at_' } else { '_' })
                        }

                        Write-Host "          '$($TemplateFile.Name)' -> '$($tempOutlookSignatureName)'" -ForegroundColor Yellow

                        $TemplateFileTargetName = $tempOutlookSignatureName
                    } else {
                        $TemplateFileTargetName = $TemplateFile.Name
                    }
                }
            } else {
                $TemplateFilePart = ''
                $TemplateFileTargetName = $TemplateFile.Name
            }

            $TemplateFilePartRegexTimeAllow = '\[(?!-:)\d{12}Z?-\d{12}Z?\]'
            $TemplateFilePartRegexTimeDeny = '\[-:\d{12}Z?-\d{12}Z?\]'
            $TemplateFilePartRegexGroupAllow = '(?i)\[(?!-:|-CURRENTUSER:)\S+?(?<!]) .+?\]'
            $TemplateFilePartRegexGroupDeny = '(?i)\[(-:|-CURRENTUSER:)\S+?(?<!]) .+?\]'
            $TemplateFilePartRegexMailaddressAllow = '(?i)\[(?!-:|-CURRENTUSER:)(\S+?)@(\S+?)\.(\S+?)\]'
            $TemplateFilePartRegexMailaddressDeny = '(?i)\[(-:|-CURRENTUSER:)(\S+?)@(\S+?)\.(\S+?)\]'
            $TemplateFilePartRegexReplacementvariableAllow = '(?i)\[(?!-:)\$.*\$\]'
            $TemplateFilePartRegexReplacementvariableDeny = '(?i)\[(-:)\$.*\$\]'

            if ($SigOrOOF -ieq 'signature') {
                $TemplateFilePartRegexDefaultneworinternal = '(?i)\[DefaultNew\]'
                $TemplateFilePartRegexDefaultreplyfwdorexternal = '(?i)\[DefaultReplyFwd\]'
                $TemplateFilePartRegexWriteprotect = '(?i)\[WriteProtect\]'
            } else {
                $TemplateFilePartRegexDefaultneworinternal = '(?i)\[internal\]'
                $TemplateFilePartRegexDefaultreplyfwdorexternal = '(?i)\[external\]'
                $TemplateFilePartRegexWriteprotect = ''
            }

            $TemplateFilePartRegexKnown = '(' + (($TemplateFilePartRegexTimeAllow, $TemplateFilePartRegexTimeDeny, $TemplateFilePartRegexGroupAllow, $TemplateFilePartRegexGroupDeny, $TemplateFilePartRegexMailaddressAllow, $TemplateFilePartRegexMailaddressDeny, $TemplateFilePartRegexReplacementvariableAllow, $TemplateFilePartRegexReplacementvariableDeny, $TemplateFilePartRegexDefaultneworinternal, $TemplateFilePartRegexDefaultreplyfwdorexternal, $TemplateFilePartRegexWriteprotect) -join '|') + ')'

            try { WatchCatchableExitSignal } catch { }

            # time based template
            $TemplateFileTimeActive = $true
            if (($TemplateFilePart -imatch $TemplateFilePartRegexTimeAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexTimeDeny)) {
                Write-Host '      Time based template'
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('TimeBasedTemplate')))) {
                    Write-Host '        Templates cannot be activated or deactivated for specified time ranges.' -ForegroundColor Green
                    Write-Host "        The 'time based template' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                    Write-Host '        Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                } else {
                    try { WatchCatchableExitSignal } catch { }
                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::TimeBasedTemplate()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '        Error evaluating time based templates.' -ForegroundColor Yellow
                        Write-Host "        $FeatureResult" -ForegroundColor Yellow
                    }
                }
            }

            if ($TemplateFileTimeActive -ne $true) {
                continue
            }

            try { WatchCatchableExitSignal } catch { }

            # common template
            if (($TemplateFilePart -inotmatch $TemplateFilePartRegexGroupAllow) -and ($TemplateFilePart -inotmatch $TemplateFilePartRegexMailaddressAllow) -and ($TemplateFilePart -inotmatch $TemplateFilePartRegexReplacementvariableAllow)) {
                Write-Host '      Common template (no group, email address or replacement variable allow tags specified)'
                if (-not $TemplateFilesCommon.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesCommon.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesCommon[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                }

                $TemplateClassificationDisplayOrder = ('group', 'mail', 'replacementvariable')
            } elseif ($TemplateFilePart -imatch $TemplateFilePartRegexGroupAllow) {
                $TemplateClassificationDisplayOrder = ('group', 'mail', 'replacementvariable')
            } elseif ($TemplateFilePart -imatch $TemplateFilePartRegexMailaddressAllow) {
                $TemplateClassificationDisplayOrder = ('mail', 'group', 'replacementvariable')
            } elseif ($TemplateFilePart -imatch $TemplateFilePartRegexReplacementvariableAllow) {
                $TemplateClassificationDisplayOrder = ('replacementvariable', 'group', 'mail')
            }

            foreach ($TemplateClassificationDisplayOrderEntry in $TemplateClassificationDisplayOrder) {
                try { WatchCatchableExitSignal } catch { }

                # group specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'group') {
                    if (($TemplateFilePart -imatch $TemplateFilePartRegexGroupAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexGroupDeny)) {
                        if (-not $TemplateFilesGroup.ContainsKey($TemplateIniSettingsIndex)) {
                            $TemplateFilesGroup.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesGroup[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                        }

                        $InclusionCount = $null
                        $ExclusionCount = $null

                        foreach ($TemplateFilePartTag in @(@(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexGroupAllow).captures.value) + @([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexGroupDeny).captures.value)) | Where-Object { $_ })) {
                            if (($TemplateFilePartTag -imatch $TemplateFilePartRegexGroupAllow) -and ($null -eq $InclusionCount)) {
                                Write-Host '      Group specific template'
                                $InclusionCount++
                            } elseif (($TemplateFilePartTag -imatch $TemplateFilePartRegexGroupDeny) -and ($null -eq $ExclusionCount)) {
                                Write-Host '      Group specific exclusions'
                                $ExclusionCount++
                            }

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                            $NTName = $TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|)(.*?) (.*)(\])$', '${3}\${4}'

                            # Check cache
                            #   $TemplateFilesGroupSIDsOverall contains tags without prefix only: [xxx xxx]
                            #   $TemplateFilesGroupSIDsOverall contains tag with extracted prefix: -:[xxx xxx]

                            if ($TemplateFilesGroupSIDsOverall.ContainsKey($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${1}${3}'))) {
                                $TemplateFileGroupSIDs.add($TemplateFilePartTag, "$($TemplateFilePartTag -ireplace '(?i)(^\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${2}')$($TemplateFilesGroupSIDsOverall[$($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${1}${3}')])")
                            }

                            if ((-not $TemplateFileGroupSIDs.ContainsKey($TemplateFilePartTag))) {
                                $tempSid = ResolveToSid($NTName)

                                if ($tempSid) {
                                    $TemplateFilesGroupSIDsOverall.add($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${1}${3}'), $tempSid)
                                    $TemplateFileGroupSIDs.add($TemplateFilePartTag, "$($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${2}')$($TemplateFilesGroupSIDsOverall[$($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${1}${3}')])")
                                }
                            }

                            if ($TemplateFileGroupSIDs.containskey($TemplateFilePartTag)) {
                                if ($null -ne $TemplateFileGroupSIDs[$TemplateFilePartTag]) {
                                    Write-Host "          $($TemplateFileGroupSIDs[$TemplateFilePartTag] -ireplace '(?i)^(-:|-CURRENTUSER:|CURRENTUSER:|)', '')"
                                    $TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesGroupFilePart[$TemplateIniSettingsIndex] + '[' + $TemplateFileGroupSIDs[$TemplateFilePartTag] + ']')
                                } else {
                                    Write-Host '          Not found' -ForegroundColor Yellow
                                }
                            } else {
                                Write-Host '          Not found' -ForegroundColor Yellow
                                $TemplateFilesGroupSIDsOverall.add($($TemplateFilePartTag -ireplace '(?i)^(\[)(-:|-CURRENTUSER:|CURRENTUSER:|)(.*)', '${1}${3}'), $null)
                            }
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # mailbox specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'mail') {
                    if (($TemplateFilePart -imatch $TemplateFilePartRegexMailaddressAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexMailaddressDeny)) {
                        if (-not $TemplateFilesMailbox.ContainsKey($TemplateIniSettingsIndex)) {
                            $TemplateFilesMailbox.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesMailbox[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                        }

                        $InclusionCount = $null
                        $ExclusionCount = $null

                        foreach ($TemplateFilePartTag in @(@(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexMailaddressAllow).captures.value) + @([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexMailaddressDeny).captures.value)) | Where-Object { $_ })) {
                            if (($TemplateFilePartTag -imatch $TemplateFilePartRegexMailaddressAllow) -and ($null -eq $InclusionCount)) {
                                Write-Host '      Mailbox specific template'
                                $InclusionCount++
                            } elseif (($TemplateFilePartTag -imatch $TemplateFilePartRegexMailaddressDeny) -and ($null -eq $ExclusionCount)) {
                                Write-Host '      Mailbox specific exclusions'
                                $ExclusionCount++
                            }

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                            $TemplateFilesMailboxFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesMailboxFilePart[$TemplateIniSettingsIndex] + $TemplateFilePartTag)
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                # Replacement variable specific template
                if ($TemplateClassificationDisplayOrderEntry -ieq 'replacementvariable') {
                    if (($TemplateFilePart -imatch $TemplateFilePartRegexReplacementvariableAllow) -or ($TemplateFilePart -imatch $TemplateFilePartRegexReplacementvariableDeny)) {
                        if (-not $TemplateFilesReplacementvariable.ContainsKey($TemplateIniSettingsIndex)) {
                            $TemplateFilesReplacementvariable.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesReplacementvariable[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                        }

                        $InclusionCount = $null
                        $ExclusionCount = $null

                        foreach ($TemplateFilePartTag in @(@(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexReplacementvariableAllow).captures.value) + @([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexReplacementvariableDeny).captures.value)) | Where-Object { $_ })) {
                            if (($TemplateFilePartTag -imatch $TemplateFilePartRegexReplacementvariableAllow) -and ($null -eq $InclusionCount)) {
                                Write-Host '      Replacement variable specific template'
                                $InclusionCount++
                            } elseif (($TemplateFilePartTag -imatch $TemplateFilePartRegexReplacementvariableDeny) -and ($null -eq $ExclusionCount)) {
                                Write-Host '      Replacement variable exclusions'
                                $ExclusionCount++
                            }

                            Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                            $TemplateFilesReplacementvariableFilePart[$TemplateIniSettingsIndex] = ($TemplateFilesReplacementvariableFilePart[$TemplateIniSettingsIndex] + $TemplateFilePartTag)
                        }
                    }
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # DefaultNew, DefaultReplyFwd, Internal, External
            if ($TemplateFilePart -imatch $TemplateFilePartRegexDefaultneworinternal) {
                foreach ($TemplateFilePartTag in @(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexDefaultneworinternal).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '      Default signature for new emails'
                    } else {
                        Write-Host '      Default internal OOF message'
                    }

                    Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                }

                if (-not $TemplateFilesDefaultnewOrInternal.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesDefaultnewOrInternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultnewOrInternal[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                }
            }

            if ($TemplateFilePart -imatch $TemplateFilePartRegexDefaultreplyfwdorexternal) {
                foreach ($TemplateFilePartTag in @(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexDefaultreplyfwdorexternal).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '      Default signature for replies and forwards'
                    } else {
                        Write-Host '      Default external OOF message'
                    }

                    Write-Host "        $(($TemplateFilePartTag -ireplace '^\[', '') -ireplace '\]$', '')"
                }

                if (-not $TemplateFilesDefaultreplyfwdOrExternal.containskey($TemplateIniSettingsIndex)) {
                    $TemplateFilesDefaultreplyfwdOrExternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultreplyfwdOrExternal[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                }
            }

            if ($SigOrOOF -ieq 'OOF') {
                if (($TemplateFilePart -notmatch $TemplateFilePartRegexDefaultreplyfwdorexternal) -and ($TemplateFilePart -notmatch $TemplateFilePartRegexDefaultneworinternal)) {
                    Write-Host '      Default internal OOF message (neither internal nor external tag specified)'
                    $TemplateFilesDefaultnewOrInternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultnewOrInternal[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)

                    Write-Host '      Default external OOF message (neither internal nor external tag specified)'
                    $TemplateFilesDefaultreplyfwdOrExternal.add($TemplateIniSettingsIndex, @{})
                    $TemplateFilesDefaultreplyfwdOrExternal[$TemplateIniSettingsIndex].add($TemplateFile.FullName, $TemplateFileTargetName)
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # WriteProtect
            if ($TemplateFilePart -imatch $TemplateFilePartRegexWriteprotect) {
                foreach ($TemplateFilePartTag in @(@([regex]::Matches($TemplateFilePart, $TemplateFilePartRegexWriteprotect).captures.value) | Where-Object { $_ })) {
                    if ($SigOrOOF -ieq 'signature') {
                        Write-Host '      Signature will be write protected'
                        if (-not $TemplateFilesWriteProtect.containskey($TemplateIniSettingsIndex)) {
                            $TemplateFilesWriteProtect.add($TemplateIniSettingsIndex, @{})
                            $TemplateFilesWriteProtect[$TemplateIniSettingsIndex].add($TemplateFile.fullname, $TemplateFileTargetName)
                        }
                    }
                }

            }

            try { WatchCatchableExitSignal } catch { }

            # unknown tags
            $x = ($TemplateFilePart -ireplace $TemplateFilePartRegexKnown, '').trim()
            if ($x) {
                Write-Host '      Unknown tags' -ForegroundColor yellow
                Write-Host "        $(($x -ireplace '^\[', '') -ireplace '\]$', '')"
            }

            Set-Variable -Name "$($SigOrOOF)Files" -Value $TemplateFiles
            Set-Variable -Name "$($SigOrOOF)FilesCommon" -Value $TemplateFilesCommon
            Set-Variable -Name "$($SigOrOOF)FilesGroup" -Value $TemplateFilesGroup
            Set-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -Value $TemplateFilesGroupFilePart
            Set-Variable -Name "$($SigOrOOF)FilesMailbox" -Value $TemplateFilesMailbox
            Set-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -Value $TemplateFilesMailboxFilePart
            Set-Variable -Name "$($SigOrOOF)FilesReplacementvariable" -Value $TemplateFilesReplacementvariable
            Set-Variable -Name "$($SigOrOOF)FilesReplacementvariableFilePart" -Value $TemplateFilesReplacementvariableFilePart

            if ($SigOrOOF -ieq 'signature') {
                $SignatureFilesDefaultNew = $TemplateFilesDefaultnewOrInternal
                $SignatureFilesDefaultReplyFwd = $TemplateFilesDefaultreplyfwdOrExternal
                $SignatureFilesWriteProtect = $TemplateFilesWriteProtect
            } else {
                $OOFFilesInternal = $TemplateFilesDefaultnewOrInternal
                $OOFFilesExternal = $TemplateFilesDefaultreplyfwdOrExternal
            }
        }
    }


    try { WatchCatchableExitSignal } catch { }


    if ($macOSSignaturesScriptable) {
        Write-Host
        Write-Host "Create copies of Outlook for Mac signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        $SignaturePaths | ForEach-Object {
            try { WatchCatchableExitSignal } catch { }

            @(@(@"
tell application "Microsoft Outlook"
    set allSignatures to every signature

    repeat with aSignature in allSignatures
        set sigName to name of aSignature

        set sigContent to content of aSignature
        set fileName to sigName & ".htm"
        set filePath to "$($_)/" & fileName
        try
            log "  '" & fileName & "'"
            set fileRef to open for access POSIX file filePath with write permission as «class utf8»
            set eof of fileRef to 0
            write sigContent to fileRef as «class utf8»
            close access fileRef
        on error errorMessage
            log "    Error copying to '" & filepath & "': " & errorMessage
        end try

        set sigContent to plain text content of aSignature
        set fileName to sigName & ".txt"
        set filePath to "$($_)/" & fileName
        try
            log "  '" & filename & "'"
            set fileRef to open for access POSIX file filePath with write permission as «class utf8»
            set eof of fileRef to 0
            write sigContent to fileRef as «class utf8»
            close access fileRef
        on error errorMessage
            log "    Error copying to '" & filePath & "': " & errorMessage
        end try
    end repeat
end tell
"@ | osascript *>&1)) | ForEach-Object { Write-Host $_.tostring() }
        }
    }


    try { WatchCatchableExitSignal } catch { }


    Write-Host
    Write-Host "Start Word background process @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    if (($UseHtmTemplates -eq $true) -and (($CreateRtfSignatures -eq $false))) {
        Write-Host '  Not required: UseHtmTemplates = $true, CreateRtfSignatures = $false'
    } else {
        Write-Verbose "  WordProcessPriority: '$($WordProcessPriorityText)' ('$($WordProcessPriority)')"

        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class Win32Api
{
[System.Runtime.InteropServices.DllImportAttribute( "User32.dll", EntryPoint =  "GetWindowThreadProcessId" )]
public static extern int GetWindowThreadProcessId ( [System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, out int lpdwProcessId );

[DllImport("User32.dll", CharSet = CharSet.Auto)]
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
'@

        # Start Word dummy object, set process priority, start real Word object, set process priority, close dummy object - this seems to avoid a rare problem where a manually started Word instance connects to the Word process created by the software
        try {
            try { WatchCatchableExitSignal } catch { }

            Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value 0 -ErrorAction SilentlyContinue

            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $script:COMWordDummy = New-Object -ComObject Word.Application
            $VerbosePreference = $tempVerbosePreference
            $script:COMWordDummy.Visible = $false

            # Restore original Word AlertIfNotDefault setting
            Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null


            if ($script:COMWordDummy) {
                try { WatchCatchableExitSignal } catch { }

                # Set Word process priority
                $script:COMWordDummyCaption = $script:COMWordDummy.Caption
                $script:COMWordDummy.Caption = "Set-OutlookSignatures $([guid]::NewGuid())"
                $script:COMWordDummyHWND = [Win32Api]::FindWindow( 'OpusApp', $($script:COMWordDummy.Caption) )
                $script:COMWordDummyPid = [IntPtr]::Zero
                $null = [Win32Api]::GetWindowThreadProcessId( $script:COMWordDummyHWND, [ref] $script:COMWordDummyPid );
                $script:COMWordDummy.Caption = $script:COMWordDummyCaption
                try {
                    $((Get-Process -PID $script:COMWordDummyPid).PriorityClass = $WordProcessPriorityText)

                    if ((Get-Process -PID $script:COMWordDummyPid).PriorityClass.ToString() -ne $WordProcessPriorityText) {
                        throw "No error, but Word dummy process priority set to '$((Get-Process -PID $script:COMWordDummyPid).PriorityClass.ToString())' ('$((Get-Process -PID $script:COMWordDummyPid).PriorityClass.value__)') instead of '$($WordProcessPriorityText)' ('$($WordProcessPriority)')."
                    }
                } catch {
                    Write-Host "    Error setting Word dummy process priority: $($_)" -ForegroundColor Yellow
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value 0 -ErrorAction SilentlyContinue

            $tempVerbosePreference = $VerbosePreference
            $VerbosePreference = 'SilentlyContinue'
            $script:COMWord = New-Object -ComObject Word.Application
            $VerbosePreference = $tempVerbosePreference
            $script:COMWord.Visible = $false

            # Restore original Word AlertIfNotDefault setting
            Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null


            if ($script:COMWord) {
                try { WatchCatchableExitSignal } catch { }

                # Set Word process priority
                $script:COMWordCaption = $script:COMWord.Caption
                $script:COMWord.Caption = "Set-OutlookSignatures $([guid]::NewGuid())"
                $script:COMWordHWND = [Win32Api]::FindWindow( 'OpusApp', $($script:COMWord.Caption) )
                $script:COMWordPid = [IntPtr]::Zero
                $null = [Win32Api]::GetWindowThreadProcessId( $script:COMWordHWND, [ref] $script:COMWordPid );
                $script:COMWord.Caption = $script:COMWordCaption
                try {
                    $((Get-Process -PID $script:COMWordPid).PriorityClass = $WordProcessPriorityText)

                    if ((Get-Process -PID $script:COMWordPid).PriorityClass.ToString() -ne $WordProcessPriorityText) {
                        throw "No error, but Word process priority set to '$((Get-Process -PID $script:COMWordPid).PriorityClass.ToString())' ('$((Get-Process -PID $script:COMWordPid).PriorityClass.value__)') instead of '$($WordProcessPriorityText)' ('$($WordProcessPriority)')."
                    }
                } catch {
                    Write-Host "    Error setting Word process priority: $($_)" -ForegroundColor Yellow
                }

                # Open blank document and get the default view value
                $script:ComWord.Documents.Add([type]::Missing, $false, [type]::Missing, [type]::Missing) | Out-Null
                $script:COMWordViewTypeOriginal = $script:COMWord.ActiveDocument.ActiveWindow.View.Type
                $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)
            }

            if ($script:COMWordDummy) {
                $script:COMWordDummy.Quit([ref]$false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWordDummy) | Out-Null
                Remove-Variable COMWordDummy -Scope 'script'
            }

            try { WatchCatchableExitSignal } catch { }

            Add-Type -LiteralPath (Get-ChildItem -LiteralPath ((Join-Path -Path ($env:SystemRoot) -ChildPath 'assembly\GAC_MSIL\Microsoft.Office.Interop.Word')) -Filter 'Microsoft.Office.Interop.Word.dll' -Recurse | Select-Object -ExpandProperty FullName -Last 1)
        } catch {
            Write-Host $error[0]
            Write-Host '  Word not installed or not working correctly. Install or repair Word and the registry information about Word, or consider using HTM templates instead of DOCX templates. Exit.' -ForegroundColor Red

            # Restore original Word AlertIfNotDefault setting
            Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null

            $script:ExitCode = 17
            $script:ExitCodeDescription = 'Word not installed or not working correctly.'
            exit
        }
    }


    # Process each email address only once
    $script:SignatureFilesDone = @()

    if ($SimulateUser) {
        try { WatchCatchableExitSignal } catch { }

        Get-ChildItem (Join-Path -Path ($SignaturePaths[0]) -ChildPath '___Mailbox *') -Attributes Directory -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ForEach-Object {
            try { WatchCatchableExitSignal } catch { }

            RemoveItemAlternativeRecurse $($_.FullName)
        }
    }

    for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
        try { WatchCatchableExitSignal } catch { }

        if (($AccountNumberRunning -eq $MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])) -and ($($MailAddresses[$AccountNumberRunning]) -like '*@*')) {
            Write-Host
            Write-Host "Mailbox $($MailAddresses[$AccountNumberRunning]) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            $UserDomain = ''
            $GroupsSIDs = @()
            $ADPropsCurrentMailbox = @()
            $ADPropsCurrentMailboxManager = @()

            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                $ADPropsCurrentMailbox = $ADPropsMailboxes[$AccountNumberRunning]
                $ADPropsCurrentMailboxManager = $ADPropsMailboxManagers[$AccountNumberRunning]
                $GroupsSIDs = $ADPropsMailboxes[$AccountNumberRunning].GroupsSIDs
            }


            if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('CLCGM')))) {
                Write-Host '  Mailbox is member of license group: False (no valid Benefactor Circle license file found)'
            } else {
                try { WatchCatchableExitSignal } catch { }

                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::CLCGM()

                if ($FeatureResult -ine 'true') {
                    Write-Host "  Mailbox is member of license group: False ($($FeatureResult))"
                } else {
                    Write-Host '  Mailbox is member of license group: True'
                }
            }


            try { WatchCatchableExitSignal } catch { }


            Write-Host "  Extract SMTP addresses @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            $CurrentMailboxSMTPAddresses = @()
            if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
                foreach ($ProxyAddress in $ADPropsCurrentMailbox.proxyaddresses) {
                    if ([string]$ProxyAddress -ilike 'smtp:*') {
                        $CurrentMailboxSmtpaddresses += [string]$ProxyAddress -ireplace 'smtp:', ''
                        Write-Verbose "    $($CurrentMailboxSMTPAddresses[-1])"
                    }
                }
            } else {
                $CurrentMailboxSmtpaddresses += $($MailAddresses[$AccountNumberRunning])
                Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox.' -ForegroundColor Yellow
                Write-Host "    Using '$($MailAddresses[$AccountNumberRunning])' as single known SMTP address." -ForegroundColor Yellow
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "  Calculate replacement variables @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            $ReplaceHash = @{}

            if (Test-Path -LiteralPath $ReplacementVariableConfigFile -PathType Leaf) {
                try {
                    Write-Host "    '$ReplacementVariableConfigFile'"
                    . ([System.Management.Automation.ScriptBlock]::Create((ConvertEncoding -InFile $ReplacementVariableConfigFile -InIsHtml $false)))
                } catch {
                    Write-Host $error[0]
                    Write-Host "    Problem executing content of '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                    $script:ExitCode = 18
                    $script:ExitCodeDescription = 'Problem executing content of ReplacementVariableConfigFile.'
                    exit
                }
            } else {
                Write-Host "    Problem connecting or reading '$ReplacementVariableConfigFile'. Exit." -ForegroundColor Red
                $script:ExitCode = 19
                $script:ExitCodeDescription = 'Problem connecting or reading ReplacementVariableConfigFile.'
                exit
            }

            try { WatchCatchableExitSignal } catch { }

            $PictureVariablesArray = @()

            foreach ($VariableName in @(foreach ($VariableName in @(
                            @(
                                foreach ($ReplacementVariableScope in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
                                    @(1..10) | ForEach-Object { "`$$($ReplacementVariableScope)CustomImage$($_)`$" }
                                }
                            ) +
                            @('$CurrentMailboxManagerPhoto$', '$CurrentMailboxPhoto$', '$CurrentUserManagerPhoto$', '$CurrentUserPhoto$')
                        )
                    ) {
                        $VariableName
                    }
                )
            ) {
                try { WatchCatchableExitSignal } catch { }

                New-Variable -Name $($($VariableName).Trim('$') + 'Guid') -Value (New-Guid).Guid -Force

                $PictureVariablesArray += , @($VariableName, $(Get-Variable -Name $($VariableName.Trim('$') + 'Guid') -ValueOnly))
            }

            foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture 127)) {
                try { WatchCatchableExitSignal } catch { }

                if ($replaceKey -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' })) {
                    Write-Verbose "    $($replaceKey): '$($replaceHash[$replaceKey])'"
                } else {
                    if ($null -ne $($replaceHash[$replaceKey])) {
                        Write-Verbose "    $($replaceKey): Photo available, $([math]::ceiling($($replaceHash[$replaceKey]).Length / 1KB)) KiB"
                    } else {
                        Write-Verbose "    $($replaceKey): Photo not available"
                    }
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host '    Export available images'
            foreach ($VariableName in $PictureVariablesArray) {
                try { WatchCatchableExitSignal } catch { }

                Write-Verbose "    $($VariableName[0]), $([math]::ceiling(($ReplaceHash[$VariableName[0]]).Length / 1KB)) KiB @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                if ($null -ne $($ReplaceHash[$VariableName[0]])) {
                    [System.IO.File]::WriteAllBytes($(((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))), $($ReplaceHash[$VariableName[0]]))
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "  Download roaming signatures from Exchange Online @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

            if ($MirrorCloudSignatures -ne $false) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesDownload')))) {
                    Write-Host '    Roaming signatures cannot be downloaded from Exchange Online.' -ForegroundColor Green
                    Write-Host "    The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                    Write-Host '    Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesDownload()

                    if ($FeatureResult -ne 'true') {
                        Write-Host '    Error downloading roaming signatures from the cloud.' -ForegroundColor Yellow
                        Write-Host "    $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "    Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
            }

            try { WatchCatchableExitSignal } catch { }


            $CurrentTemplateIsForAliasSmtp = $null

            EvaluateAndSetSignatures


            # Delete photos from file system
            foreach ($VariableName in $PictureVariablesArray) {
                try { WatchCatchableExitSignal } catch { }

                Remove-Item -LiteralPath (((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')))) -Force -ErrorAction SilentlyContinue
                $ReplaceHash.Remove($VariableName[0])
                $ReplaceHash.Remove(($VariableName[0][-999..-2] -join '') + 'DELETEEMPTY$')
            }


            # Set OOF message and Outlook Web signature
            if (
                ((($SetCurrentUserOutlookWebSignature -eq $true)) -or ($SetCurrentUserOOFMessage -eq $true)) -and
                ($MailAddresses[$AccountNumberRunning] -ieq $PrimaryMailboxAddress)
            ) {
                Write-Host "  Set default signature(s) in Outlook Web @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                if ($SetCurrentUserOutlookWebSignature) {
                    if ($SimulateUser -and (-not $SimulateAndDeploy)) {
                        Write-Host '      Simulation mode enabled, skipping task.' -ForegroundColor Yellow
                    } else {
                        Write-Host "    Set default classic (not roaming) Outlook Web signature @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SetCurrentUserOutlookWebSignature')))) {
                            Write-Host '      Default classic Outlook Web signature cannot be set.' -ForegroundColor Green
                            Write-Host "      The 'SetCurrentUserOutlookWebSignature' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                            Write-Host '      Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                        } else {
                            try { WatchCatchableExitSignal } catch { }

                            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SetCurrentUserOutlookWebSignature()

                            if ($FeatureResult -ne 'true') {
                                Write-Host '      Error setting current user Outlook web signature.' -ForegroundColor Yellow
                                Write-Host "      $FeatureResult" -ForegroundColor Yellow
                            }
                        }

                        Write-Host "    Set default roaming Outlook Web signature(s) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                        if ($MirrorCloudSignatures -ne $false) {
                            if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesSetDefaults')))) {
                                Write-Host '      Default roaming Outlook Web signature(s) cannot be set. This also affects New Outlook on Windows.' -ForegroundColor Green
                                Write-Host "      The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                                Write-Host '      Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                            } else {
                                try { WatchCatchableExitSignal } catch { }

                                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesSetDefaults()

                                if ($FeatureResult -ne 'true') {
                                    Write-Host '      Error setting default roaming signatures in the cloud.' -ForegroundColor Yellow
                                    Write-Host "      $FeatureResult" -ForegroundColor Yellow
                                }
                            }
                        } else {
                            Write-Host "      Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
                        }
                    }
                } else {
                    Write-Host "    Parameter 'SetCurrentUserOutlookWebSignature' is not enabled, skipping task."
                }

                Write-Host "  Process out-of-office (OOF) auto replies @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

                if ($SetCurrentUserOOFMessage) {
                    if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SetCurrentUserOOFMessage')))) {
                        Write-Host '    The out-of-office replies cannot be set.' -ForegroundColor Green
                        Write-Host "    The 'SetCurrentUserOOFMessage' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                        Write-Host '    Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
                    } else {
                        try { WatchCatchableExitSignal } catch { }

                        $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SetCurrentUserOOFMessage()

                        if ($FeatureResult -ne 'true') {
                            Write-Host '    Error setting current user out-of-office message.' -ForegroundColor Yellow
                            Write-Host "    $FeatureResult" -ForegroundColor Yellow
                        }
                    }
                } else {
                    Write-Host "    Parameter 'SetCurrentUserOOFMessage' is not enabled, skipping task."
                }
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # Close Word, as it is no longer needed
    if ($script:COMWord) {
        try {
            $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
        } catch {
        }

        $script:COMWord.Quit([ref]$false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null
        Remove-Variable -Name 'COMWord' -Scope 'script'
    }


    try { WatchCatchableExitSignal } catch { }


    # Delete old signatures created by this script, which are no longer available in $SignatureTemplatePath
    # We check all local signatures for a specific marker in HTML code, so we don't touch user-created signatures
    Write-Host
    Write-Host "Remove old signatures created by this script, which are no longer centrally available @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($DeleteScriptCreatedSignaturesWithoutTemplate -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DeleteScriptCreatedSignaturesWithoutTemplate')))) {
            Write-Host '  Cannot delete old signatures created by Set-OutlookSignatures, which are no longer centrally available.' -ForegroundColor Green
            Write-Host "  The 'DeleteScriptCreatedSignaturesWithoutTemplate' feature requires the Benefactor Circle add-on." -ForegroundColor Green
            Write-Host '  Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
        } else {
            try { WatchCatchableExitSignal } catch { }
            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DeleteScriptCreatedSignaturesWithoutTemplate()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error deleting script created signature which no longer have a corresponding template.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "  Parameter 'DeleteScriptCreatedSignaturesWithoutTemplate' is not enabled, skipping task."
    }


    try { WatchCatchableExitSignal } catch { }


    # Delete user-created signatures if $DeleteUserCreatedSignatures -eq $true
    Write-Host
    Write-Host "Remove user-created signatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($DeleteUserCreatedSignatures -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DeleteUserCreatedSignatures')))) {
            Write-Host '  Cannot remove user-created signatures.' -ForegroundColor Green
            Write-Host "  The 'DeleteUserCreatedSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Green
            Write-Host '  Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
        } else {
            try { WatchCatchableExitSignal } catch { }

            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DeleteUserCreatedSignatures()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error removing user-created signatures.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "  Parameter 'DeleteUserCreatedSignatures' is not enabled, skipping task."
    }

    try { WatchCatchableExitSignal } catch { }

    # Upload local signatures to Exchange Online as roaming signatures
    Write-Host
    Write-Host "Upload local signatures to Exchange Online as roaming signatures for current user @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($MirrorCloudSignatures -ne $false) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesUpload')))) {
            Write-Host '  Signature(s) cannot be uploaded to Exchange Online. This affects Outlook Web and New Outlook on Windows.' -ForegroundColor Green
            Write-Host "  The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Green
            Write-Host '  Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
        } else {
            try { WatchCatchableExitSignal } catch { }

            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesUpload()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error uploading roaming signatures to the cloud.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "    Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
    }


    try { WatchCatchableExitSignal } catch { }


    # Prepare data for Outlook add-in
    Write-Host
    Write-Host "Prepare data for Outlook add-in @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host '  Required because Microsoft actively blocks Outlook add-ins from using roaming signatures'

    [SetOutlookSignatures.Common]::PrepareOutlookAddinDataCommon()


    try { WatchCatchableExitSignal } catch { }


    # Create/update 'My signatures, powered by Set-OutlookSignatures Benefactor Circle' email draft
    Write-Host
    Write-Host "Create 'My signatures, powered by Set-OutlookSignatures Benefactor Circle' email draft for current user @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($SignatureCollectionInDrafts -eq $true) {
        if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('SignatureCollectionInDrafts')))) {
            Write-Host '  Cannot create email draft containing all signatures.' -ForegroundColor Green
            Write-Host "  The 'SignatureCollectionInDrafts' feature requires the Benefactor Circle add-on." -ForegroundColor Green
            Write-Host '  Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
        } else {
            try { WatchCatchableExitSignal } catch { }

            $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::SignatureCollectionInDrafts()

            if ($FeatureResult -ne 'true') {
                Write-Host '  Error creating ''My signatures, powered by Set-OutlookSignatures Benefactor Circle'' email draft.' -ForegroundColor Yellow
                Write-Host "  $FeatureResult" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "    Parameter 'SignatureCollectionInDrafts' is not enabled, skipping task."
    }


    try { WatchCatchableExitSignal } catch { }


    # Copy signatures to additional path if $AdditionalSignaturePath is set
    Write-Host
    Write-Host "Copy signatures to AdditionalSignaturePath @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($AdditionalSignaturePath) {
        Write-Host "  '$AdditionalSignaturePath'"

        if ($SimulateUser) {
            Write-Host '    Simulation mode enabled, AdditionalSignaturePath already used as output directory'
        } else {
            if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('AdditionalSignaturePath')))) {
                Write-Host '    Cannot copy signatures to additional signature path.' -ForegroundColor Green
                Write-Host "    The 'AdditionalSignaturePath' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                Write-Host '    Visit https://set-outlooksignatures.com/benefactorcircle for details.' -ForegroundColor Green
            } else {
                try { WatchCatchableExitSignal } catch { }

                $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::AdditionalSignaturePath()

                if ($FeatureResult -ne 'true') {
                    Write-Host '    Error copying signatures to additional signature path.' -ForegroundColor Yellow
                    Write-Host "    $FeatureResult" -ForegroundColor Yellow
                }
            }
        }
    } else {
        Write-Host "    Parameter 'AdditionalSignaturePath' is not enabled, skipping task."
    }

    try { WatchCatchableExitSignal } catch { }

    if (
        ($script:GraphUserDummyMailbox -eq $true) -or
        ($OutlookUseNewOutlook -eq $true)
    ) {
        RemoveItemAlternativeRecurse $SignaturePaths[0] -SkipFolder
    }
}


function ResolveToSid($string) {
    try { WatchCatchableExitSignal } catch { }

    # Find the last ':', use everything right from it and remove surrounding whitespace
    $string = (($string -split ':')[-1]).trim()

    if ($string.contains('\')) {
        # is already in pre-Windows 2000 format
        $local:NTName = $string
    } elseif ($string.contains(' ')) {
        # format it in pre-Windows 2000 format
        $local:NTName = ([regex]' ').replace($string, '\', 1)
    } else {
        # Invalid
        return $null
    }

    if (($null -ne $TrustsToCheckForGroups[0]) -and ($local:NTName -inotmatch '^(EntraID\\|AzureAD\\|EntraID_\S+\\|AzureAD_\S+\\)')) {
        try {
            try { WatchCatchableExitSignal } catch { }
            $local:x = (New-Object System.Security.Principal.NTAccount($local:NTName)).Translate([System.Security.Principal.SecurityIdentifier]).value

            if ($local:x) {
                return $local:x
            }
        } catch {
            try { WatchCatchableExitSignal } catch { }

            try {
                # No group with this sAMAccountName found. Interpreting it as a display name.

                $objTrans = New-Object -ComObject 'NameTranslate'
                $objNT = $objTrans.GetType()
                $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($local:NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (4, ($local:NTName -split '\\')[1])) # 4 = ADS_NAME_TYPE_DISPLAY

                try { WatchCatchableExitSignal } catch { }
                $local:x = $(((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                try { WatchCatchableExitSignal } catch { }

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                Remove-Variable -Name 'objTrans'
                Remove-Variable -Name 'objNT'

                if ($local:x) {
                    return $local:x
                }
            } catch {
                try { WatchCatchableExitSignal } catch { }

                try {
                    # Let the API guess what it is

                    $objTrans = New-Object -ComObject 'NameTranslate'
                    $objNT = $objTrans.GetType()
                    $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (1, ($local:NTName -split '\\')[0])) # 1 = ADS_NAME_INITTYPE_DOMAIN
                    $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (8, ($local:NTName -split '\\')[1])) # 8 = ADS_NAME_TYPE_UNKNOWN

                    try { WatchCatchableExitSignal } catch { }
                    $local:x = $(((New-Object System.Security.Principal.NTAccount(($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)))).Translate([System.Security.Principal.SecurityIdentifier])).value)
                    try { WatchCatchableExitSignal } catch { }

                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objTrans) | Out-Null
                    Remove-Variable -Name 'objTrans'
                    Remove-Variable -Name 'objNT'

                    if ($local:x) {
                        return $local:x
                    }
                } catch {
                    # Nothing found
                    return $null
                }
            }
        }
    } else {
        $tempFilterOrder = @()

        # Object ID
        if ([guid]::TryParse($local:NTName.Split('\')[1], $([ref][guid]::Empty))) {
            $tempFilterOrder += "(id eq '$($local:NTName.Split('\')[1])')"
        }

        # securityIdentifier
        try {
            $null = [System.Security.Principal.SecurityIdentifier]$($local:NTName.Split('\')[1])
            $tempFilterOrder += "(securityIdentifier eq '$($local:NTName.Split('\')[1])')"
        } catch {
            # Do nothing
        }


        if ($local:NTName -inotmatch '^(EntraID\\|AzureAD\\|EntraID_\S+?\\|AzureAD_\S+?\\)') {
            if ($local:NTName.Split('\')[0] -inotlike '*.*') {
                # NetBIOS domain name pattern
                $tempFilterOrder += "((onPremisesNetBiosName eq '$($local:NTName.Split('\')[0])') and (onPremisesSamAccountName eq '$($local:NTName.Split('\')[1])'))"
                $tempFilterOrder += "((onPremisesNetBiosName eq '$($local:NTName.Split('\')[0])') and (displayName eq '$($local:NTName.Split('\')[1])'))"
            } else {
                # DNS domain name pattern
                $tempFilterOrder += "((onPremisesDomainName eq '$($local:NTName.Split('\')[0])') and (onPremisesSamAccountName eq '$($local:NTName.Split('\')[1])'))"
                $tempFilterOrder += "((onPremisesDomainName eq '$($local:NTName.Split('\')[0])') and (displayName eq '$($local:NTName.Split('\')[1])'))"
            }
        }

        # Email address pattern
        if ($local:NTName.Split('\')[1] -ilike '*@*') {
            $tempFilterOrder += "(proxyAddresses/any(x:x eq 'smtp:$($local:NTName.Split('\')[1])'))"
        }

        $tempFilterOrder += "(mailNickname eq '$($local:NTName.Split('\')[1])')"
        $tempFilterOrder += "(displayName eq '$($local:NTName.Split('\')[1])')"

        # Search Graph for groups
        foreach ($tempFilter in $tempFilterOrder) {
            try { WatchCatchableExitSignal } catch { }

            $tempResults = (GraphFilterGroups $tempFilter -GraphContext $($local:NTName.split('\')[0].split('_')[1]))

            if (($tempResults.error -eq $false) -and ($tempResults.groups.count -eq 1) -and $($tempResults.groups[0].value)) {
                if ($($tempResults.groups[0].value.securityidentifier)) {
                    return $($tempResults.groups[0].value.securityidentifier)
                }
            }
        }

        # Search Graph for users
        foreach ($tempFilter in $tempFilterOrder) {
            try { WatchCatchableExitSignal } catch { }

            $tempResults = (GraphFilterUsers $tempFilter -GraphContext $($local:NTName.split('\')[0].split('_')[1]))

            if (($tempResults.error -eq $false) -and ($tempResults.users.count -eq 1) -and $($tempResults.users[0].value)) {
                if ($($tempResults.users[0].value.securityidentifier)) {
                    return $($tempResults.users[0].value.securityidentifier)
                }
            }
        }

        # Nothing found
        return $null
    }
}


function GetBitness {
    [CmdletBinding()]

    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'files', HelpMessage = 'Comma separated list of files to process', ValueFromPipelineByPropertyName = $true)]
        [string[]]$fullname ,
        [Parameter(Mandatory = $true, ParameterSetName = 'folders', HelpMessage = 'Comma separated list of folders to process')]
        [string[]]$folders ,
        [Parameter(Mandatory = $false, ParameterSetName = 'folders')]
        [switch]$recurse ,
        [switch]$explain ,
        [switch]$quiet ,
        [switch]$dotnetOnly
    )

    begin {
        try { WatchCatchableExitSignal } catch { }

        [int]$MACHINE_OFFSET = 4
        [int]$PE_POINTER_OFFSET = 60

        [hashtable]$machineTypes = @{
            # Source: https://learn.microsoft.com/en-us/windows/win32/debug/pe-format#machine-types
            0x0    = 'UNKNOWN' # IMAGE_FILE_MACHINE_UNKNOWN; The content of this field is assumed to be applicable to any machine type
            0x14c  = 'x86' # IMAGE_FILE_MACHINE_I386; Intel 386 or later processors and compatible processors
            0x166  = 'R4000' # IMAGE_FILE_MACHINE_R4000; MIPS little endian
            0x169  = 'WCEMIPSV2' # IMAGE_FILE_MACHINE_WCEMIPSV2; MIPS little-endian WCE v2
            0x1a2  = 'SH3' # IMAGE_FILE_MACHINE_SH3; Hitachi SH3
            0x1a3  = 'SH3DSP' # IMAGE_FILE_MACHINE_SH3DSP; Hitachi SH3 DSP
            0x1a6  = 'SH4' # IMAGE_FILE_MACHINE_SH4; Hitachi SH4
            0x1a8  = 'SH5' # IMAGE_FILE_MACHINE_SH5; Hitachi SH5
            0x1c0  = 'ARM' # IMAGE_FILE_MACHINE_ARM; ARM little endian
            0x1c2  = 'THUMB' # IMAGE_FILE_MACHINE_THUMB; Thumb
            0x1c4  = 'ARMNT' # IMAGE_FILE_MACHINE_ARMNT; ARM Thumb-2 little endian
            0x1d3  = 'AM33' # IMAGE_FILE_MACHINE_AM33; Matsushita AM33
            0x1f0  = 'POWERPC' # IMAGE_FILE_MACHINE_POWERPC; Power PC little endian
            0x1f1  = 'POWERPCFP' # IMAGE_FILE_MACHINE_POWERPCFP; Power PC with floating point support
            0x200  = 'IA64' # IMAGE_FILE_MACHINE_IA64; Intel Itanium processor family
            0x266  = 'MIPS16' # IMAGE_FILE_MACHINE_MIPS16; MIPS16
            0x366  = 'MIPSFPU' # IMAGE_FILE_MACHINE_MIPSFPU; MIPS with FPU
            0x466  = 'MIPSFPU16' # IMAGE_FILE_MACHINE_MIPSFPU16; MIPS16 with FPU
            0x5032 = 'RISCV32' # IMAGE_FILE_MACHINE_RISCV32; RISC-V 32-bit address space
            0x5064 = 'RISCV64' # IMAGE_FILE_MACHINE_RISCV64; RISC-V 64-bit address space
            0x5128 = 'RISCV128' # IMAGE_FILE_MACHINE_RISCV128; RISC-V 128-bit address space
            0x6232 = 'LOONGARCH32' # IMAGE_FILE_MACHINE_LOONGARCH32; LoongArch 32-bit processor family
            0x6264 = 'LOONGARCH64' # IMAGE_FILE_MACHINE_LOONGARCH64; LoongArch 64-bit processor family
            0x8664 = 'x64' # IMAGE_FILE_MACHINE_AMD64; x64
            0x9041 = 'M32R' # IMAGE_FILE_MACHINE_M32R; Mitsubishi M32R little endian
            0xaa64 = 'ARM64' # IMAGE_FILE_MACHINE_ARM64; ARM64 little endian
            0xebc  = 'EBC' # IMAGE_FILE_MACHINE_EBC; EFI byte code
        }

        [hashtable]$processorAchitectures = @{
            'None'  = 'None'
            'MSIL'  = 'AnyCPU'
            'X86'   = 'x86'
            'I386'  = 'x86'
            'IA64'  = 'Itanium'
            'Amd64' = 'x64'
            'Arm'   = 'ARM'
        }

        [hashtable]$pekindsExplanations = @{
            'ILOnly'                      = 'MSIL processor neutral'
            'NotAPortableExecutableImage' = 'Not in portable executable (PE) file format'
            'PE32Plus'                    = 'Requires a 64-bit platform'
            'Preferred32Bit'              = 'Platform-agnostic but should be run on 32-bit platform'
            'Required32Bit'               = 'Runs on a 32-bit platform or in the 32-bit WOW environment on a 64-bit platform'
            'Unmanaged32Bit'              = 'Contains pure unmanaged code'
        }

        if ($PSBoundParameters[ 'folders' ]) {
            $fullname = @(foreach ($folder in $folders) {
                    Get-ChildItem -LiteralPath $folder -File -Recurse:$recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
                })
        }
    }

    process {
        foreach ($file in $fullname) {
            try {
                try { WatchCatchableExitSignal } catch { }
                $runtimeAssembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($file)
            } catch {
                $runtimeAssembly = $null
            }

            try {
                try { WatchCatchableExitSignal } catch { }
                $assembly = [System.Reflection.AssemblyName]::GetAssemblyName($file)
            } catch {
                $assembly = $null
            }

            if ((-not $dotnetOnly) -or ($assembly -and $runtimeAssembly)) {
                $data = New-Object System.Byte[] 4096

                try {
                    $stream = New-Object System.IO.FileStream -ArgumentList $file, Open, Read
                } catch {
                    $stream = $null

                    if (-not $quiet) {
                        Write-Verbose $_
                    }
                }

                if ($stream) {
                    try { WatchCatchableExitSignal } catch { }

                    [uint16]$machineUint = 0xffff
                    [int]$read = $stream.Read($data , 0 , $data.Count)

                    if ($read -gt $PE_POINTER_OFFSET) {
                        if (($data[0] -eq 0x4d) -and ($data[1] -eq 0x5a)) {
                            ## MZ
                            [int]$PE_HEADER_ADDR = [System.BitConverter]::ToInt32($data, $PE_POINTER_OFFSET)
                            [int]$typeOffset = $PE_HEADER_ADDR + $MACHINE_OFFSET
                            if ($data[$PE_HEADER_ADDR] -eq 0x50 -and $data[$PE_HEADER_ADDR + 1] -eq 0x45) {
                                ## PE
                                if ($read -gt $typeOffset + [System.Runtime.InteropServices.Marshal]::SizeOf($machineUint)) {
                                    [uint16]$machineUint = [System.BitConverter]::ToUInt16($data, $typeOffset)
                                    $versionInfo = Get-ItemProperty -LiteralPath $file -ErrorAction SilentlyContinue | Select-Object -ExpandProperty VersionInfo
                                    if ($runtimeAssembly -and ($module = ($runtimeAssembly.GetModules() | Select-Object -First 1))) {
                                        $pekinds = New-Object -TypeName System.Reflection.PortableExecutableKinds
                                        $imageFileMachine = New-Object -TypeName System.Reflection.ImageFileMachine
                                        $module.GetPEKind([ref]$pekinds, [ref]$imageFileMachine)
                                    } else {
                                        $pekinds = $null
                                        $imageFileMachine = $null
                                    }

                                    try { WatchCatchableExitSignal } catch { }

                                    [pscustomobject][ordered]@{
                                        'File'                = $file
                                        'Architecture'        = $machineTypes[[int]$machineUint]
                                        'NET Architecture'    = $(if ($assembly) { $processorAchitectures[$assembly.ProcessorArchitecture.ToString()] } else { 'Not .NET' })
                                        'NET PE Kind'         = $(if ($pekinds) { if ($explain) { ($pekinds.ToString() -split ',\s?' | ForEach-Object { $pekindsExplanations[$_] }) -join ',' } else { $pekinds.ToString() } }  else { 'Not .NET' })
                                        'NET Platform'        = $(if ($imageFileMachine) { $processorAchitectures[ $imageFileMachine.ToString() ] } else { 'Not .NET' })
                                        'NET Runtime Version' = $(if ($runtimeAssembly) { $runtimeAssembly.ImageRuntimeVersion } else { 'Not .NET' })
                                        'Company'             = $versionInfo | Select-Object -ExpandProperty CompanyName
                                        'File Version'        = $versionInfo | Select-Object -ExpandProperty FileVersionRaw
                                        'Product Name'        = $versionInfo | Select-Object -ExpandProperty ProductName
                                    }
                                } else {
                                    Write-Verbose "Only read $($data.Count) bytes from '$file' so can't read header at offset $typeOffset"
                                }
                            } elseif (-not $quiet) {
                                Write-Verbose "'$file' does not have a PE header signature"
                            }
                        } elseif (-not $quiet) {
                            Write-Verbose "'$file' is not an executable"
                        }
                    } elseif (-not $quiet) {
                        Write-Verbose "Only read $read bytes from '$file', not enough to get header at $PE_POINTER_OFFSET"
                    }
                    $stream.Close()
                    $stream = $null
                }
            }
        }
    }
}


function EvaluateAndSetSignatures {
    param(
        [switch]$ProcessOOF = $false
    )

    if ($ProcessOOF -eq $true) {
        $SigOrOOF = 'OOF'
        $Indent = '  '
    } else {
        $SigOrOOF = 'Signature'
        $Indent = ''
    }

    foreach ($TemplateGroup in ('common', 'group', 'mailbox', 'replacementvariable')) {
        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent  Process $TemplateGroup $(if($TemplateGroup -iin ('group', 'mailbox', 'replacementvariable')){'specific '})templates @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if (-not (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly -ErrorAction SilentlyContinue)) {
            continue
        }

        for ($TemplateFileIndex = 0; $TemplateFileIndex -lt (Get-Variable -Name "$($SigOrOOF)Files" -ValueOnly).count; $TemplateFileIndex++) {
            try { WatchCatchableExitSignal } catch { }

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

            Write-Host "$Indent    '$([System.IO.Path]::GetFileName($Template.key))' ($($SigOrOOF) INI index #$($TemplateIniSettingsIndex)) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            Write-Host "$Indent      Check permissions"

            $TemplateAllowed = $false
            $CurrentTemplateIsForAliasSmtp = $null

            try { WatchCatchableExitSignal } catch { }

            # check for allow entries
            Write-Host "$Indent        Allows"
            if ($TemplateGroup -ieq 'common') {
                $TemplateAllowed = $true
                Write-Host "$Indent          Common: Template is classified as common template valid for all mailboxes"
            } elseif ($TemplateGroup -ieq 'group') {
                try { WatchCatchableExitSignal } catch { }

                $tempAllowCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($GroupsSid)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '${1}') -join '|') = $($GroupsSid) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          No group match for current mailbox, checking current user specific allows"

                    foreach ($GroupsSid in $ADPropsCurrentUser.GroupsSIDs) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[CURRENTUSER:$($GroupsSid)``]*") {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', 'CURRENTUSER:${1}') -join '|') = $($GroupsSid) (current user)"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox and current user are not member of any allowed group"
                }
            } elseif ($TemplateGroup -ieq 'mailbox') {
                try { WatchCatchableExitSignal } catch { }

                $tempAllowCount = 0

                foreach ($CurrentMailboxSmtpaddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $true
                        $tempAllowCount++
                        $CurrentTemplateIsForAliasSmtp = $CurrentMailboxSmtpaddress
                        Write-Host "$Indent          First email address match: $($CurrentMailboxSmtpAddress) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          No email address match for current mailbox, checking current user specific allows"

                    try { WatchCatchableExitSignal } catch { }

                    foreach ($CurrentUserSmtpaddress in $ADPropsCurrentUser.proxyaddresses) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[CURRENTUSER:$($CurrentUserSmtpAddress -ireplace '^smtp:', '')``]*") {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            $CurrentTemplateIsForAliasSmtp = $CurrentUserSmtpaddress
                            Write-Host "$Indent          First email address match: $($CurrentUserSmtpAddress -ireplace '^smtp:', '') (current user)"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Email address: Mailbox and current user do not have any allowed email address"
                }
            } elseif ($TemplateGroup -ieq 'replacementvariable') {
                try { WatchCatchableExitSignal } catch { }

                $tempAllowCount = 0

                foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture 127)) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesReplacementvariableFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[$($replaceKey)``]*") {
                        if ([bool]($ReplaceHash[$replaceKey])) {
                            $TemplateAllowed = $true
                            $tempAllowCount++
                            Write-Host "$Indent          First replacement variable match: $($replaceKey) evaluates to true"
                            break
                        }
                    }
                }

                if ($tempAllowCount -eq 0) {
                    Write-Host "$Indent          Replacement variable: No allowed replacement variable evaluates to true"
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # check for deny entries
            if ($TemplateAllowed -eq $true) {
                Write-Host "$Indent        Denies"
                # check for group deny
                $tempDenyCount = 0

                foreach ($GroupsSid in $GroupsSIDs) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($GroupsSid)``]*") {
                        $TemplateAllowed = $false
                        $tempDenyCount++
                        Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '-:${1}') -join '|') = $($GroupsSid) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          No group match for current mailbox, checking current user specific denies"

                    foreach ($GroupsSid in $ADPropsCurrentUser.GroupsSIDs) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesGroupFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-CURRENTUSER:$($GroupsSid)``]*") {
                            $TemplateAllowed = $false
                            $tempDenyCount++
                            Write-Host "$Indent          First group match: $(@(@($TemplateFilesGroupSIDsOverall.getenumerator() | Where-Object { $_.value -ieq $GroupsSid }).name -ireplace '^\[(.*)\]$', '-CURRENTUSER:${1}') -join '|') = $($GroupsSid) (current user)"
                            break
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Group: Mailbox and current user are not member of any denied group"
                }

                try { WatchCatchableExitSignal } catch { }

                # check for mail address deny
                $tempDenyCount = 0

                foreach ($CurrentMailboxSmtpaddress in $CurrentMailboxSmtpAddresses) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($CurrentMailboxSmtpAddress)``]*") {
                        $TemplateAllowed = $false
                        $tempDenyCount++
                        Write-Host "$Indent          First email address match: $($CurrentMailboxSmtpAddress) (current mailbox)"
                        break
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          No email address match for current mailbox, checking current user specific denies"

                    foreach ($CurrentUserSmtpaddress in $ADPropsCurrentUser.proxyaddresses) {
                        if ((Get-Variable -Name "$($SigOrOOF)FilesMailboxFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-CURRENTUSER:$($CurrentUserSmtpAddress -ireplace '^smtp:', '')``]*") {
                            $TemplateAllowed = $false
                            $tempDenyCount++
                            Write-Host "$Indent          First email address match: $($CurrentUserSmtpAddress -ireplace '^smtp:', '') (current user)"
                            break
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Email address: Mailbox and current user do not have any denied email address"
                }

                try { WatchCatchableExitSignal } catch { }

                # check for replacement variable deny
                $tempDenyCount = 0

                foreach ($replaceKey in @($replaceHash.Keys | Sort-Object -Culture 127)) {
                    if ((Get-Variable -Name "$($SigOrOOF)FilesReplacementvariableFilePart" -ValueOnly)[$TemplateIniSettingsIndex] -ilike "*``[-:$($replaceKey)``]*") {
                        if ([bool]($ReplaceHash[$replaceKey])) {
                            $TemplateAllowed = $false
                            $tempDenyCount++
                            Write-Host "$Indent          First replacement variable match: $($replaceKey) evaluates to true"
                            break
                        }
                    }
                }

                if ($tempDenyCount -eq 0) {
                    Write-Host "$Indent          Replacement variable: No deny replacement variable evaluates to true"
                }
            }

            try { WatchCatchableExitSignal } catch { }

            # result
            if ($Template -and ($TemplateAllowed -eq $true)) {
                Write-Host "$Indent        Use template as there is at least one allow and no deny"
                if ($ProcessOOF) {
                    if ($OOFFilesInternal.contains($TemplateIniSettingsIndex)) {
                        $OOFInternal = $Template
                        $script:OOFInternalValueBasename = $(($OOFInternal.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }

                    if ($OOFFilesExternal.contains($TemplateIniSettingsIndex)) {
                        $OOFExternal = $Template
                        $script:OOFExternalValueBasename = $(($OOFExternal.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }
                } else {
                    $Signature = $Template

                    try { WatchCatchableExitSignal } catch { }

                    SetSignatures -ProcessOOF:$ProcessOOF
                }
            } else {
                Write-Host "$Indent        Do not use template as there is no allow or at least one deny"
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ($ProcessOOF) {
        # Internal OOF message
        if ($OOFInternal -or $OOFExternal) {
            Write-Host "$Indent  Convert final OOF templates to HTM format @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        }

        if ($OOFInternal) {
            $Signature = $OOFInternal

            if ($OOFExternal -eq $OOFInternal) {
                Write-Host "$Indent    Common OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            } else {
                Write-Host "$Indent    Internal OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
            }

            if ($UseHtmTemplates) {
                $Signature.value = "$OOFInternalGUID OOFInternal.htm"
            } else {
                $Signature.value = "$OOFInternalGUID OOFInternal.docx"
            }

            try { WatchCatchableExitSignal } catch { }

            SetSignatures -ProcessOOF:$ProcessOOF

            try { WatchCatchableExitSignal } catch { }

            if ($OOFExternal -eq $OOFInternal) {
                Copy-Item -LiteralPath (Join-Path -Path $script:tempDir -ChildPath "$OOFInternalGUID OOFInternal.htm") -Destination (Join-Path -Path $script:tempDir -ChildPath "$OOFExternalGUID OOFExternal.htm")
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # External OOF message
    if ($OOFExternal -and ($OOFExternal -ne $OOFInternal)) {
        $Signature = $OOFExternal

        Write-Host "$Indent    External OOF message: '$($Signature.value)' @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        if ($UseHtmTemplates) {
            $Signature.value = "$OOFExternalGUID OOFExternal.htm"
        } else {
            $Signature.value = "$OOFExternalGUID OOFExternal.docx"
        }

        try { WatchCatchableExitSignal } catch { }

        SetSignatures -ProcessOOF:$ProcessOOF
    }

    try { WatchCatchableExitSignal } catch { }
}


function SetSignatures {
    param(
        [switch]$ProcessOOF = $false
    )

    try { WatchCatchableExitSignal } catch { }

    if ($ProcessOOF) {
        $Indent = '  '
    }

    if (-not $ProcessOOF) {
        Write-Host "      Outlook signature name: '$([System.IO.Path]::ChangeExtension($($Signature.value), $null) -ireplace '\.$')'"

        if ($MailboxSpecificSignatureNames) {
            Write-Host "        Mailbox specific signature name: '$([System.IO.Path]::GetFileNameWithoutExtension($Signature.Value)) ($($MailAddresses[$AccountNumberRunning]))'"

            $SignatureFileAlreadyDone = $false
        } else {
            $SignatureFileAlreadyDone = ($script:SignatureFilesDone -contains $TemplateIniSettingsIndex)

            if ($SignatureFileAlreadyDone) {
                Write-Host "$Indent      $($SigOrOOF) INI index #$($TemplateIniSettingsIndex) already processed before with higher priority mailbox"
                Write-Host "$Indent        Not overwriting signature. Consider using parameter MailboxSpecificSignatureNames."
            } else {
                $script:SignatureFilesDone += $TemplateIniSettingsIndex
            }
        }
    }

    if (($SignatureFileAlreadyDone -eq $false) -or $ProcessOOF) {
        Write-Host "$Indent      Create temporary file copy"

        $pathGUID = (New-Guid).guid
        $path = Join-Path -Path $script:tempDir -ChildPath "$($pathGUID).htm"

        $pathConnectedFolderNames = @()

        foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
            $pathConnectedFolderNames += "$($pathGUID)$($ConnectedFilesFolderName)"
            $pathConnectedFolderNames += [uri]::EscapeDataString($pathConnectedFolderNames[-1])
            $pathConnectedFolderNames += [uri]::EscapeUriString($pathConnectedFolderNames[-2])

            $pathConnectedFolderNames += "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$($ConnectedFilesFolderName)"
            $pathConnectedFolderNames += [uri]::EscapeDataString($pathConnectedFolderNames[-1])
            $pathConnectedFolderNames += [uri]::EscapeUriString($pathConnectedFolderNames[-2])
        }

        $pathConnectedFolderNames = $pathConnectedFolderNames | Select-Object -Unique

        try { WatchCatchableExitSignal } catch { }

        if ($UseHtmTemplates) {
            try {
                if ($script:SpoDownloadUrls -and $script:SpoDownloadUrls["$($Signature.name)"]) {
                    $(New-Object Net.WebClient).DownloadFile(
                        $script:SpoDownloadUrls["$($Signature.name)"],
                        $path
                    )
                } else {
                    Copy-Item -LiteralPath $Signature.name -Destination $path -Force
                }

                try { WatchCatchableExitSignal } catch { }

                foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                    try { WatchCatchableExitSignal } catch { }

                    $pathTemp = (Join-Path -Path (Split-Path -LiteralPath $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName")

                    if (Test-Path -LiteralPath $pathTemp) {
                        if ($script:SpoDownloadUrls) {
                            # Work around a bug in WebDAV or .Net (https://github.com/dotnet/runtime/issues/49803)
                            #   Do not use 'Get-ChildItem'
                            $tempFiles = @()

                            [System.IO.Directory]::EnumerateFiles((Join-Path -Path (Split-Path -LiteralPath $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName"), '*', [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                                $tempX = $_ -replace $([char]0)

                                if (
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)."))$") -or
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).$([IO.Path]::DirectorySeparatorChar)"))") -or
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).."))$") -or
                                    $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)..$([IO.Path]::DirectorySeparatorChar)"))")
                                ) {
                                    # do nothing
                                } else {
                                    $tempFiles += $tempX
                                }
                            }

                            $tempFiles = $tempFiles | Select-Object -Unique

                            foreach ($tempX in $tempFiles) {
                                if ($script:SpoDownloadUrls -and $script:SpoDownloadUrls["$($tempX)"]) {
                                    try { WatchCatchableExitSignal } catch { }

                                    $(New-Object Net.WebClient).DownloadFile(
                                        $script:SpoDownloadUrls["$($tempX)"],
                                        $tempX
                                    )
                                }
                            }
                        }

                        try { WatchCatchableExitSignal } catch { }


                        # Work around a bug in WebDAV or .Net (https://github.com/dotnet/runtime/issues/49803)
                        #   Do not use 'Get-ChildItem'
                        $tempFiles = @()

                        [System.IO.Directory]::EnumerateFiles((Join-Path -Path (Split-Path -LiteralPath $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName"), '*', [System.IO.SearchOption]::AllDirectories) | ForEach-Object {
                            $tempX = $_ -replace $([char]0)

                            if (
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)."))$") -or
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).$([IO.Path]::DirectorySeparatorChar)"))") -or
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar).."))$") -or
                                $($tempX -imatch "$([regex]::escape("$([IO.Path]::DirectorySeparatorChar)..$([IO.Path]::DirectorySeparatorChar)"))")
                            ) {
                                # do nothing
                            } else {
                                $tempFiles += $tempX
                            }
                        }

                        $tempFiles = $tempFiles | Select-Object -Unique

                        foreach ($tempX in $tempFiles) {
                            $tempY = (Join-Path -Path (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files") -ChildPath ($tempX -ireplace "^$([regex]::escape("$(Join-Path -Path (Split-Path -LiteralPath $signature.name) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.name))$ConnectedFilesFolderName")$([IO.Path]::DirectorySeparatorChar)"))", ''))

                            $(Split-Path -LiteralPath $tempY) | ForEach-Object {
                                if (-not (Test-Path -LiteralPath $_ -PathType Container)) {
                                    $null = New-Item -ItemType Directory -Path $_
                                }
                            }

                            Copy-Item -LiteralPath $tempX -Destination $tempY -Force
                        }

                        break
                    }
                }
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                Write-Host $error[0]
                continue
            }
        } else {
            $path = $([System.IO.Path]::ChangeExtension($($path), '.docx'))

            try {
                try { WatchCatchableExitSignal } catch { }

                if ($script:SpoDownloadUrls -and $script:SpoDownloadUrls["$($Signature.name)"]) {
                    $(New-Object Net.WebClient).DownloadFile(
                        $script:SpoDownloadUrls["$($Signature.name)"],
                        $path
                    )
                } else {
                    Copy-Item -LiteralPath $Signature.name -Destination $path -Force
                }
            } catch {
                Write-Host "$Indent        Error copying file. Skip template." -ForegroundColor Red
                continue
            }
        }

        try { WatchCatchableExitSignal } catch { }


        @(
            @(Get-ChildItem -LiteralPath $path -Force) +
            @(Get-ChildItem -LiteralPath (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files") -Recurse -Force -ErrorAction SilentlyContinue)
        ) | ForEach-Object {
            if (-not $_.PSIsContainer) {
                Set-ItemProperty -LiteralPath $_.FullName -Name IsReadOnly -Value $false

                if (-not $IsLinux) {
                    Unblock-File -LiteralPath $_.FullName
                }
            }
        }


        try { WatchCatchableExitSignal } catch { }


        $Signature.value = $([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))

        if ($MailboxSpecificSignatureNames -and ($ProcessOOF -eq $false)) {
            if ($OutlookDisableRoamingSignatures -eq 0) {
                $Signature.value = ($Signature.Value -ireplace '\.htm$', " ($($MailAddresses[$AccountNumberRunning])).htm")
            } else {
                $Signature.value = ($Signature.Value -ireplace '\.htm$', " ($($MailAddresses[$AccountNumberRunning])).htm")
            }
        }

        if (-not $ProcessOOF) {
            $script:SignatureFilesDone += $Signature.Value
        }

        try { WatchCatchableExitSignal } catch { }

        if ($UseHtmTemplates) {
            # Non-picture variables first, to allow for dynamic code creation
            # Picture variables are replaced using the DOM later

            Write-Host "$Indent      Replace non-picture variables"
            $tempFileContent = (ConvertEncoding -InFile $path)

            foreach ($replaceKey in @($replaceHash.Keys | Where-Object { $_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' }) } | Sort-Object -Culture 127)) {
                $tempFileContent = $tempFileContent -ireplace [Regex]::Escape($replacekey), $replaceHash.$replaceKey
            }

            try { WatchCatchableExitSignal } catch { }

            [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $tempFileContent)

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Replace picture variables"

            $htmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
            $htmlDoc.DisableImplicitEnd = $true
            $htmlDoc.OptionAutoCloseOnEnd = $true
            $htmlDoc.OptionCheckSyntax = $true
            $htmlDoc.OptionEmptyCollection = $true
            $htmlDoc.OptionFixNestedTags = $true

            $htmlDoc.LoadHtml((ConvertEncoding -InFile $path))

            $htmlDocSelectNodeResult = $htmlDoc.DocumentNode.SelectNodes('//img')

            if ($htmlDocSelectNodeResult) {
                foreach ($image in $htmlDocSelectNodeResult) {
                    try { WatchCatchableExitSignal } catch { }

                    $DuplicateLocalSrcPaths = @()

                    $tempHtmlDocSelectNodeResult = $htmlDoc.DocumentNode.SelectNodes('//img[@src and normalize-space(@src) != '''']')

                    if ($tempHtmlDocSelectNodeResult) {
                        foreach ($tempImgNode in $tempHtmlDocSelectNodeResult) {
                            $tempSrc = $tempImgNode.GetAttributeValue('src', '')

                            if (
                                $(-not $tempSrc) -or
                                $($tempSrc.StartsWith('data:', [System.StringComparison]::OrdinalIgnoreCase)) -or
                                $($tempSrc.StartsWith('http://', [System.StringComparison]::OrdinalIgnoreCase)) -or
                                $($tempSrc.StartsWith('https://', [System.StringComparison]::OrdinalIgnoreCase)) -or
                                $($tempSrc.StartsWith('about:', [System.StringComparison]::OrdinalIgnoreCase))
                            ) {
                                continue
                            }

                            $DuplicateLocalSrcPaths += $(Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode($tempSrc.Trim()))))")
                        }
                    }

                    $DuplicateLocalSrcPaths = @($DuplicateLocalSrcPaths | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name)

                    $tempImageIsDeleted = $false
                    $srcValue = $image.GetAttributeValue('src', '')
                    $altValue = $image.GetAttributeValue('alt', '')

                    if (
                        $($srcValue -ilike '*$*$*') -or
                        $($altValue -ilike '*$*$*')
                    ) {
                        foreach ($VariableName in $PictureVariablesArray) {
                            try { WatchCatchableExitSignal } catch { }

                            $tempImageVariableString = $VariableName[0] -ireplace '\$$', 'DELETEEMPTY$'
                            $NewImageFilenameGuid = (New-Guid).Guid

                            if (
                                $($srcValue -ilike "*$($VariableName[0])*") -or
                                $($altValue -ilike "*$($VariableName[0])*")
                            ) {
                                if ($ReplaceHash[$VariableName[0]]) {
                                    if (-not $EmbedImagesInHtml) {
                                        @(
                                            $(Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode($srcValue))))")
                                        ) | ForEach-Object {
                                            if ($DuplicateLocalSrcPaths -notcontains $_) {
                                                Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue
                                            }
                                        }

                                        Copy-Item -LiteralPath (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$($NewImageFilenameGuid).jpeg") -Force

                                        $null = $image.SetAttributeValue('src', $([System.Net.WebUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($NewImageFilenameGuid).jpeg")))

                                        #if ($altValue) {
                                        #    $null = $image.SetAttributeValue('alt', $($altValue -ireplace [Regex]::Escape($VariableName[0]), ''))
                                        #}
                                    } else {
                                        $null = $image.SetAttributeValue('src', $('data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                    }
                                } else {
                                    $null = $image.SetAttributeValue('src', "$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode($srcValue))))")
                                }
                            } elseif (
                                $($srcValue -ilike "*$($tempImageVariableString)*") -or
                                $($altValue -ilike "*$($tempImageVariableString)*")
                            ) {
                                if ($ReplaceHash[$VariableName[0]]) {
                                    if (-not $EmbedImagesInHtml) {
                                        @(
                                            $(Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode($srcValue))))")
                                        ) | ForEach-Object {
                                            if ($DuplicateLocalSrcPaths -notcontains $_) {
                                                Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue
                                            }
                                        }

                                        Copy-Item -LiteralPath (Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg')) (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$($NewImageFilenameGuid).jpeg") -Force

                                        $null = $image.SetAttributeValue('src', $([System.Net.WebUtility]::UrlDecode("$([System.IO.Path]::ChangeExtension($Signature.Value, '.files'))/$($NewImageFilenameGuid).jpeg")))

                                        #if ($altValue) {
                                        #    $null = $image.SetAttributeValue('alt', $($altValue -ireplace [Regex]::Escape($tempImageVariableString), ''))
                                        #}
                                    } else {
                                        $null = $image.SetAttributeValue('src', $('data:image/jpeg;base64,' + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes((Join-Path -Path $script:tempDir -ChildPath ($VariableName[0] + $VariableName[1] + '.jpeg'))))))
                                    }
                                } else {
                                    @(
                                        $(Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode($srcValue))))")
                                    ) | ForEach-Object {
                                        if ($DuplicateLocalSrcPaths -notcontains $_) {
                                            Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue
                                        }
                                    }

                                    $null = $image.Remove()
                                    $tempImageIsDeleted = $true
                                    break
                                }
                            }

                            if (
                                $(-not $tempImageIsDeleted) -and
                                $altValue
                            ) {
                                $altValue = $($altValue -ireplace [Regex]::Escape($VariableName[0]), '' -ireplace [Regex]::Escape($tempImageVariableString), '')
                                $null = $image.SetAttributeValue('alt', $altValue)
                            }
                        }

                        if ($tempImageIsDeleted) {
                            continue
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Other images
                    if (
                        $($srcValue -ilike '*$*DELETEEMPTY$*') -or
                        $($altValue -ilike '*$*DELETEEMPTY$*')
                    ) {
                        foreach ($VariableName in @(@($ReplaceHash.Keys) | Where-Object { $_ -inotin @('$CurrentMailboxPhoto$', '$CurrentMailboxManagerPhoto$', '$CurrentUserPhoto$', '$CurrentUserManagerPhoto$') } | Sort-Object -Culture 127)) {
                            try { WatchCatchableExitSignal } catch { }

                            $tempImageVariableString = $VariableName -ireplace '\$$', 'DELETEEMPTY$'

                            if (
                                $($srcValue -ilike "*$($tempImageVariableString)*") -or
                                $($altValue -ilike "*$($tempImageVariableString)*")
                            ) {
                                if ($ReplaceHash[$VariableName]) {
                                    if ($altValue) {
                                        $altValue = $($altValue -ireplace [Regex]::Escape($tempImageVariableString), '')
                                        $null = $image.SetAttributeValue('alt', $altValue)
                                    }
                                } else {
                                    @(
                                        $(Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$($pathGUID).files/$([System.IO.Path]::GetFileName(([System.Net.WebUtility]::UrlDecode($srcValue))))")
                                    ) | ForEach-Object {
                                        if ($DuplicateLocalSrcPaths -notcontains $_) {
                                            Remove-Item -LiteralPath $_ -Force -ErrorAction SilentlyContinue
                                        }
                                    }

                                    $null = $image.Remove()
                                    $tempImageIsDeleted = $true
                                    break
                                }
                            }
                        }

                        if ($tempImageIsDeleted) {
                            continue
                        }
                    }
                }
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Export to HTM format"
            [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $htmlDoc.DocumentNode.OuterHtml)
        } else {
            $script:COMWord.Documents.Open($path, $false, $false, $false) | Out-Null
            $script:COMWord.ActiveDocument.ActiveWindow.View.Type = [Microsoft.Office.Interop.Word.WdViewType]::wdWebView

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Replace non-picture variables"
            $script:COMWordShowFieldCodesOriginal = $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes

            try {
                # Replace in view without field codes
                if ($script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes -ne $false) {
                    $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $false
                }

                $script:COMWord.ActiveDocument.Select()
                $script:COMWord.Selection.Collapse()

                foreach ($replaceKey in @($replaceHash.Keys | Where-Object { ($_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' })) } | Sort-Object -Culture 127 )) {
                    try { WatchCatchableExitSignal } catch { }

                    $null = $script:COMWord.Selection.Find.Execute($replaceKey, $false, $true, $false, $false, $false, $true, 1, $false, $(($replaceHash.$replaceKey -ireplace "`r`n", '^p') -ireplace "`n", '^l'), 2)
                }

                # Restore original view
                if ($script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes -ne $script:COMWordShowFieldCodesOriginal) {
                    $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
                }

                try { WatchCatchableExitSignal } catch { }

                # Replace in field codes
                foreach ($field in $script:COMWord.ActiveDocument.Fields) {
                    try { WatchCatchableExitSignal } catch { }

                    $tempWordFieldCodeOriginal = $field.Code.Text
                    $tempWordFieldCodeNew = $tempWordFieldCodeOriginal

                    foreach ($replaceKey in @($replaceHash.Keys | Where-Object { ($_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' })) } | Sort-Object -Culture 127 )) {
                        $tempWordFieldCodeNew = $tempWordFieldCodeNew -ireplace [regex]::escape($replaceKey), $($replaceHash.$replaceKey)
                    }

                    if ($tempWordFieldCodeOriginal -ne $tempWordFieldCodeNew) {
                        $field.Code.Text = $tempWordFieldCodeNew
                    }
                }
            } catch {
                Write-Host $error[0]
                Write-Host "$Indent        Error replacing non-picture variables in Word. Exit." -ForegroundColor Red
                Write-Host "$Indent        If the error says 'Access denied', your environment may require to assign a Microsoft Purview Information Protection sensitivity label to your DOCX templates." -ForegroundColor Red
                $script:ExitCode = 21
                $script:ExitCodeDescription = 'Error replacing non-picture variables in Word.'
                exit
            }

            try { WatchCatchableExitSignal } catch { }

            Write-Host "$Indent      Replace picture variables"
            if ($script:COMWord.ActiveDocument.Shapes.Count -gt 0) {
                Write-Host "$Indent        Warning: Template contains $($script:COMWord.ActiveDocument.Shapes.Count) image(s) configured as non-inline shapes." -ForegroundColor Yellow
                Write-Host "$Indent        Set the text wrapping to 'inline with text' to avoid incorrect positioning and other problems." -ForegroundColor Yellow
            }

            try {
                foreach ($image in @(@($script:COMWord.ActiveDocument.Shapes) + @($script:COMWord.ActiveDocument.InlineShapes))) {
                    try { WatchCatchableExitSignal } catch { }

                    # Setting the values in word is very slow, so we use temporay variables
                    $tempImageIsDeleted = $false

                    $tempImageSourceFullName = $image.LinkFormat.SourceFullName
                    $tempImageAlternativeText = $image.AlternativeText
                    $tempImageHyperlinkAddress = $image.Hyperlink.Address
                    $tempImageHyperlinkSubAddress = $image.Hyperlink.SubAddress
                    $tempImageHyperlinkEmailSubject = $image.Hyperlink.EmailSubject
                    $tempImageHyperlinkScreenTip = $image.Hyperlink.ScreenTip


                    # Mailbox photos
                    if ($tempImageSourceFullName -or $tempImageAlternativeText) {
                        foreach ($VariableName in @($PictureVariablesArray)) {
                            try { WatchCatchableExitSignal } catch { }

                            if (
                                $(if ($tempImageSourceFullName) { ((Split-Path -Path $tempImageSourceFullName -Leaf) -ilike "*$($Variablename[0])*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($Variablename[0])*") })
                            ) {
                                if ($null -ne $($ReplaceHash[$Variablename[0]])) {
                                    $tempImageSourceFullName = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                }
                            } elseif (
                                $(if ($tempImageSourceFullName) { ((Split-Path -Path $tempImageSourceFullName -Leaf) -ilike "*$($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')*") })
                            ) {
                                if ($null -ne $($ReplaceHash[$Variablename[0]])) {
                                    $tempImageSourceFullName = (Join-Path -Path $script:tempDir -ChildPath ($Variablename[0] + $Variablename[1] + '.jpeg'))
                                } else {
                                    $image.delete()
                                    $tempImageIsDeleted = $true
                                    break
                                }
                            }

                            if ((-not $tempImageIsDeleted) -and ($tempImageAlternativeText)) {
                                $tempImageAlternativeText = $($tempImageAlternativeText) -ireplace [Regex]::Escape($Variablename[0]), ''
                                $tempImageAlternativeText = $($tempImageAlternativeText) -ireplace [Regex]::Escape($($Variablename[0] -ireplace '\$$', 'DELETEEMPTY$')), ''
                            }
                        }

                        if ($tempImageIsDeleted) {
                            continue
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Other images
                    if (
                        $(if ($tempImageSourceFullName) { ((Split-Path -Path $tempImageSourceFullName -Leaf) -ilike '*$*DELETEEMPTY$*') }) -or
                        $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike '*$*DELETEEMPTY$*') })
                    ) {
                        foreach ($Variablename in @(@($ReplaceHash.Keys) | Where-Object { $_ -inotin @('$CurrentMailboxPhoto$', '$CurrentMailboxManagerPhoto$', '$CurrentUserPhoto$', '$CurrentUserManagerPhoto$') } | Sort-Object -Culture 127)) {
                            $tempImageVariableString = $Variablename -ireplace '\$$', 'DELETEEMPTY$'

                            if (
                                $(if ($tempImageSourceFullName) { ((Split-Path -Path $tempImageSourceFullName -Leaf) -ilike "*$($tempImageVariableString)*") }) -or
                                $(if ($tempImageAlternativeText) { ($tempImageAlternativeText -ilike "*$($tempImageVariableString)*") })
                            ) {
                                if ($($ReplaceHash[$Variablename])) {
                                    if ($tempImageAlternativeText) {
                                        $tempImageAlternativeText = $tempImageAlternativeText -ireplace [Regex]::Escape($tempImageVariableString), ''
                                    }
                                } else {
                                    $image.delete()
                                    $tempImageIsDeleted = $true
                                    break
                                }
                            }
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }

                    if ($tempImageIsDeleted) {
                        continue
                    }

                    try { WatchCatchableExitSignal } catch { }

                    foreach ($replaceKey in @($replaceHash.Keys | Where-Object { $_ -inotin @($PictureVariablesArray | ForEach-Object { $_[0]; $_[0] -replace '\$$', 'DeleteEmpty$' }) } | Sort-Object -Culture 127)) {
                        if ($replaceKey) {
                            if ($tempImageAlternativeText) {
                                $tempImageAlternativeText = $tempImageAlternativeText -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($tempImageHyperlinkAddress) {
                                $tempImageHyperlinkAddress = $tempImageHyperlinkAddress -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($tempImageHyperlinkSubAddress) {
                                $tempImageHyperlinkSubAddress = $tempImageHyperlinkSubAddress -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($tempImageHyperlinkEmailSubject) {
                                $tempImageHyperlinkEmailSubject = $tempImageHyperlinkEmailSubject -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }

                            if ($tempImageHyperlinkScreenTip) {
                                $tempImageHyperlinkScreenTip = $tempImageHyperlinkScreenTip -ireplace [Regex]::Escape($replaceKey), $replaceHash.$replaceKey
                            }
                        }
                    }

                    try { WatchCatchableExitSignal } catch { }

                    if (
                        $($null -ne $tempImageSourceFullname) -and
                        $($null -ne $image.linkformat.sourcefullname) -and
                        ($tempImageSourceFullName -ne $image.LinkFormat.SourceFullName)
                    ) {
                        $image.LinkFormat.SourceFullName = $tempImageSourceFullName
                    }

                    if (
                        $($null -ne $tempImageAlternativeText) -and
                        $($null -ne $image.AlternativeText) -and
                        ($tempImageAlternativeText -ne $image.AlternativeText)
                    ) {
                        $image.AlternativeText = $tempImageAlternativeText
                    }

                    if (
                        $($null -ne $tempImageHyperlinkAddress) -and
                        $($null -ne $image.Hyperlink.Address) -and
                        ($tempImageHyperlinkAddress -ne $image.Hyperlink.Address)
                    ) {
                        $image.Hyperlink.Address = $tempImageHyperlinkAddress
                    }

                    if (
                        $($null -ne $tempImageHyperlinkSubAddress) -and
                        $($null -ne $image.Hyperlink.SubAddress) -and
                        ($tempImageHyperlinkSubAddress -ne $image.Hyperlink.SubAddress)
                    ) {
                        $image.Hyperlink.SubAddress = $tempImageHyperlinkSubAddress
                    }

                    if (
                        $($null -ne $tempImageHyperlinkEmailSubject) -and
                        $($null -ne $image.Hyperlink.EmailSubject) -and
                        ($tempImageHyperlinkEmailSubject -ne $image.Hyperlink.EmailSubject)
                    ) {
                        $image.Hyperlink.EmailSubject = $tempImageHyperlinkEmailSubject
                    }

                    if (
                        $($null -ne $tempImageHyperlinkScreenTip) -and
                        $($null -ne $image.Hyperlink.ScreenTip) -and
                        ($tempImageHyperlinkScreenTip -ne $image.Hyperlink.ScreenTip)
                    ) {
                        $image.Hyperlink.ScreenTip = $tempImageHyperlinkScreenTip
                    }
                }
            } catch {
                Write-Host $error[0]
                Write-Host "$Indent        Error replacing picture variables in Word. Exit." -ForegroundColor Red
                Write-Host "$Indent        If the error says 'Access denied', your environment may require to assign a Microsoft Purview Information Protection sensitivity label to your DOCX templates." -ForegroundColor Red
                $script:ExitCode = 20
                $script:ExitCodeDescription = 'Error replacing picture variables in Word.'
                exit
            }

            try { WatchCatchableExitSignal } catch { }

            # Save changed document, it's later used for export to .htm, .rtf and .txt
            $saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault

            try { WatchCatchableExitSignal } catch { }

            try {
                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            } catch {
                # Restore original security setting after error
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                Start-Sleep -Seconds 2

                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            }

            try { WatchCatchableExitSignal } catch { }

            # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
            # Close the document to remove in-memory references to already deleted images
            if ($script:COMWord.ActiveDocument.Saved -ne $true) {
                $script:COMWord.ActiveDocument.Saved = $true
            }

            $script:COMWord.ActiveDocument.ActiveWindow.View.Type = $script:COMWordViewTypeOriginal

            $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)

            try { WatchCatchableExitSignal } catch { }

            # Export to .htm
            Write-Host "$Indent      Export to HTM format"
            $path = $([System.IO.Path]::ChangeExtension($path, '.docx'))

            try { WatchCatchableExitSignal } catch { }

            $script:COMWord.Documents.Open($path, $false, $false, $false) | Out-Null
            $script:COMWord.ActiveDocument.ActiveWindow.View.Type = [Microsoft.Office.Interop.Word.WdViewType]::wdWebView

            try { WatchCatchableExitSignal } catch { }

            $saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatFilteredHTML
            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

            $script:WordWebOptions = $script:COMWord.ActiveDocument.WebOptions

            if ($script:COMWord.ActiveDocument.WebOptions.TargetBrowser -ne 4) {
                $script:COMWord.ActiveDocument.WebOptions.TargetBrowser = 4 # IE6, which is the maximum
            }
            if ($script:COMWord.ActiveDocument.WebOptions.BrowserLevel -ne 2) {
                $script:COMWord.ActiveDocument.WebOptions.BrowserLevel = 2 # IE6, which is the maximum
            }
            if ($script:COMWord.ActiveDocument.WebOptions.AllowPNG -ne $true) {
                $script:COMWord.ActiveDocument.WebOptions.AllowPNG = $true
            }
            if ($script:COMWord.ActiveDocument.WebOptions.OptimizeForBrowser -ne $false) {
                $script:COMWord.ActiveDocument.WebOptions.OptimizeForBrowser = $false
            }
            if ($script:COMWord.ActiveDocument.WebOptions.RelyOnCSS -ne $true) {
                $script:COMWord.ActiveDocument.WebOptions.RelyOnCSS = $true
            }
            if ($script:COMWord.ActiveDocument.WebOptions.RelyOnVML -ne $false) {
                $script:COMWord.ActiveDocument.WebOptions.RelyOnVML = $false
            }
            if ($script:COMWord.ActiveDocument.WebOptions.Encoding -ne 65001) {
                $script:COMWord.ActiveDocument.WebOptions.Encoding = 65001 # Outlook uses 65001 (UTF8) for .htm, but 1200 (UTF16LE a.k.a Unicode) for .txt
            }
            if ($script:COMWord.ActiveDocument.WebOptions.OrganizeInFolder -ne $true) {
                $script:COMWord.ActiveDocument.WebOptions.OrganizeInFolder = $true
            }
            if ($script:COMWord.ActiveDocument.WebOptions.PixelsPerInch -ne 96) {
                $script:COMWord.ActiveDocument.WebOptions.PixelsPerInch = 96
            }
            if ($script:COMWord.ActiveDocument.WebOptions.ScreenSize -ne 10) {
                $script:COMWord.ActiveDocument.WebOptions.ScreenSize = 10 # 1920x1200
            }
            if ($script:COMWord.ActiveDocument.WebOptions.UseLongFileNames -ne $true) {
                $script:COMWord.ActiveDocument.WebOptions.UseLongFileNames = $true
            }

            $script:COMWord.ActiveDocument.WebOptions.UseDefaultFolderSuffix()
            $pathHtmlFolderSuffix = $script:COMWord.ActiveDocument.WebOptions.FolderSuffix

            try {
                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            } catch {
                # Restore original security setting after error
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                Start-Sleep -Seconds 2

                # Overcome Word security warning when export contains embedded pictures
                if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                }

                if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                    $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                }

                if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                    $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                }

                try { WatchCatchableExitSignal } catch { }

                # Save
                $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
            }

            try { WatchCatchableExitSignal } catch { }

            # Restore original WebOptions
            try {
                if ($script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        if ($script:COMWord.ActiveDocument.WebOptions.$property -ne $script:WordWebOptions.$property) {
                            $script:COMWord.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                        }
                    }
                }
            } catch {}

            try { WatchCatchableExitSignal } catch { }

            # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
            if ($script:COMWord.ActiveDocument.Saved -ne $true) {
                $script:COMWord.ActiveDocument.Saved = $true
            }

            Write-Host "$Indent        Export high-res images"

            if ($DocxHighResImageConversion) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('DocxHighResImageConversion')))) {
                    $script:COMWord.ActiveDocument.ActiveWindow.View.Type = $script:COMWordViewTypeOriginal

                    $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)

                    Write-Host "$Indent          Cannot export high-res images." -ForegroundColor Green
                    Write-Host "$Indent          The 'DocxHighResImageConversion' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                    Write-Host "$Indent          Visit https://set-outlooksignatures.com/benefactorcircle for details." -ForegroundColor Green
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::DocxHighResImageConversion()

                    if ($FeatureResult -ne 'true') {
                        try {
                            $script:COMWord.ActiveDocument.ActiveWindow.View.Type = $script:COMWordViewTypeOriginal

                            $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)
                        } catch {
                        }
                        Write-Host "$Indent          Error converting high resolution images from DOCX template." -ForegroundColor Yellow
                        Write-Host "$Indent          $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "$Indent          Parameter 'DocxHighResImageConversion' is not enabled, skipping task."

                $script:COMWord.ActiveDocument.ActiveWindow.View.Type = $script:COMWordViewTypeOriginal

                $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Copy HTM image width and height attributes to style attribute"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

        $htmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
        $htmlDoc.DisableImplicitEnd = $true
        $htmlDoc.OptionAutoCloseOnEnd = $true
        $htmlDoc.OptionCheckSyntax = $true
        $htmlDoc.OptionEmptyCollection = $true
        $htmlDoc.OptionFixNestedTags = $true

        $htmlDoc.LoadHtml((ConvertEncoding -InFile $path))

        $htmlDocSelectNodeResult = $htmlDoc.DocumentNode.SelectNodes('//img')

        if ($htmlDocSelectNodeResult) {
            foreach ($image in $htmlDocSelectNodeResult) {
                $currentStyle = $image.GetAttributeValue('style', '')

                if ($null -ne $currentStyle) {
                    $currentStyle = $currentStyle.Trim()
                } else {
                    continue
                }

                $newStyleParts = @()

                $widthAttribute = $image.Attributes['width']

                if ($null -ne $widthAttribute) {
                    $width = $widthAttribute.Value

                    if (-not [string]::IsNullOrWhiteSpace($width) -and $currentStyle -notmatch 'width:') {
                        $newStyleParts += "width:$($width)"
                    }
                }


                $heightAttribute = $image.Attributes['height']

                if ($null -ne $heightAttribute) {
                    $height = $heightAttribute.Value

                    if (-not [string]::IsNullOrWhiteSpace($height) -and $currentStyle -notmatch 'height:') {
                        $newStyleParts += "height:$($height)"
                    }
                }

                if ($newStyleParts.Count -gt 0) {
                    $combinedStyle = $newStyleParts -join ';'

                    if (-not [string]::IsNullOrWhiteSpace($currentStyle)) {
                        $null = $image.SetAttributeValue('style', "$currentStyle;$combinedStyle")
                    } else {
                        $null = $image.SetAttributeValue('style', $combinedStyle)
                    }
                }
            }
        }

        [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $htmlDoc.DocumentNode.OuterHtml)


        try { WatchCatchableExitSignal } catch { }


        if ($MoveCSSInline) {
            Write-Host "$Indent        Move CSS inline"

            $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
            $tempFileContent = ConvertEncoding -InFile $path

            # Use a separate runspace for PreMailer.Net, as there are DLL conflicts in PowerShell 5.x with Invoke-RestMethod
            # Do not use jobs, as they fall back to Constrained Language Mode in secured environments, which makes Import-Module fail
            $MoveCSSInlineResult = MoveCssInline $tempFileContent

            if ($MoveCSSInlineResult.StartsWith('Failed: ')) {
                Write-Host "$Indent          $MoveCSSInlineResult" -ForegroundColor Yellow
            } else {
                [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $MoveCSSInlineResult)
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Remove empty CSS properties from style attributes"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

        $htmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
        $htmlDoc.DisableImplicitEnd = $true
        $htmlDoc.OptionAutoCloseOnEnd = $true
        $htmlDoc.OptionCheckSyntax = $true
        $htmlDoc.OptionEmptyCollection = $true
        $htmlDoc.OptionFixNestedTags = $true

        $htmlDoc.LoadHtml((ConvertEncoding -InFile $path))

        $htmlDocSelectNodeResult = $htmlDoc.DocumentNode.SelectNodes('//*[@style]')

        if ($htmlDocSelectNodeResult) {
            foreach ($node in $htmlDocSelectNodeResult) {
                if (-not [string]::IsNullOrWhiteSpace($node.GetAttributeValue('style', ''))) {
                    $null = $node.SetAttributeValue(
                        'style',
                        $(
                            @(
                                ParseHtmlStyleAttribute ($node.GetAttributeValue('style', '')) | Where-Object { $_.Property } | ForEach-Object {
                                    "$($_.Property): $($_.Value)"
                                }
                            ) -join '; '
                        )
                    )
                }
            }
        }

        [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $htmlDoc.DocumentNode.OuterHtml)

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Add marker to final HTM file"
        $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
        $tempFileContent = (ConvertEncoding -InFile $path)


        # Load the HTML content
        $htmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
        $htmlDoc.DisableImplicitEnd = $true
        $htmlDoc.OptionAutoCloseOnEnd = $true
        $htmlDoc.OptionCheckSyntax = $true
        $htmlDoc.OptionEmptyCollection = $true
        $htmlDoc.OptionFixNestedTags = $true

        $htmlDoc.LoadHtml($tempFileContent)

        # Ensure there's a <head> element to work with
        $headNode = $htmlDoc.DocumentNode.SelectSingleNode('//head')

        if (-not $headNode) {
            $htmlNode = $htmlDoc.DocumentNode.SelectSingleNode('//html')

            if (-not $htmlNode) {
                $htmlNode = $htmlDoc.CreateElement('html')
                $null = $htmlDoc.DocumentNode.AppendChild($htmlNode)
            }

            $headNode = $htmlDoc.CreateElement('head')
            $null = $htmlNode.PrependChild($headNode)
        }

        # Check for the meta tag
        $metaExists = $headNode.SelectSingleNode("meta[@name='data-SignatureFileInfo' and @content='Set-OutlookSignatures']")

        if (-not $metaExists) {
            $meta = $htmlDoc.CreateElement('meta')
            $null = $meta.SetAttributeValue('name', 'data-SignatureFileInfo')
            $null = $meta.SetAttributeValue('content', 'Set-OutlookSignatures')
            $null = $headNode.AppendChild($meta)
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent        Modify connected folder name"

        foreach ($pathConnectedFolderName in $pathConnectedFolderNames) {
            try { WatchCatchableExitSignal } catch { }

            $newFolderName = "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value)).files"

            # Update src attributes in <img> tags
            $htmlDocSelectNodeResult = $htmlDoc.DocumentNode.SelectNodes('//img[@src]')

            if ($htmlDocSelectNodeResult) {
                foreach ($img in $htmlDocSelectNodeResult) {
                    $src = $img.GetAttributeValue('src', '')

                    if ($null -ne $src) {
                        $src = $src.Trim()
                    } else {
                        continue
                    }

                    if ($src -like "$($pathConnectedFolderName)/*") {
                        $null = $img.SetAttributeValue('src', "$($newFolderName)/$($src.Substring($pathConnectedFolderName.Length + 1))")
                    }
                }
            }

            # Rename the folder
            $oldFolderPath = Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath $pathConnectedFolderName

            Rename-Item -LiteralPath $oldFolderPath -NewName $newFolderName -ErrorAction SilentlyContinue
        }

        try { WatchCatchableExitSignal } catch { }

        [SetOutlookSignatures.Common]::WriteAllTextWithEncodingCorrections($path, $htmlDoc.DocumentNode.OuterHtml)

        try { WatchCatchableExitSignal } catch { }

        if (-not $ProcessOOF) {
            if ($EmbedImagesInHtml) {
                Write-Host "$Indent        Embed local images"

                [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($path, $path)
            }
        } else {
            [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($path, ((Join-Path -Path $script:tempDir -ChildPath $Signature.Value)))
        }

        try { WatchCatchableExitSignal } catch { }

        if (-not $ProcessOOF) {
            if ($CreateRtfSignatures) {
                Write-Host "$Indent      Export to RTF format"

                try { WatchCatchableExitSignal } catch { }

                # If possible, use .docx file to avoid problems with MS Information Protection
                $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))
                $script:COMWord.Documents.Open($path, $false, $false, $false, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, 65001) | Out-Null
                $script:COMWord.ActiveDocument.ActiveWindow.View.Type = [Microsoft.Office.Interop.Word.WdViewType]::wdWebView

                try { WatchCatchableExitSignal } catch { }

                $saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatRTF
                $path = $([System.IO.Path]::ChangeExtension($path, '.rtf'))

                try {
                    # Overcome Word security warning when export contains embedded pictures
                    if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                    }

                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    }

                    if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Save
                    $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                    # Restore original security setting
                    Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                } catch {
                    # Restore original security setting after error
                    Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                    Start-Sleep -Seconds 2

                    # Overcome Word security warning when export contains embedded pictures
                    if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
                    }

                    if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
                        $script:WordDisableWarningOnIncludeFieldsUpdate = Get-ItemPropertyValue -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
                    }

                    if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
                        $null = "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
                    }

                    try { WatchCatchableExitSignal } catch { }

                    # Save
                    $script:COMWord.ActiveDocument.SaveAs2($path, $saveFormat, [Type]::Missing, [Type]::Missing, $false)

                    # Restore original security setting
                    Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
                }

                try { WatchCatchableExitSignal } catch { }

                # Mark document as saved to avoid MS Information Protection asking for setting a sensitivity label when closing the document
                # Close the document as conversion to .rtf happens from .htm
                if ($script:COMWord.ActiveDocument.Saved -ne $true) {
                    $script:COMWord.ActiveDocument.Saved = $true
                }

                $script:COMWord.ActiveDocument.ActiveWindow.View.Type = $script:COMWordViewTypeOriginal

                $script:COMWord.ActiveDocument.Close($false, [Type]::Missing, $false)

                # Restore original security setting
                Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null

                try { WatchCatchableExitSignal } catch { }

                Write-Host "$Indent        Shrink RTF file"
                # No need to use ConvertEncoding, as RTF must be encoded in ASCII
                $((Get-Content -LiteralPath $path -Raw -Encoding Ascii) -ireplace '\{\\nonshppict[\s\S]*?\}\}', '') | Set-Content -LiteralPath $path -Encoding Ascii
            }

            try { WatchCatchableExitSignal } catch { }

            if ($CreateTxtSignatures) {
                Write-Host "$Indent      Export to TXT format"

                $path = $([System.IO.Path]::ChangeExtension($path, '.htm'))

                $null = ConvertHtmlToPlainText -InFile $path -OutFile $([System.IO.Path]::ChangeExtension($path, '.txt')) -OutEncoding ([System.Text.Encoding]::Unicode)

                try { WatchCatchableExitSignal } catch { }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        if (-not $ProcessOOF) {
            Write-Host "$Indent      Upload signature to Exchange Online as roaming signature"

            if ($MirrorCloudSignatures -ne $false) {
                if (-not (($BenefactorCircleLicenseFile) -and ($null -ne [SetOutlookSignatures.BenefactorCircle].GetMethod('RoamingSignaturesUpload')))) {
                    Write-Host "$Indent        Cannot upload signature to Exchange Online." -ForegroundColor Green
                    Write-Host "$Indent        The 'MirrorCloudSignatures' feature requires the Benefactor Circle add-on." -ForegroundColor Green
                    Write-Host "$Indent        Visit https://set-outlooksignatures.com/benefactorcircle for details." -ForegroundColor Green
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    $FeatureResult = [SetOutlookSignatures.BenefactorCircle]::RoamingSignaturesUpload()

                    if ($FeatureResult -ne 'true') {
                        Write-Host "$Indent        Error uploading roaming signatures to the cloud." -ForegroundColor Yellow
                        Write-Host "$Indent        $FeatureResult" -ForegroundColor Yellow
                    }
                }
            } else {
                Write-Host "$Indent        Parameter 'MirrorCloudSignatures' is not enabled, skipping task."
            }

            foreach ($SignaturePath in $SignaturePaths) {
                try { WatchCatchableExitSignal } catch { }

                Write-Host "$Indent      Copy signature files to '$SignaturePath'"

                RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))

                foreach ($ConnectedFilesFolderName in $ConnectedFilesFolderNames) {
                    try { WatchCatchableExitSignal } catch { }

                    RemoveItemAlternativeRecurse -LiteralPath ((Join-Path -Path $SignaturePath -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($Signature.value))") + $ConnectedFilesFolderName)
                }

                Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.htm')) -Destination $((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.htm')))) -Force

                try { WatchCatchableExitSignal } catch { }

                if ($EmbedImagesInHtml -eq $false) {
                    if (Test-Path -LiteralPath (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))")) {
                        Copy-Item -LiteralPath (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath "$([System.IO.Path]::ChangeExtension($Signature.value, '.files'))") -Destination $SignaturePath -Force -Recurse
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($CreateRtfSignatures -eq $true) {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.rtf')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))) -Force
                } else {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.rtf')))
                }

                try { WatchCatchableExitSignal } catch { }

                if ($CreateTxtSignatures -eq $true) {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))
                    Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, '.txt')) -Destination ((Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))) -Force
                } else {
                    RemoveItemAlternativeRecurse (Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, '.txt')))
                }

                try { WatchCatchableExitSignal } catch { }

                if ($SignatureFilesWriteProtect.containskey($TemplateIniSettingsIndex)) {
                    Write-Host "$Indent      Write protect signature files"
                    @('.htm', '.rtf', '.txt') | ForEach-Object {
                        $file = Join-Path -Path ($SignaturePath) -ChildPath $([System.IO.Path]::ChangeExtension($Signature.Value, $_))
                        if (Test-Path -LiteralPath $file -PathType Leaf) {
                            (Get-Item -LiteralPath $file -Force).Attributes += 'ReadOnly'
                        }
                    }
                }

                try { WatchCatchableExitSignal } catch { }

                if ($macOSSignaturesScriptable) {
                    Write-Host "$Indent      Create Outlook for Mac signature"

                    @($(@"
tell application "Microsoft Outlook"
    try
        set signatureName to "$(Split-Path -Path $signature.value -LeafBase)"
        set htmlContent to (read POSIX file "$(([System.IO.Path]::ChangeExtension($path, '.htm')))" as «class utf8»)

        -- Check if the signature exists
        set signatureList to signatures
        set signatureExists to false

        repeat with aSignature in signatureList
            if name of aSignature is signatureName then
                set signatureExists to true
                exit repeat
            end if
        end repeat

        if signatureExists then
            -- Update the existing signature
            set content of signature signatureName to htmlContent
        else
            -- Create a new signature
            make new signature with properties {name:signatureName, content:htmlContent}
        end if
    on error errorMessage
        log "$Indent        Error: " & errorMessage
    end try
end tell
"@ | osascript *>&1)) | ForEach-Object { Write-Host $_.tostring() }
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Write-Host "$Indent      Remove temporary files"
        foreach ($extension in ('.docx', '.htm', '.rtf', '.txt')) {
            Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($path, $extension)) -ErrorAction SilentlyContinue | Out-Null

            if ($pathHighResHtml) {
                Remove-Item -LiteralPath $([System.IO.Path]::ChangeExtension($pathHighResHtml, $extension)) -ErrorAction SilentlyContinue | Out-Null
            }
        }

        try { WatchCatchableExitSignal } catch { }

        foreach ($file in @(Get-ChildItem ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($path) + '*') -Directory).FullName) {
            Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
        }

        try { WatchCatchableExitSignal } catch { }

        if ($pathHighResHtml) {
            foreach ($file in @(Get-ChildItem ("$($script:tempDir)\*" + [System.IO.Path]::GetFileNameWithoutExtension($pathHighResHtml) + '*') -Directory).FullName) {
                Remove-Item -LiteralPath $file -Force -Recurse -ErrorAction SilentlyContinue
            }
        }

        try { WatchCatchableExitSignal } catch { }

        Remove-Item -LiteralPath (Join-Path -Path (Split-Path -LiteralPath $path) -ChildPath $([System.IO.Path]::ChangeExtension($signature.value, '.files'))) -Force -Recurse -ErrorAction SilentlyContinue
    }

    try { WatchCatchableExitSignal } catch { }

    if ((-not $ProcessOOF)) {
        # Set default signature for new emails
        if ($SignatureFilesDefaultNew.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                try { WatchCatchableExitSignal } catch { }

                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if ($CurrentTemplateIsForAliasSmtp) {
                        $NewSigExpected."$($CurrentTemplateIsForAliasSmtp.ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }

                    $NewSigExpected."$(($MailAddresses[$AccountNumberRunning]).ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                    if (-not $SimulateUser) {
                        if ($RegistryPaths[$j] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*') {
                            Write-Host "$Indent      Set signature as default for new messages (Outlook profile '$(($RegistryPaths[$j] -split '\\')[8])')"

                            if ($OutlookFileVersion -ge '16.0.0.0') {
                                New-ItemProperty -LiteralPath $RegistryPaths[$j] -Name 'New Signature' -PropertyType String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force | Out-Null
                            } else {
                                New-ItemProperty -LiteralPath $RegistryPaths[$j] -Name 'New Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) + "`0")))) -Force | Out-Null
                            }
                        } else {
                            $script:GraphUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        @('htm', 'rtf', 'txt') | ForEach-Object {
                            if (Test-Path -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))) {
                                $script:GraphUserDummyMailboxDefaultSigNew = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                                if ($_ -ieq 'htm') {
                                    [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")), $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultNew.$($_)")))
                                } else {
                                    Copy-Item -LiteralPath $(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")) -Destination $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultNew.$($_)")) -Force
                                }
                            }
                        }
                    }
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        # Set default signature for replies and forwarded emails
        try { WatchCatchableExitSignal } catch { }

        if ($SignatureFilesDefaultReplyFwd.containskey($TemplateIniSettingsIndex)) {
            for ($j = 0; $j -lt $MailAddresses.count; $j++) {
                try { WatchCatchableExitSignal } catch { }

                if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                    if ($CurrentTemplateIsForAliasSmtp) {
                        $ReplySigExpected."$($CurrentTemplateIsForAliasSmtp.ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                    }

                    $ReplySigExpected."$(($MailAddresses[$AccountNumberRunning]).ToLower())" = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                    if (-not $SimulateUser) {
                        if ($RegistryPaths[$j] -ilike '*\9375CFF0413111d3B88A00104B2A6676\*') {
                            Write-Host "$Indent      Set signature as default for reply/forward messages (Outlook profile '$(($RegistryPaths[$j] -split '\\')[8])')"

                            if ($OutlookFileVersion -ge '16.0.0.0') {
                                New-ItemProperty -LiteralPath $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force | Out-Null
                            } else {
                                New-ItemProperty -LiteralPath $RegistryPaths[$j] -Name 'Reply-Forward Signature' -PropertyType Binary -Value ([byte[]](([System.Text.Encoding]::Unicode.GetBytes(((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')) + "`0")))) -Force | Out-Null
                            }
                        } else {
                            $script:GraphUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')
                        }
                    } else {
                        @('htm', 'rtf', 'txt') | ForEach-Object {
                            if (Test-Path -LiteralPath (Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)"))) {
                                $script:GraphUserDummyMailboxDefaultSigReply = (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.')

                                if ($_ -ieq 'htm') {
                                    [SetOutlookSignatures.Common]::ConvertToSingleFileHtml($(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")), $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultReplyFwd.$($_)")))
                                } else {
                                    Copy-Item -LiteralPath $(Join-Path -Path ($SignaturePaths[0]) -ChildPath ((($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') + ".$($_)")) -Destination $((Join-Path -Path ((New-Item -ItemType Directory -Path (Join-Path -Path ($SignaturePaths[0]) -ChildPath "___Mailbox $($MailAddresses[$AccountNumberRunning])\") -Force).fullname) -ChildPath "DefaultReplyFwd.$($_)")) -Force
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function CheckADConnectivity {
    param (
        [array]$CheckDomains,
        [string]$CheckProtocolText,
        [string]$Indent
    )

    try { WatchCatchableExitSignal } catch { }

    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 25)
    $RunspacePool.Open()

    for ($DomainNumber = 0; $DomainNumber -lt $CheckDomains.count; $DomainNumber++) {
        try { WatchCatchableExitSignal } catch { }

        if ($($CheckDomains[$DomainNumber]) -eq '') {
            continue
        }

        $PowerShell = [powershell]::Create()
        $PowerShell.RunspacePool = $RunspacePool

        [void]$PowerShell.AddScript({
                param (
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
                    try { WatchCatchableExitSignal } catch { }
                    $null = ([ADSI]"$(($Search.FindOne()).path)")
                    try { WatchCatchableExitSignal } catch { }
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
        try { WatchCatchableExitSignal } catch { }

        foreach ($job in $script:jobs) {
            try { WatchCatchableExitSignal } catch { }

            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -imatch 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -ireplace '[^0-9]'
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
                        Write-Host "$Indent  If this error is permanent, check firewalls, DNS and AD trust. Consider parameter 'TrustsToCheckForGroups' to not use this domain." -ForegroundColor Red

                        if ($TrustsToCheckForGroups -icontains $data[0]) {
                            $TrustsToCheckForGroups.remove($data[0])
                        }

                        $LookupDomainsToTrusts.remove($data[0])

                        $returnvalue = $false
                    }

                    $job.Done = $true
                }
            }
        }

        Start-Sleep -Seconds 1
    }

    try { WatchCatchableExitSignal } catch { }

    return $returnvalue
}


function ConvertEncoding {
    [CmdletBinding()]

    param (
        [Parameter()]
        $InFile = $null,

        [Parameter()]
        $InString = $null,

        [Parameter()]
        $InEncoding = $null,

        [Parameter()]
        $OutFile = $null,

        [Parameter()]
        $OutEncoding = [System.Text.UTF8Encoding]::new($false),

        [Parameter()]
        [bool]$InIsHtml = $true
    )


    # To use the systems default codepage, use one of the following values for $InEncoding or $OutEncoding:
    #   ([System.Text.Encoding]::Default)
    #   ([System.Text.Encoding]::GetEncoding(0))
    #   ([System.Text.Encoding]::GetEncoding($null))
    # UTF8 without BOM
    #   ([System.Text.UTF8Encoding]::new($false))


    Write-Verbose 'ConvertEncoding Start'


    try { WatchCatchableExitSignal } catch { }


    if ($InString) {
        try {
            $InFileBytes = ([System.Text.UTF8Encoding]::new($false)).GetBytes("$($InString -join [Environment]::NewLine)")
            Write-Verbose '  InString: Converted to bytes using UTF-8 encoding.'
        } catch {
            Write-Verbose "  InString: Error converting to bytes: $($_)"
            return
        }
    } else {
        try {
            $InFile = (Resolve-Path -LiteralPath $InFile).ProviderPath

            try {
                try { WatchCatchableExitSignal } catch { }

                $InFileBytes = [System.IO.File]::ReadAllBytes($InFile)

                Write-Verbose "  InFile: '$($InFile)'"
            } catch {
                Write-Verbose "  InFile: Error reading '$($InFile)': $($_)"

                Write-Verbose 'ConvertEncoding End'

                return
            }
        } catch {
            Write-Verbose "  InFile: '$($InFile)' not found: $($_)"

            Write-Verbose 'ConvertEncoding End'

            return
        }
    }

    if ($InEncoding) {
        if (-not ($InEncoding -is [System.Text.Encoding])) {
            if (-not [string]::IsNullOrWhiteSpace($InEncoding.ToString())) {
                try {
                    . ([System.Management.Automation.ScriptBlock]::Create("`$InEncoding = $(@([System.Text.Encoding] | Get-Member -Static -MemberType Property | ForEach-Object { "[$($_.TypeName)]::$($_.Name)" }) | Where-Object { $_ -ieq $InEncoding.ToString() })"))
                } catch {
                    try {
                        $InEncoding = [System.Text.Encoding]::GetEncoding($InEncoding.ToString())
                    } catch {
                    }
                }
            }

            if (-not ($InEncoding -is [System.Text.Encoding])) {
                throw "InEncoding: WebName '$($InEncoding)' not found. Exiting."
            }
        }
    }

    Write-Verbose "  InEncoding: WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"


    Write-Verbose "  InIsHtml: $($InIsHtml)"


    if ($OutFile) {
        try {
            if ([System.IO.Path]::IsPathRooted($OutFile)) {
                $OutFile = [System.IO.Path]::GetFullPath($OutFile)
            } else {
                $OutFile = [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location) -ChildPath $OutFile))
            }
        } catch {
            Write-Verbose "  OutFile: '$($OutFile)' error: $($_)"

            Write-Verbose 'ConvertEncoding End'

            return
        }
    }

    Write-Verbose "  OutFile: '$($OutFile)'"


    if ($OutEncoding) {
        if (-not ($OutEncoding -is [System.Text.Encoding])) {
            if (-not [string]::IsNullOrWhiteSpace($OutEncoding.ToString())) {
                try {
                    . ([System.Management.Automation.ScriptBlock]::Create("`$OutEncoding = $(@([System.Text.Encoding] | Get-Member -Static -MemberType Property | ForEach-Object { "[$($_.TypeName)]::$($_.Name)" }) | Where-Object { $_ -ieq $OutEncoding.ToString() })"))
                } catch {
                    try {
                        $OutEncoding = [System.Text.Encoding]::GetEncoding($OutEncoding.ToString())
                    } catch {
                    }
                }
            }

            if (-not ($OutEncoding -is [System.Text.Encoding])) {
                throw "OutEncoding: WebName '$($OutEncoding)' not found. Exiting."
            }
        }
    }

    Write-Verbose "  OutEncoding: WebName '$($OutEncoding.WebName)', WindowsCodePage '$($OutEncoding.WindowsCodePage)', CodePage '$($OutEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $OutEncoding.GetPreamble() } catch { , @() })))'"


    # $InEncoding has not been defined, so we detect it
    # Check for BOM (Byte Order Mark)
    if (-not $InEncoding) {
        try { WatchCatchableExitSignal } catch { }

        foreach ($encodingInfo in [System.Text.Encoding]::GetEncodings()) {
            try {
                $encoding = $encodingInfo.GetEncoding()
                $preamble = $encoding.GetPreamble()

                if ($preamble.Length -gt 0 -and $InFileBytes.Length -ge $preamble.Length) {
                    $fileStart = $InFileBytes[0..($preamble.Length - 1)]

                    if ([BitConverter]::ToString($fileStart) -ceq [BitConverter]::ToString($preamble)) {
                        $InEncoding = $encoding
                        break
                    }
                }
            } catch {
                # Some encodings may throw exceptions (e.g., unsupported code pages)
                continue
            }
        }

        Write-Verbose "  InEncoding: BOM: WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"
    }

    # No BOM detected, so we need to go deeper
    if (-not $InEncoding) {
        # If $InIsHtml, check HTML meta tags for encoding information
        if ($InIsHtml -eq $true) {
            try { WatchCatchableExitSignal } catch { }

            # Initialize Html Agility Pack
            $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
            # Use default settings to keep HTML as original as possible
            # $htmlDoc.DisableImplicitEnd = $true
            # $htmlDoc.OptionAutoCloseOnEnd = $true
            # $htmlDoc.OptionCheckSyntax = $true
            # $htmlDoc.OptionEmptyCollection = $true
            # $htmlDoc.OptionFixNestedTags = $true

            $htmlDoc.LoadHtml(([System.Text.UTF8Encoding]::new($false)).GetString($InFileBytes))

            $charset = $null

            # Get all meta tags
            $htmlDocSelectNodeResult = $htmlDoc.DocumentNode.SelectNodes('//meta')

            # First, try to find <meta charset="...">
            if ($htmlDocSelectNodeResult) {
                foreach ($meta in $htmlDocSelectNodeResult) {
                    if ($meta.Attributes['charset']) {
                        $charset = $meta.Attributes['charset'].Value

                        try {
                            $InEncoding = [System.Text.Encoding]::GetEncoding($charset)
                        } catch {
                            $InEncoding = $null
                        }

                        break
                    }
                }

                Write-Verbose "  InEncoding: HTML charset '$($charset)': WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"


                # Fallback: Try to find <meta http-equiv="Content-Type" content="...; charset=...">
                if (-not $InEncoding) {
                    foreach ($meta in $htmlDocSelectNodeResult) {
                        $httpEquiv = $meta.GetAttributeValue('http-equiv', '')

                        if ($httpEquiv -ieq 'content-type') {
                            $content = $meta.GetAttributeValue('content', '')

                            if ($content -imatch 'charset=([\w-]+)') {
                                $charset = $matches[1]

                                try {
                                    $InEncoding = [System.Text.Encoding]::GetEncoding($charset)
                                } catch {
                                    $InEncoding = $null
                                }

                                break
                            }
                        }
                    }
                }

                Write-Verbose "  InEncoding: HTML http-equiv content '$($charset)': WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"
            }
        }

        # No BOM, no info from HTML, so we need to detect the encoding heuristically
        if (-not $InEncoding) {
            try { WatchCatchableExitSignal } catch { }

            $InEncoding = [UtfUnknown.CharsetDetector]::DetectFromBytes($InFileBytes).Detected.Encoding

            Write-Verbose "  InEncoding: Heuristics: WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"
        }

        # BOM has already been checked before, so switch to bomless if possible
        if ($InEncoding -and ($InEncoding.GetPreamble().Length -gt 0)) {
            try { WatchCatchableExitSignal } catch { }

            foreach ($encodingInfo in [System.Text.Encoding]::GetEncodings()) {
                $encoding = $encodingInfo.GetEncoding()

                if (($encoding.CodePage -eq $InEncoding.CodePage) -and ($encoding.GetPreamble().Length -eq 0)) {
                    $InEncoding = $encoding
                    break
                }
            }

            if ($InEncoding.GetPreamble().Length -gt 0) {
                # Try to find a constructor that can create a bomless version
                $encodingType = $InEncoding.GetType()
                $constructors = $encodingType.GetConstructors()

                foreach ($constructor in $constructors) {
                    $parameters = $constructor.GetParameters()

                    $emitBOMParameter = $parameters | Where-Object {
                        $($_.ParameterType -eq [bool]) -and
                        $(($_.Name -eq 'encoderShouldEmitUTF8Identifier') -or ($_.Name -eq 'emitBOM'))
                    }

                    if ($emitBOMParameter) {
                        $constructorArguments = @()

                        foreach ($param in $parameters) {
                            if ($param.Name -eq 'encoderShouldEmitUTF8Identifier' -or $param.Name -eq 'emitBOM') {
                                $constructorArguments += $false
                            } else {
                                try {
                                    $propertyValue = $InEncoding | Select-Object -ExpandProperty $param.Name -ErrorAction SilentlyContinue

                                    if ($null -ne $propertyValue) {
                                        $constructorArguments += $propertyValue
                                    } else {
                                        $constructorArguments += $null # Placeholder, might cause issues
                                    }
                                } catch {
                                    $constructorArguments += $null
                                }
                            }
                        }

                        try {
                            $bomlessEncoding = [System.Activator]::CreateInstance($encodingType, $constructorArguments)

                            if ($bomlessEncoding.GetPreamble().Length -eq 0) {
                                $InEncoding = $bomlessEncoding
                            }
                        } catch {
                        }
                    }
                }
            }

            Write-Verbose "  InEncoding: To BOM-less: WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"
        }
    }

    # If we still don't have an encoding, we need to fallback to UTF-8 without BOM
    if (-not $InEncoding) {
        $InEncoding = [System.Text.UTF8Encoding]::new($false)

        Write-Verbose "  InEncoding: Fallback: WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"
    }

    Write-Verbose "  InEncoding: Final: WebName '$($InEncoding.WebName)', WindowsCodePage '$($InEncoding.WindowsCodePage)', CodePage '$($InEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $InEncoding.GetPreamble() } catch { , @() })))'"


    try { WatchCatchableExitSignal } catch { }


    # Strip BOM from $InFileBytes if present, so it is not accidently added later
    # Convert from bytes and $InEncoding to string and UTF-8
    # Ensures that we have a .Net string object
    # and that HTML Agility Pack can work with a UTF8 string
    $preamble = $InEncoding.GetPreamble()

    if ($preamble.Length -gt 0 -and $InFileBytes.Length -ge $preamble.Length) {
        $fileStart = $InFileBytes[0..($preamble.Length - 1)]
        if ([BitConverter]::ToString($fileStart) -ceq [BitConverter]::ToString($preamble)) {
            $InFileBytes = $InFileBytes[$preamble.Length..($InFileBytes.Length - 1)]
        }
    }

    $OutString = ([System.Text.UTF8Encoding]::new($false)).GetString(
        [System.Text.Encoding]::Convert(
            $InEncoding,
            ([System.Text.UTF8Encoding]::new($false)),
            $InFileBytes
        )
    )


    # Modify HTML to reflect the new encoding
    if ($InIsHtml -eq $true) {
        $htmlDoc.LoadHtml($OutString)

        # Desired encoding
        $newCharset = $OutEncoding.WebName

        # Find or create <head>
        $headNode = $htmlDoc.DocumentNode.SelectSingleNode('//head')

        if (-not $headNode) {
            $headNode = $htmlDoc.CreateElement('head')
            $null = $htmlDoc.DocumentNode.PrependChild($headNode)
        }

        # Update or insert <meta http-equiv="Content-Type">
        $metaHttpEquiv = $htmlDoc.DocumentNode.SelectSingleNode("//meta[@http-equiv='Content-Type']")

        if ($metaHttpEquiv) {
            $null = $metaHttpEquiv.SetAttributeValue('content', "text/html; charset=$newCharset")
        } else {
            $newMeta = $htmlDoc.CreateElement('meta')
            $null = $newMeta.SetAttributeValue('http-equiv', 'Content-Type')
            $null = $newMeta.SetAttributeValue('content', "text/html; charset=$newCharset")
            $null = $headNode.AppendChild($newMeta)
        }

        # Update or insert <meta charset="...">
        $metaCharset = $htmlDoc.DocumentNode.SelectSingleNode('//meta[@charset]')

        if ($metaCharset) {
            $null = $metaCharset.SetAttributeValue('charset', $newCharset)
        } else {
            $newCharsetMeta = $htmlDoc.CreateElement('meta')
            $null = $newCharsetMeta.SetAttributeValue('charset', $newCharset)
            $null = $headNode.AppendChild($newCharsetMeta)
        }

        # Update $OutString
        $OutString = $htmlDoc.DocumentNode.OuterHtml
    }


    # Convert OutString from UTF-8 to target encoding
    $OutString = $OutEncoding.GetString(
        [System.Text.Encoding]::Convert(
            ([System.Text.UTF8Encoding]::new($false)),
            $OutEncoding,
            ([System.Text.UTF8Encoding]::new($false)).GetBytes("$($OutString)")
        )
    )


    # Save to file if $OutFile is specified
    if ($OutFile) {
        try { WatchCatchableExitSignal } catch { }

        try {
            if (Test-Path -LiteralPath $OutFile) {
                Remove-Item -LiteralPath $OutFile -Force
            }

            # As WriteAllText works with .Net Strings only, the input encoding does not need to be (and can not be) defined
            [System.IO.File]::WriteAllText($OutFile, $OutString, $OutEncoding)

            Write-Verbose "  OutFile: Success writing '$($OutFile)'"
        } catch {
            Write-Verbose "  OutFile: Error writing '$($OutFile)': $($_)"

            Write-Verbose 'ConvertEncoding End'

            return
        }
    }

    Write-Verbose 'ConvertEncoding End'

    return $OutString
}


function ConvertHtmlToPlainText {
    [CmdletBinding()]

    param (
        [Parameter()]
        $InFile = $null,

        [Parameter()]
        $InString = $null,

        [Parameter()]
        $InEncoding = $null,

        [Parameter()]
        $OutFile = $null,

        [Parameter()]
        $OutEncoding = $null
    )


    try { WatchCatchableExitSignal } catch { }


    if ($OutEncoding) {
        if (-not ($OutEncoding -is [System.Text.Encoding])) {
            if (-not [string]::IsNullOrWhiteSpace($OutEncoding.ToString())) {
                try {
                    . ([System.Management.Automation.ScriptBlock]::Create("`$OutEncoding = $(@([System.Text.Encoding] | Get-Member -Static -MemberType Property | ForEach-Object { "[$($_.TypeName)]::$($_.Name)" }) | Where-Object { $_ -ieq $OutEncoding.ToString() })"))
                } catch {
                    try {
                        $OutEncoding = [System.Text.Encoding]::GetEncoding($OutEncoding.ToString())
                    } catch {
                    }
                }
            }

            if (-not ($OutEncoding -is [System.Text.Encoding])) {
                throw "OutEncoding WebName '$($OutEncoding)' not found. Exiting."
            }
        }
    } else {
        # Default to UTF-8 without BOM
        $OutEncoding = [System.Text.UTF8Encoding]::new($false)
    }

    Write-Verbose "OutEncoding: WebName '$($OutEncoding.WebName)', WindowsCodePage '$($OutEncoding.WindowsCodePage)', CodePage '$($OutEncoding.CodePage)', Preamble/BOM '$([BitConverter]::ToString($(try{ , $OutEncoding.GetPreamble() } catch { , @() })))'"


    $ConvertEncodingParams = @{
        InFile      = $InFile
        InString    = $InString
        InEncoding  = $InEncoding
        OutFile     = $null
        OutEncoding = ([System.Text.Encoding]::Unicode) # This ist what the Sytem.Text.StringBuilder expects
    }

    $htmlContent = ConvertEncoding @ConvertEncodingParams


    enum ConvertHtmlToPlainTextToPlainTextState {
        StartLine = 0
        NotWhiteSpace
        WhiteSpace
    }


    $ConvertHtmlToPlainTextInlineTags = @(
        'b', 'big', 'i', 'small', 'tt', 'abbr', 'acronym',
        'cite', 'code', 'dfn', 'em', 'kbd', 'strong', 'samp',
        'var', 'a', 'bdo', 'br', 'img', 'map', 'object', 'q',
        'script', 'span', 'sub', 'sup', 'button', 'input', 'label',
        'select', 'textarea'
    )


    $ConvertHtmlToPlainTextNonVisibleTags = @(
        'script', 'style'
    )


    function ConvertHtmlToPlainTextIsHardSpace($ch) {
        return $(
            $($ch -eq 0xA0) -or
            $($ch -eq 0x2007) -or
            $($ch -eq 0x202F)
        )
    }


    function ConvertHtmlToPlainTextProcessText([ref]$builder, [ref]$state, [char[]]$chars) {
        foreach ($ch in $chars) {
            if ([string]::IsNullOrWhiteSpace($ch)) {
                if (ConvertHtmlToPlainTextIsHardSpace $ch) {
                    if ($state.Value -eq [ConvertHtmlToPlainTextToPlainTextState]::WhiteSpace) {
                        $builder.Value.Append(' ') | Out-Null
                    }

                    $builder.Value.Append(' ') | Out-Null
                    $state.Value = [ConvertHtmlToPlainTextToPlainTextState]::NotWhiteSpace
                } else {
                    if ($state.Value -eq [ConvertHtmlToPlainTextToPlainTextState]::NotWhiteSpace) {
                        $state.Value = [ConvertHtmlToPlainTextToPlainTextState]::WhiteSpace
                    }
                }
            } else {
                if ($state.Value -eq [ConvertHtmlToPlainTextToPlainTextState]::WhiteSpace) {
                    $builder.Value.Append(' ') | Out-Null
                }

                $builder.Value.Append($ch) | Out-Null
                $state.Value = [ConvertHtmlToPlainTextToPlainTextState]::NotWhiteSpace
            }
        }
    }


    function ConvertHtmlToPlainTextProcessNodes([ref]$builder, [ref]$state, $nodes) {
        foreach ($node in $nodes) {
            if ($node -is [HtmlAgilityPack.HtmlTextNode]) {
                $text = [HtmlAgilityPack.HtmlEntity]::DeEntitize($node.Text)

                ConvertHtmlToPlainTextProcessText -builder $builder -state $state -chars $text.ToCharArray()
            } else {
                $tag = $node.Name

                if ($tag -ieq 'br') {
                    $builder.Value.AppendLine() | Out-Null
                    $state.Value = [ConvertHtmlToPlainTextToPlainTextState]::StartLine
                } elseif ($ConvertHtmlToPlainTextNonVisibleTags -icontains $tag) {
                    continue
                } elseif ($ConvertHtmlToPlainTextInlineTags -icontains $tag) {
                    ConvertHtmlToPlainTextProcessNodes -builder $builder -state $state -nodes $node.ChildNodes
                } else {
                    if ($state.Value -ne [ConvertHtmlToPlainTextToPlainTextState]::StartLine) {
                        $builder.Value.AppendLine() | Out-Null
                        $state.Value = [ConvertHtmlToPlainTextToPlainTextState]::StartLine
                    }

                    ConvertHtmlToPlainTextProcessNodes -builder $builder -state $state -nodes $node.ChildNodes

                    if ($state.Value -ne [ConvertHtmlToPlainTextToPlainTextState]::StartLine) {
                        $builder.Value.AppendLine() | Out-Null
                        $state.Value = [ConvertHtmlToPlainTextToPlainTextState]::StartLine
                    }
                }
            }
        }
    }


    try { WatchCatchableExitSignal } catch { }


    $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
    $htmlDoc.DisableImplicitEnd = $true
    $htmlDoc.OptionAutoCloseOnEnd = $true
    $htmlDoc.OptionCheckSyntax = $true
    $htmlDoc.OptionEmptyCollection = $true
    $htmlDoc.OptionFixNestedTags = $true

    $htmlDoc.LoadHtml($htmlContent)


    try { WatchCatchableExitSignal } catch { }


    $builder = New-Object System.Text.StringBuilder
    $state = [ConvertHtmlToPlainTextToPlainTextState]::StartLine

    ConvertHtmlToPlainTextProcessNodes -builder ([ref]$builder) -state ([ref]$state) -nodes @($htmlDoc.DocumentNode)


    try { WatchCatchableExitSignal } catch { }


    # Convert from System.Text.StringBuilder UTF16 to target encoding
    $htmlContent = $OutEncoding.GetString(
        [System.Text.Encoding]::Convert(
            [System.Text.Encoding]::Unicode,
            $OutEncoding,
            [System.Text.Encoding]::Unicode.GetBytes($builder.ToString())
        )
    )


    # Save to file if $OutFile is specified
    if ($OutFile) {
        try {
            try { WatchCatchableExitSignal } catch { }

            [System.IO.File]::WriteAllText($OutFile, $htmlContent, $OutEncoding)
        } catch {
            Write-Verbose "Error writing to file '$($OutFile)': $($_)"

            return
        }
    }


    return $htmlContent
}


function ParseHtmlStyleAttribute {
    param (
        [Parameter(Mandatory = $true)]
        [string]$StyleString
    )

    # Initialize result array
    $properties = @()

    # Decode HTML entities if present
    $decodedStyle = [System.Net.WebUtility]::HtmlDecode($StyleString)

    # State variables
    $currentProperty = ''
    $currentValue = ''
    $inValue = $false
    $inQuote = $false
    $quoteChar = ''
    $parenCount = 0

    # Process character by character
    $chars = $decodedStyle.ToCharArray()

    for ($i = 0; $i -lt $chars.Length; $i++) {
        $char = $chars[$i]

        # Handle quotes
        if (($char -eq '"' -or $char -eq "'") -and $chars[$i - 1] -ne '\') {
            if ($inQuote) {
                if ($char -eq $quoteChar) {
                    $inQuote = $false
                    $currentValue += $char
                    continue
                }
            } else {
                $inQuote = $true
                $quoteChar = $char
                $currentValue += $char
                continue
            }
        }

        # Handle parentheses
        if ($char -eq '(' -and -not $inQuote) {
            $parenCount++
            $currentValue += $char
            continue
        }
        if ($char -eq ')' -and -not $inQuote) {
            $parenCount--
            $currentValue += $char
            continue
        }

        # Property-value separator
        if ($char -eq ':' -and -not $inQuote -and $parenCount -eq 0 -and -not $inValue) {
            $inValue = $true
            continue
        }

        # Property separator
        if ($char -eq ';' -and -not $inQuote -and $parenCount -eq 0) {
            if ($currentProperty -and $currentValue) {
                $properties += [PSCustomObject]@{
                    Property = $currentProperty.Trim().ToLower()
                    Value    = [System.Net.WebUtility]::HtmlEncode($currentValue.Trim())
                }
            }
            $currentProperty = ''
            $currentValue = ''
            $inValue = $false
            continue
        }

        # Add character to current property or value
        if ($inValue) {
            $currentValue += $char
        } else {
            $currentProperty += $char
        }
    }

    # Add final property if exists
    if ($currentProperty -and $currentValue) {
        $properties += [PSCustomObject]@{
            Property = $currentProperty.Trim().ToLower()
            Value    = [System.Net.WebUtility]::HtmlEncode($currentValue.Trim())
        }
    }

    return $properties
}


function GetHtmlBody {
    param (
        [string]$htmlContent = ''
    )


    try { WatchCatchableExitSignal } catch { }


    $htmlDoc = [HtmlAgilityPack.HtmlDocument]::new()
    $htmlDoc.DisableImplicitEnd = $true
    $htmlDoc.OptionAutoCloseOnEnd = $true
    $htmlDoc.OptionCheckSyntax = $true
    $htmlDoc.OptionEmptyCollection = $true
    $htmlDoc.OptionFixNestedTags = $true

    $htmlDoc.LoadHtml($htmlContent)

    $bodyNode = $htmlDoc.DocumentNode.SelectSingleNode('//body')

    if ($bodyNode) {
        $bodyHtml = $bodyNode.InnerHtml
    } else {
        $bodyChildren = $htmlDoc.DocumentNode.ChildNodes | Where-Object {
            $($_.Name -ine 'head') -and
            $($_.Name -ine 'html') -and
            $($_.NodeType -ieq 'Element')
        }

        $bodyHtml = ($bodyChildren | ForEach-Object { $_.OuterHtml }) -join "`n"
    }

    return $bodyHtml
}


function MoveCssInline {
    param (
        $HtmlCode
    )

    try { WatchCatchableExitSignal } catch { }

    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, 1)
    $RunspacePool.Open()

    $PowerShell = [powershell]::Create()
    $PowerShell.RunspacePool = $RunspacePool

    [void]$PowerShell.AddScript({
            param (
                $HtmlCode,
                $path
            )

            $assemblyResolveHandler = $null
            $currentlyResolving = @{}

            function EnableAssemblyResolver {
                if ($null -ne $assemblyResolveHandler) {
                    return
                }

                $assemblyResolveHandler = {
                    param($senderDetails, $arguments)

                    $assemblyName = [System.Reflection.AssemblyName]::new($arguments.Name).Name

                    if ($assemblyName -like '*.resources') {
                        return $null
                    }


                    if ($currentlyResolving.ContainsKey($assemblyName)) {
                        return $null
                    }

                    $currentlyResolving[$assemblyName] = $true

                    try {
                        $dllPath = Join-Path -Path $path -ChildPath "$($assemblyName).dll"
                        if (Test-Path -LiteralPath $dllPath) {
                            return [System.Reflection.Assembly]::LoadFrom($dllPath)
                        }
                    } catch {
                        Write-Debug "Failed: Load '$($assemblyName)' from '$($dllPath)'."
                    } finally {
                        $currentlyResolving.Remove($assemblyName) | Out-Null
                    }

                    try {
                        return [System.Reflection.Assembly]::Load($assemblyName)
                    } catch {
                        Write-Debug "Failed: Default load for '$($assemblyName)'."
                    }

                    return $null
                }


                [System.AppDomain]::CurrentDomain.add_AssemblyResolve($assemblyResolveHandler)
            }

            function DisableAssemblyResolver {
                if ($null -eq $assemblyResolveHandler) {
                    return
                }

                [System.AppDomain]::CurrentDomain.remove_AssemblyResolve($assemblyResolveHandler)

                $assemblyResolveHandler = $null
            }

            if ($($PSVersionTable.PSEdition) -ine 'Core') {
                EnableAssemblyResolver
            }

            $DebugPreference = 'Continue'
            Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"

            try {
                Import-Module (Join-Path -Path $path -ChildPath 'PreMailer.Net.dll') -Force -ErrorAction Stop

                $PreMailer = [PreMailer.Net.PreMailer]::New(
                    $HtmlCode, # string html
                    $null # uri baseUri
                )

                Write-Debug $(
                    $PreMailer.MoveCssInline(
                        $true, # bool removeStyleElements = False
                        $null, # string ignoreElements = null
                        $null, # string css = null
                        $true, # bool stripIdAndClassAttributes = False
                        $false, # bool removeComments = False
                        $null, # AngleSharp.IMarkupFormatter customFormatter = null
                        $false, # bool preserveMediaQueries = False
                        $false # bool useEmailFormatter = False # Messes up Outlook formatting!
                    ).html
                )
            } catch {
                Write-Debug "Failed: $($_ | Format-List * | Out-String)"
            }

            if ($($PSVersionTable.PSEdition) -ine 'Core') {
                DisableAssemblyResolver
            }
        }).AddArgument($HtmlCode).AddArgument($script:PreMailerNetModulePath)

    $Object = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $Handle = $PowerShell.BeginInvoke($Object, $Object)
    $temp = '' | Select-Object PowerShell, Handle, Object, StartTime, Done
    $temp.PowerShell = $PowerShell
    $temp.Handle = $Handle
    $temp.Object = $Object
    $temp.StartTime = $null
    $temp.Done = $false
    [void]$script:jobs.Add($Temp)

    while (($script:jobs.Done | Where-Object { $_ -eq $false }).count -ne 0) {
        try { WatchCatchableExitSignal } catch { }

        foreach ($job in $script:jobs) {
            try { WatchCatchableExitSignal } catch { }

            if (($null -eq $job.StartTime) -and ($job.Powershell.Streams.Debug[0].Message -imatch 'Start')) {
                $StartTicks = $job.powershell.Streams.Debug[0].Message -ireplace '[^0-9]'
                $job.StartTime = [Datetime]::MinValue + [TimeSpan]::FromTicks($StartTicks)
            }

            if ($null -ne $job.StartTime) {
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -ge 5))) {
                    $data = $job.Object[0..$(($job.object).count - 1)]
                    #if ($job.Powershell.Streams.Debug[1].Message.StartsWith('Failed: ')) {
                    #    $returnvalue = $HtmlCode
                    #} else {
                    $returnvalue = $job.Powershell.Streams.Debug[1].Message
                    #}
                    $job.Done = $true
                }
            }
        }

        Start-Sleep -Seconds 1
    }

    try { WatchCatchableExitSignal } catch { }

    return $returnvalue
}


$CheckPathScriptblock = {
    # A script block runs in the scope of the caller, which is different from functions
    # This makes it interesting for manipulating variables, so take care of variable names
    [cmdletbinding()]
    param (
        [ref]$CheckPathRefPath,
        [switch]$CheckPathSilent = $false,
        [switch]$CheckPathCreate = $false,
        [string]$ExpectedPathType = 'Container'
    )

    try { WatchCatchableExitSignal } catch { }

    $CheckPathPath = $CheckPathRefPath.Value

    try {
        Write-Verbose "      Execute config file '$GraphConfigFile'"

        $GraphClientIDOld = $GraphClientID

        if (Test-Path -LiteralPath $GraphConfigFile -PathType Leaf) {
            . ([System.Management.Automation.ScriptBlock]::Create((ConvertEncoding -InFile $GraphConfigFile -InIsHtml $false)))
        } elseif (Test-Path -LiteralPath $(Join-Path -Path $PSScriptRoot -ChildPath '.\config\default graph config.ps1') -PathType Leaf) {
            Write-Verbose '        Not accessible, use default Graph config file'
            . ([System.Management.Automation.ScriptBlock]::Create((ConvertEncoding -InFile $(Join-Path -Path $PSScriptRoot -ChildPath '.\config\default graph config.ps1') -InIsHtml $false)))
        } else {
            Write-Verbose '        Not accessible, and default Graph config file not found'
        }

        if ($GraphClientIDOld -ne $GraphClientID) {
            $GraphClientIDOriginal = $GraphClientID
        }

        GraphSwitchContext -TenantID $null

        try { WatchCatchableExitSignal } catch { }

        @(
            'displayName',
            'givenName',
            'id',
            'mail',
            'mailNickname',
            'onPremisesDistinguishedName',
            'onPremisesDomainName',
            'onPremisesExtensionAttributes',
            'onPremisesImmutableId',
            'onPremisesSamAccountName',
            'onPremisesSecurityIdentifier',
            'onPremisesUserPrincipalName',
            'proxyAddresses',
            'securityIdentifier',
            'surname',
            'userPrincipalName'
        ) | ForEach-Object {
            if ($GraphUserProperties -inotcontains $_) {
                $GraphUserProperties += $_
            }
        }

        if (-not $GraphUserAttributeMapping) {
            $GraphUserAttributeMapping = @{}
        }

        $GraphUserAttributeMapping['distinguishedname'] = 'onPremisesDistinguishedName'
        $GraphUserAttributeMapping['mailboxsettings'] = 'mailboxSettings'
        $GraphUserAttributeMapping['mailNickname'] = 'mailNickname'
        $GraphUserAttributeMapping['objectguid'] = 'id'
        $GraphUserAttributeMapping['objectsid'] = 'securityIdentifier'
        $GraphUserAttributeMapping['onpremisesdomainname'] = 'onPremisesDomainName'
        $GraphUserAttributeMapping['onpremisessecurityidentifier'] = 'onPremisesSecurityIdentifier'
        $GraphUserAttributeMapping['userprincipalname'] = 'userPrincipalName'
    } catch {
        Write-Host $error[0]
        Write-Host "        Problem executing content of '$GraphConfigFile'. Exit." -ForegroundColor Red
        $script:ExitCode = 22
        $script:ExitCodeDescription = 'Problem executing content of GraphConfigFile';
        exit
    }

    try { WatchCatchableExitSignal } catch { }

    if ($CheckPathCreate -eq $false) {
        Write-Verbose "      Try to access '$($CheckPathPath)'."

        if (
            -not $(
                $(
                    (((
                            [uri]$(
                                if (-not [System.Uri]::IsWellFormedUriString($CheckPathPath, [System.UriKind]::Absolute)) {
                                    $([uri]($CheckPathPath -ireplace '@SSL\\', '/' -ireplace '^\\\\', 'https://' -ireplace '\\', '/')).AbsoluteUri
                                } else {
                                    $CheckPathPath
                                }
                            )
                        ).DnsSafeHost -split '\.')[1..999] -join '.') -iin $script:CloudEnvironmentSharePointOnlineDomains
                ) -and
                $GraphClientID
            ) -and
            $(Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)
        ) {
            Write-Verbose "        '$($CheckPathPath)' is accessible, nothing more to do."
        } else {
            Write-Verbose "        '$($CheckPathPath)' is not yet accessible."

            if (-not [System.Uri]::IsWellFormedUriString($CheckPathPath, [System.UriKind]::Absolute)) {
                $CheckPathPath = ([uri]($CheckPathPath -ireplace '@SSL\\', '/' -ireplace '^\\\\', 'https://' -ireplace '\\', '/')).AbsoluteUri
            }

            if (
                (((([uri]$CheckPathPath).DnsSafeHost -split '\.')[1..999] -join '.') -iin $script:CloudEnvironmentSharePointOnlineDomains) -and
                $GraphClientID
            ) {
                # SharePoint Online with Graph client ID
                if (-not $CheckPathSilent) {
                    Write-Host '    SharePoint via Graph, may be slow'
                }

                $CheckPathPath = [uri]::UnescapeDataString($CheckPathPath.Trimend('/'))
                $CheckPathPathSplitBySlash = @($CheckPathPath -split '\/' | Where-Object { $_ })

                try { WatchCatchableExitSignal } catch { }

                # graph auth
                if (-not $script:GraphToken) {
                    GraphGetTokenWrapper -indent '      '
                }

                if ($script:GraphToken.error -eq $false) {
                    Write-Verbose "        Graph Token metadata: $((ParseJwtToken $script:GraphToken.AccessToken) | ConvertTo-Json)"

                    if ($SimulateAndDeployGraphCredentialFile) {
                        Write-Verbose "        Graph Token App metadata: $((ParseJwtToken $script:GraphToken.AppAccessToken) | ConvertTo-Json)"
                    }
                } else {
                    Write-Host '      Problem connecting to Microsoft Graph. Exit.' -ForegroundColor Red
                    Write-Host $script:GraphToken.error -ForegroundColor Red
                    $script:ExitCode = 23
                    $script:ExitCodeDescription = 'Problem connecting to Microsoft Graph.';
                    exit
                }

                if ($SimulateUser) {
                    $script:GraphUser = $SimulateUser
                }

                try { WatchCatchableExitSignal } catch { }

                Write-Verbose '    Get SharePoint Online site ID'

                $(
                    if ($CheckPathPathSplitbySlash[2] -iin @('sites', 'teams')) {
                        "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/sites/$(([uri]$CheckPathPath).DnsSafeHost):/$($CheckPathPathSplitbySlash[2])/$($CheckPathPathSplitbySlash[3])"
                    } else {
                        "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/sites/$(([uri]$CheckPathPath).DnsSafeHost)"
                    }
                ) | ForEach-Object {
                    Write-Verbose "      Query: '$($_)'"

                    $siteId = (GraphGenericQuery -method GET -uri $_ -GraphContext $(([uri]$CheckPathPath).DnsSafeHost) -body $null).result.id

                    Write-Verbose "      siteId: $($siteID)"
                }


                try { WatchCatchableExitSignal } catch { }

                if ($siteid) {
                    Write-Verbose '    Get DocLib drive ID'

                    "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/sites/$($siteId)/drives" | ForEach-Object {
                        $docLibDriveIdQueryResult = (GraphGenericQuery -method GET -uri $_ -GraphContext $(([uri]$CheckPathPath).DnsSafeHost) -body $null).result.value
                        $docLibDriveId = ($docLibDriveIdQueryResult | Where-Object {
                                $_.webUrl -ieq $(
                                    if ($CheckPathPathSplitbySlash[2] -iin @('sites', 'teams')) {
                                        [uri]::EscapeUriString($(($CheckPathPath -split '/')[0..5] -join '/'))
                                    } else {
                                        [uri]::EscapeUriString($(($CheckPathPath -split '/')[0..3] -join '/'))
                                    }
                                )
                            }
                        ).id

                        Write-Verbose "      Query: '$($_)'"
                        Write-Verbose "      Return value: '$(ConvertTo-Json $docLibDriveIdQueryResult -Compress -Depth 10)'"
                        Write-Verbose "      webUrl: '$([uri]::EscapeUriString($(($CheckPathPath -split '/')[0..5] -join '/')))'"

                        Write-Verbose "      docLibDriveId: $docLibDriveId"
                    }

                    try { WatchCatchableExitSignal } catch { }

                    if ($docLibDriveId) {
                        Write-Verbose '      Get DocLib drive items'
                        $docLibDriveItems = (GraphGenericQuery -method GET -uri "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/drives/$($docLibDriveId)/list/items?`$expand=DriveItem" -GraphContext $(([uri]$CheckPathPath).DnsSafeHost) -body $null).result.value

                        $tempDir = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid)))
                        $null = New-Item $tempDir -ItemType Directory

                        $docLibDriveItem = $docLibDriveItems | Where-Object { ([uri]($_.webUrl)).AbsoluteUri -eq ([uri]($CheckPathPath)).AbsoluteUri }

                        if ($docLibDriveItem) {
                            if ($docLibDriveItem.driveItem.file) {
                                Write-Verbose '    Download file to local temp folder'

                                $CheckPathPathNew = $(Join-Path -Path $tempDir -ChildPath $([uri]::UnEscapeDataString((Split-Path -Path $docLibDriveItem.webUrl -Leaf))))

                                $(New-Object Net.WebClient).DownloadFile(
                                    $docLibDriveItem.driveItem.'@microsoft.graph.downloadUrl',
                                    $CheckPathPathNew
                                )

                                Write-Verbose "      '$($CheckPathRefPath.Value)' -> '$($CheckPathPathNew)'"
                                $CheckPathPath = $CheckPathRefPath.Value = $CheckPathPathNew
                            } elseif ($docLibDriveItem.driveItem.folder) {
                                Write-Verbose '    Create temp folders locally'

                                @(
                                    @($docLibDriveItems | Where-Object { ($_.driveItem.folder) -and ($_.webUrl -ilike "$([uri]::EscapeUriString($CheckPathPath))/*") }).webUrl | ForEach-Object {
                                        [uri]::UnescapeDataString(($_ -ireplace "^$([uri]::EscapeUriString($CheckPathPath))/", '')) -replace '/', '\'
                                    }
                                ) | Sort-Object -Culture 127 | ForEach-Object {
                                    if (-not (Test-Path -LiteralPath (Join-Path -Path $tempDir -ChildPath $_) -PathType Container)) {
                                        $null = New-Item -ItemType Directory -Path (Join-Path -Path $tempDir -ChildPath $_)
                                    }
                                }

                                Write-Verbose '      Create dummy files in local temp folders'
                                @($docLibDriveItems | Where-Object { ($_.driveItem.file) -and ($_.webUrl -ilike "$([uri]::EscapeUriString($CheckPathPath))/*") }) | Sort-Object -Culture 127 -Property { $_.webUrl } | ForEach-Object {
                                    $CheckPathPathNew = $(Join-Path -Path $tempDir -ChildPath ([uri]::UnescapeDataString(($_.webUrl -ireplace "^$([uri]::EscapeUriString($CheckPathPath))/", '')) -replace '/', '\'))

                                    if (-not $script:SpoDownloadUrls) {
                                        $script:SpoDownloadUrls = @{}
                                    }

                                    $script:SpoDownloadUrls.Add(
                                        $CheckPathPathNew,
                                        $_.driveItem.'@microsoft.graph.downloadUrl'
                                    )

                                    $null = New-Item -Path $CheckPathPathNew -ItemType File
                                }

                                Write-Verbose "      '$($CheckPathRefPath.Value)' -> '$($tempDir)'"
                                $CheckPathPath = $CheckPathRefPath.Value = $tempDir
                            }
                        } else {
                            Write-Host " '$($CheckPathPath)' does not exist. Exiting." -ForegroundColor Red
                            $script:ExitCode = 24
                            $script:ExitCodeDescription = "Path '$($CheckPathPath)' does not exist.";
                            exit
                        }
                    } else {
                        Write-Host '    SharePoint via Graph: No DriveID. Wrong path or missing permission in SharePoint?' -ForegroundColor Yellow
                    }
                } else {
                    Write-Host '    SharePoint via Graph: No SiteID. Wrong path or missing permission in Entra ID app?' -ForegroundColor Yellow
                }
            }

            try { WatchCatchableExitSignal } catch { }

            if ((Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)) {
                Write-Verbose "      '$($CheckPathPath)' is accessible, nothing more to do."
            } else {
                # SharePoint Online without Graph client ID or SharePoint on-prem
                # Or normal file path that does not exist

                if ($IsWindows) {
                    # Windows. Use old way with "net use", Internet-Explorer-Cookie.

                    if (($CheckPathPath.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($CheckPathPath -ilike '*@SSL\*')) {
                        Write-Host '    SharePoint via WebDAV, may be slow and path length problems may occur (fully qualified file names must be less than 260 characters).' -ForegroundColor Yellow
                        $CheckPathPath = $CheckPathPath -ireplace '@SSL\\', '\'
                        $CheckPathPath = ([uri]::UnescapeDataString($CheckPathPath) -ireplace ('https://', '\\'))
                        $CheckPathPath = ([System.URI]$CheckPathPath).AbsoluteURI -ireplace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\${2}' -ireplace '/', '\'
                        $CheckPathPath = [uri]::UnescapeDataString($CheckPathPath)
                    } else {
                        try {
                            $CheckPathPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CheckPathPath)
                            $CheckPathPath = ([System.URI]$CheckPathPath).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -ireplace '/', '\'
                            $CheckPathPath = [uri]::UnescapeDataString($CheckPathPath)
                        } catch {
                            if ($CheckPathSilent -eq $false) {
                                Write-Host "Problem connecting or reading '$CheckPathPath'. Exit." -ForegroundColor Red
                                $script:ExitCode = 25
                                $script:ExitCodeDescription = "Problem connecting or reading '$CheckPathPath'.";
                                exit
                            }
                        }
                    }

                    if (-not (Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)) {
                        # Reconnect already connected network drives at the OS level
                        # New-PSDrive is not enough for this
                        foreach ($NetworkConnection in @(Get-CimInstance Win32_NetworkConnection)) {
                            try { WatchCatchableExitSignal } catch { }
                            & net use $NetworkConnection.LocalName $NetworkConnection.RemoteName 2>&1 | Out-Null
                        }

                        if (-not (Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue)) {
                            try { WatchCatchableExitSignal } catch { }

                            # Connect network drives
                            $([System.Environment]::NewLine) | & net use "$CheckPathPath" 2>&1 | Out-Null

                            try { WatchCatchableExitSignal } catch { }

                            try {
                                (Test-Path -LiteralPath $CheckPathPath -ErrorAction Stop) | Out-Null
                            } catch {
                                if ($_.CategoryInfo.Category -eq 'PermissionDenied') {
                                    try { WatchCatchableExitSignal } catch { }
                                    & net use "$CheckPathPath" 2>&1
                                }
                            }

                            try { WatchCatchableExitSignal } catch { }

                            & net use "$CheckPathPath" /d 2>&1 | Out-Null
                        }

                        try { WatchCatchableExitSignal } catch { }

                        if (($CheckPathPath -ilike '*@SSL\*') -and (-not (Test-Path -LiteralPath $CheckPathPath -ErrorAction SilentlyContinue))) {
                            if ((Get-Service -ServiceName 'WebClient' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).Status -ine 'Running') {
                                if (-not $CheckPathSilent) {
                                    Write-Host
                                    Write-Host 'WebClient service not running.' -ForegroundColor Red
                                }
                            } else {
                                try {
                                    if (-not [string]::IsNullOrWhitespace($GraphHtmlMessageboxText)) {
                                        if ($IsWindows -and (-not (Test-Path -LiteralPath env:SSH_CLIENT))) {
                                            Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

                                            $window = New-Object System.Windows.Window -Property @{
                                                Width                 = 1
                                                Height                = 1
                                                WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
                                                ShowActivated         = $false
                                                Topmost               = $true
                                            }

                                            $window.Show()
                                            $window.Hide()

                                            $MessageBoxResult = [System.Windows.MessageBox]::Show($window, "$($GraphHtmlMessageboxText)", $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }), [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Information, [System.Windows.MessageBoxResult]::None)

                                            if ($MessageBoxResult -ieq 'Cancel') {
                                                $window.Close()

                                                Write-Host
                                                Write-Host 'Authentication cancelled by user. Exiting.' -ForegroundColor Red

                                                $script:ExitCode = 26
                                                $script:ExitCodeDescription = 'Authentication cancelled by user.';
                                                exit
                                            }

                                            $window.Close()
                                        }
                                    }

                                    # Add site to trusted sites in internet options
                                    New-Item ('HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\' + (New-Object System.Uri -ArgumentList ($CheckPathPath -ireplace ('@SSL', ''))).Host) -Force | New-ItemProperty -Name * -Value 1 -Type DWORD -Force | Out-Null

                                    # Open site in new IE process
                                    $oIE = New-Object -com InternetExplorer.Application
                                    $oIE.Navigate('https://' + ((($CheckPathPath -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')) + '?web=1')
                                    $oIE.Visible = $true

                                    # Wait until an IE tab with the corresponding URL is open
                                    $app = New-Object -com shell.application

                                    $i = 0

                                    $compareurl = ('*' + ([uri]::UnescapeDataString([uri]::UnescapeDataString((($CheckPathPath -ireplace ('@SSL', '')).replace('\\', '')).replace('\', '/')))) + '*') -split '\/' -join '*'

                                    while ($i -lt 1) {
                                        $i += @($app.windows() | Where-Object {
                                                ([uri]::UnescapeDataString([uri]::UnescapeDataString($_.LocationURL)) -ilike $compareurl)
                                            }).count

                                        Start-Sleep -Seconds 1

                                        try { WatchCatchableExitSignal } catch { }
                                    }

                                    # Wait until the corresponding URL is fully loaded, then close the tab
                                    @($app.windows() | Where-Object {
                                            ([uri]::UnescapeDataString([uri]::UnescapeDataString($_.LocationURL)) -ilike $compareurl)
                                        }) | ForEach-Object {

                                        while ($_.Busy) {
                                            Start-Sleep -Milliseconds 100

                                            try { WatchCatchableExitSignal } catch { }
                                        }

                                        $_.Quit()
                                    }

                                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
                                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($oIE) | Out-Null

                                    Remove-Variable -Name 'app'
                                    Remove-Variable -Name 'oIE'
                                } catch {
                                    $_
                                }
                            }
                        }
                    }
                } else {
                    if (($CheckPathPath.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($CheckPathPath -ilike '*@SSL\*')) {
                        Write-Host '    SharePoint via WebDAV is only supported on Windows platforms.' -ForegroundColor Yellow
                    }
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        if ((Test-Path -LiteralPath $CheckPathPath -PathType $ExpectedPathType) -eq $false) {
            if ($CheckPathSilent -eq $false) {
                Write-Host "Problem connecting or reading $($ExpectedPathType) '$($CheckPathPath)'. Exit." -ForegroundColor Red
                $script:ExitCode = 27
                $script:ExitCodeDescription = "Problem connecting or reading $($ExpectedPathType) '$($CheckPathPath)'.";
                exit
            } else {
                return $false
            }
        } else {
            if ($CheckPathSilent -eq $false) {
                # Write-Host
            } else {
                return $true
            }
        }
    } else {
        Write-Verbose "      Try to create '$($CheckPathPath)'."

        if ($CheckPathPath.StartsWith('https://', 'CurrentCultureIgnoreCase')) {
            $CheckPathPath = ((([uri]::UnescapeDataString($CheckPathPath) -ireplace ('https://', '\\')) -ireplace ('(.*?)/(.*)', '${1}@SSL\${2}')) -ireplace ('/', '\'))
        } else {
            # '@SSL' seems to be case sensitive, so we make sure that the first occurrence is in uppercase letters
            $CheckPathPath = ([regex]"(?i)$([regex]::escape('@ssl\'))").replace($CheckPathPath, '@SSL\', 1)
        }

        $CheckPathPathTarget = $CheckPathPath

        for (
            $i = 1
            $i -lt @($CheckPathPathTarget -split [regex]::escape([IO.Path]::DirectorySeparatorChar)).count
            $i++
        ) {
            try { WatchCatchableExitSignal } catch { }

            $CheckPathPathTemp = @($CheckPathPathTarget -split [regex]::escape([IO.Path]::DirectorySeparatorChar))[0..$i] -join [IO.Path]::DirectorySeparatorChar

            if ((. $CheckPathScriptblock ([ref]$CheckPathPathTemp) -CheckPathSilent) -eq $true) {
                if (-not (Test-Path -LiteralPath $CheckPathPathTemp -PathType Container -ErrorAction SilentlyContinue)) {
                    Write-Host "'$CheckPathPathTemp' is a file, '$CheckPathPathTarget' is not valid. Exit." -ForegroundColor Red
                    $script:ExitCode = 28
                    $script:ExitCodeDescription = "'$CheckPathPathTemp' is a file, '$CheckPathPathTarget' is not valid.";
                    exit
                }

                if ($CheckPathPathTemp -eq $CheckPathPathTarget) {
                    break
                } else {
                    try { WatchCatchableExitSignal } catch { }

                    Write-Verbose "      Try to create '$($CheckPathPathTarget)'."

                    New-Item -ItemType Directory -Path $CheckPathPathTarget -ErrorAction SilentlyContinue | Out-Null

                    if (Test-Path -LiteralPath $CheckPathPathTarget -PathType Container) {
                        break
                    }
                }
            }
        }

        if ((. $CheckPathScriptblock ([ref]$CheckPathPathTarget) -CheckPathSilent) -ne $true) {
            Write-Host "Problem connecting or reading '$CheckPathPathTarget'. Exit." -ForegroundColor Red
            $script:ExitCode = 29
            $script:ExitCodeDescription = "Problem connecting or reading '$CheckPathPathTarget'.";
            exit
        } else {
            # Write-Host
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function ConnectEWS([string]$MailAddress = $MailAddresses[0], [string]$Indent = '') {
    Write-Host "$($Indent)Connect to Outlook Web"

    GraphSwitchContext -TenantID $MailAddress

    if (-not $script:WebServicesDllPath) {
        Write-Host "$Indent  Set up environment for connection to Outlook Web"

        try { WatchCatchableExitSignal } catch { }

        $script:WebServicesDllPath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))

        try {
            Copy-Item -LiteralPath ((Join-Path -Path '.' -ChildPath 'bin\EWS\netstandard2.0\Microsoft.Exchange.WebServices.Data.dll')) -Destination $script:WebServicesDllPath -Force
            if (-not $IsLinux) {
                Unblock-File -LiteralPath $script:WebServicesDllPath
            }
        } catch {
            Write-Verbose "$Indent    $($_)"
        }
    }

    $local:exchServiceAvailable = $false

    if ($script:exchService) {
        try {
            if (
                $($script:exchService.SetOutlookSignaturesMailaddress -ieq $MailAddress) -and
                $(([Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:exchService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)).DisplayName)
            ) {
                Write-Host "$($Indent)  Existing connection matches required parameters and is working"

                $local:exchServiceAvailable = $true
            } else {
                Write-Host "$($Indent)  Existing connecting does not match required parameters or does not work"
            }
        } catch {
            Write-Host "$($Indent)  Existing connecting does not match required parameters or does not work"
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ($local:exchServiceAvailable -eq $false) {
        Write-Host "$($Indent)  Creating new connection"

        $script:exchService = $null

        try {
            Import-Module -Name $script:WebServicesDllPath -Force -ErrorAction Stop

            try { WatchCatchableExitSignal } catch { }

            $script:exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

            $script:exchService.Timeout = 25000

            try { WatchCatchableExitSignal } catch { }

            $tempEwsRedirectUrl = $null

            function ExchServiceEwsTraceHandler() {
                $sourceCode = @'
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;

public class ExchServiceEwsTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener
{
    public void Trace(System.String traceType, System.String traceMessage)
    {
        string tempEwsRedirectUrl = string.Empty;

        Match match = Regex.Match(traceMessage, "Redirection URL found: '(.*?)'");

        if (match.Success)
        {
            tempEwsRedirectUrl = match.Groups[1].Value;
        }

        StringBuilder sb = new StringBuilder();
        // sb.AppendLine("Write-Verbose \"$($Indent)      traceType: $($('" + System.Management.Automation.Language.CodeGeneration.EscapeSingleQuotedStringContent(traceType) + "'))\"");
        // sb.AppendLine("Write-Verbose \"$($Indent)      traceMessage: $($('" + System.Management.Automation.Language.CodeGeneration.EscapeSingleQuotedStringContent(traceMessage) + "'))\"");
        sb.AppendLine("$tempEwsRedirectUrl = $($('" + System.Management.Automation.Language.CodeGeneration.EscapeSingleQuotedStringContent(tempEwsRedirectUrl) + "'))");

        var defRunspace = System.Management.Automation.Runspaces.Runspace.DefaultRunspace;
        var pipeline = defRunspace.CreateNestedPipeline();
        pipeline.Commands.AddScript(sb.ToString());
        pipeline.Invoke();
    }
}
'@

                Add-Type -TypeDefinition $sourceCode -Language CSharp -ReferencedAssemblies $script:WebServicesDllPath, System.Management.Automation, System.Text.RegularExpressions
                $ExchServiceEwsTraceListener = New-Object ExchServiceEwsTraceListener
                return $ExchServiceEwsTraceListener
            }

            $script:exchService.TraceEnabled = $true
            $script:exchService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::AutodiscoverConfiguration, [Microsoft.Exchange.WebServices.Data.TraceFlags]::AutodiscoverRequest, [Microsoft.Exchange.WebServices.Data.TraceFlags]::AutodiscoverResponse
            $script:exchService.TraceListener = ExchServiceEwsTraceHandler

            try { WatchCatchableExitSignal } catch { }

            try {
                Write-Host "$($Indent)    Try Autodiscover with Integrated Windows Authentication"

                $script:exchService.UseDefaultCredentials = $true
                $script:exchService.ImpersonatedUserId = $null
                $script:exchService.AutodiscoverUrl($MailAddress, { $true }) | Out-Null

                Write-Host "$($Indent)      Success"
            } catch {
                Write-Host "$($Indent)      Failed. See verbose output for details."
                Write-Verbose "$($Indent)        $($_)"
                Write-Verbose "$($Indent)      This is OK when:"
                Write-Verbose "$($Indent)        - Not connected to internal network"
                Write-Verbose "$($Indent)        - Connected to internal network with no Exchange on prem."
                Write-Verbose "$($Indent)        - Connected to internal network with Exchange on prem, but your mailbox is in Exchange Online."
                Write-Verbose "$($Indent)        - Connected to internal network but not logged-on with Active Directory credentials."
                Write-Verbose "$($Indent)      Else, you should check your internal and/or external Autodiscover configuration:"
                Write-Verbose "$($Indent)        - External: https://testconnectivity.microsoft.com"
                Write-Verbose "$($Indent)        - Internal: https://learn.microsoft.com/en-us/exchange/architecture/client-access/autodiscover"
                Write-Verbose "$($Indent)        - Check your loadbalancer configuration."

                if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                    Write-Verbose "$($Indent)      Anyhow:"
                    Write-Verbose "$($Indent)        - Redirect URL '$($tempEwsRedirectUrl)' was returned."
                    Write-Verbose "$($Indent)        - No need to try Autodiscver with OAuth, skipping to OAuth with fixed URL."
                } else {
                    $tempEwsRedirectUrl = $null
                }

                if (
                    $($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile -and !$script:GraphToken.AppAccessTokenExo) -or
                    -not $script:GraphToken.AccessTokenExo
                ) {
                    throw "Integrated Windows Authentication failed, and there is no EXO OAuth access token available. Did you forget '-GraphOnly true' or are you missing AD attributes?"
                }

                try { WatchCatchableExitSignal } catch { }

                try {
                    Write-Host "$($Indent)    Try Autodiscover with OAuth"

                    if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                        throw 'Autodiscover with IWA failed before, but returned a redirect URL. We will use this fixed URL without Autodiscover.'
                    } else {
                        $tempEwsRedirectUrl = $null
                    }

                    $script:exchService.UseDefaultCredentials = $false

                    if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                        $script:exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailAddress)
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($script:GraphToken.AppAccessTokenExo)
                    } else {
                        $script:exchService.ImpersonatedUserId = $null
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($script:GraphToken.AccessTokenExo)
                    }

                    $script:exchService.AutodiscoverUrl($MailAddress, { $true }) | Out-Null

                    Write-Host "$($Indent)      Success"
                } catch {
                    if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                        Write-Host "$($Indent)      Skipping Autodiscover with OAuth because"
                        Write-Host "$($Indent)        $($_)"
                    } else {
                        Write-Host "$($Indent)      Failed.  See verbose output for details."
                        Write-Verbose "$($Indent)        $($_)"
                        Write-Verbose "$($Indent)      This is OK when"
                        Write-Verbose "$($indent)        - Connected to internal network with Exchange on prem without Hybrid Modern Authentication"
                        Write-Verbose "$($Indent)      Else, you should check your internal and/or external Autodiscover configuration:"
                        Write-Verbose "$($Indent)        - External: https://testconnectivity.microsoft.com"
                        Write-Verbose "$($Indent)        - Internal: https://learn.microsoft.com/en-us/exchange/architecture/client-access/autodiscover"
                        Write-Verbose "$($Indent)        - Check your loadbalancer configuration."
                    }

                    try { WatchCatchableExitSignal } catch { }

                    Write-Host "$($Indent)    Try OAuth with fixed URL"

                    $script:exchService.UseDefaultCredentials = $false

                    if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) {
                        $script:exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailAddress)
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($script:GraphToken.AppAccessTokenExo)
                    } else {
                        $script:exchService.ImpersonatedUserId = $null
                        $script:exchService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $($script:GraphToken.AccessTokenExo)
                    }

                    if ([System.Uri]::IsWellFormedUriString($tempEwsRedirectUrl, [System.UriKind]::Absolute)) {
                        $script:exchService.Url = "$(([uri]$tempEwsRedirectUrl).GetLeftPart([UriPartial]::Authority))/EWS/Exchange.asmx"
                    } else {
                        $script:exchService.Url = "$($script:CloudEnvironmentExchangeOnlineEndpoint)/EWS/Exchange.asmx"
                    }

                    Write-Host "$($Indent)      Success"
                    Write-Host "$($Indent)      Fixed URL: '$($script:exchService.Url)'"
                }
            }

            if (([Microsoft.Exchange.WebServices.Data.Folder]::Bind($script:exchService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)).DisplayName) {
                Add-Member -InputObject $script:exchService -MemberType NoteProperty -Name 'SetOutlookSignaturesMailaddress' -Value $MailAddress -Force
            } else {
                throw 'Could not connect to Outlook Web, although the EWS DLL threw no error.'
            }
        } catch {
            Write-Host "$($Indent)    Error connecting to Outlook Web: $($_)" -ForegroundColor Red
            Write-Host "$($Indent)    Check verbose output for details and solution hints." -ForegroundColor Red

            $script:exchService = $null
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function GraphGetToken {
    param(
        [switch]$EXO,
        [string]$indent = ''
    )

    try { WatchCatchableExitSignal } catch { }

    if (-not $EXO) {
        Write-Host "$($indent)Graph authentication"

        if ($GraphClientID -ieq 'beea8249-8c98-4c76-92f6-ce3c468a61e6') {
            Write-Host "$($indent)  You use the Entra ID app provided by the developers. It is recommended to create und use your own Entra ID app." -ForegroundColor Yellow
            Write-Host "$($indent)    Find a description on how to do this in the file '`.\config\default graph config.ps1`'." -ForegroundColor Yellow
        }

        $script:GraphUser = $null
    }

    if ($SimulateAndDeployGraphCredentialFile) {
        Write-Host "$($indent)  Via SimulateAndDeployGraphCredentialFile '$SimulateAndDeployGraphCredentialFile'"

        try {
            try {
                $auth = Import-Clixml -LiteralPath $SimulateAndDeployGraphCredentialFile
            } catch {
                Start-Sleep -Seconds 2
                $auth = Import-Clixml -LiteralPath $SimulateAndDeployGraphCredentialFile
            }

            if ($auth.AccessToken) {
                $script:AuthorizationToken = $auth.AccessToken
                $script:ExoAuthorizationToken = $auth.AccessTokenExo

                $script:AuthorizationHeader = @{
                    Authorization = $auth.AuthHeader
                }
                $script:ExoAuthorizationHeader = @{
                    Authorization = $auth.AuthHeaderExo
                }

                $script:AppAuthorizationToken = $auth.AppAccessToken
                $script:AppExoAuthorizationToken = $auth.AppAccessTokenExo

                $script:AppAuthorizationHeader = @{
                    Authorization = $auth.AppAuthHeader
                }
                $script:AppExoAuthorizationHeader = @{
                    Authorization = $auth.AppAuthHeaderExo
                }
            } else {
                $script:GraphUser = $SimulateUser
                $script:GraphDomainToTenantIDCache = $auth.GraphDomainToTenantIDCache
                $script:GraphDomainToCloudInstanceCache = $auth.GraphDomainToCloudInstanceCache
                $script:GraphTokenDictionary = $auth.GraphTokenDictionary

                GraphSwitchContext $null
            }

            return @{
                error             = $false
                AccessToken       = $script:AuthorizationToken
                AuthHeader        = $script:AuthorizationHeader
                AccessTokenExo    = $script:ExoAuthorizationToken
                AuthHeaderExo     = $script:ExoAuthorizationHeader
                AppAccessToken    = $script:AppAuthorizationToken
                AppAuthHeader     = $script:AppAuthorizationHeader
                AppAccessTokenExo = $script:AppExoAuthorizationToken
                AppAuthHeaderExo  = $script:AppExoAuthorizationHeader
            }
        } catch {
            return @{
                error             = ($error[0] | Out-String)
                AccessToken       = $null
                AuthHeader        = $null
                AccessTokenExo    = $null
                AuthHeaderExo     = $null
                AppAccessToken    = $null
                AppAuthHeader     = $null
                AppAccessTokenExo = $null
                AppAuthHeaderExo  = $null
            }
        }
    } else {
        if (-not  $script:MsalModulePath) {
            Write-Host "$($indent)  Load MSAL.PS"

            $script:MsalModulePath = (Join-Path -Path $script:tempDir -ChildPath 'MSAL.PS')

            # Copy each item to the destination
            foreach ($item in @(Get-ChildItem -LiteralPath ((Join-Path -Path '.' -ChildPath 'bin\MSAL.PS')) -Recurse)) {
                if ($item.BaseName -like '*msalruntime*') {
                    if ($IsWindows) {
                        if ($item.Name -inotlike 'msalruntime*.dll') {
                            continue
                        }
                    } elseif ($IsLinux) {
                        if ($item.Name -inotlike 'libmsalruntime.so') {
                            continue
                        }
                    } elseif ($IsMacOS) {
                        if ($item.Name -inotlike 'msalruntime*.dylib') {
                            continue
                        }
                    } else {
                        continue
                    }
                }

                $destinationPath = $item.FullName -replace [regex]::escape((Resolve-Path -LiteralPath (Join-Path -Path '.' -ChildPath 'bin\MSAL.PS')).ProviderPath), $script:MsalModulePath

                if ($item.PSIsContainer) {
                    # Create the directory if it doesn't exist
                    if (-not (Test-Path -LiteralPath $destinationPath)) {
                        $null = New-Item -ItemType Directory -Path $destinationPath -Force
                    }
                } else {
                    # Copy the file
                    Copy-Item -LiteralPath $item.FullName -Destination $destinationPath
                }
            }

            if (-not $IsLinux) {
                Get-ChildItem -LiteralPath $script:MsalModulePath -Recurse | Unblock-File
            }

            try { WatchCatchableExitSignal } catch { }

            try {
                Import-Module $script:MsalModulePath -Force -ErrorAction Stop
            } catch {
                Write-Host $error[0]
                Write-Host "$($indent)    Problem importing MSAL.PS module. Exit." -ForegroundColor Red
                $script:ExitCode = 30
                $script:ExitCodeDescription = 'Problem importing MSAL.PS module.';
                exit
            }
        }

        try { WatchCatchableExitSignal } catch { }

        # On Linux/macOS, unlock keyring/keychain if required
        if (-not [string]::IsNullOrWhitespace($GraphUnlockKeyringKeychainMessageboxText)) {
            if ($IsLinux) {
                $keyringPath = (dbus-send --session --dest=org.freedesktop.secrets --type=method_call --print-reply /org/freedesktop/secrets org.freedesktop.Secret.Service.ReadAlias string:'default' | grep -oP '(?<=object path \")/[^"]+')

                if ($((gdbus call -e -d org.freedesktop.secrets -o $keyringPath -m org.freedesktop.DBus.Properties.Get org.freedesktop.Secret.Collection Locked *>&1) -ine '(<false>,)')) {
                    if ($(Get-Command -Name 'kdialog' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                        $null = kdialog `
                            --title $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                            --msgbox "$($GraphUnlockKeyringKeychainMessageboxText)"
                    } elseif ($(Get-Command -Name 'zenity' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                        $null = zenity `
                            --info `
                            --title=$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                            --text="$($GraphUnlockKeyringKeychainMessageboxText)"
                    } else {
                        Write-Host "$($indent)  Neither kdialog nor zenity found, so no message box could be shown: $($GraphUnlockKeyringKeychainMessageboxText)"
                    }
                }
            } elseif ($IsMacOS) {
                security unlock-keychain -p 'Set-OutlookSignatures dummy password' *>$null

                if ($LastExitCode -ne 0) {
                    Write-Host $("display alert ""$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })"" message ""$($GraphUnlockKeyringKeychainMessageboxText)""  buttons { ""OK"" } default button 1" | osascript *>$1; '')
                }
            }
        }

        try { WatchCatchableExitSignal } catch { }

        # Graph authentication
        Write-Host "$($indent)  Authentication against $(if(-not $EXO) { 'Graph' } else { 'Exchange Online' })"

        try {
            Write-Host "$($indent)    Silent via Integrated Windows Authentication without login hint"

            if ($EXO) {
                throw 'Ignoring because login hint is available.'
            }

            $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

            $auth = $script:msalClientApp | Get-MsalToken -IntegratedWindowsAuth -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 1)

            Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
        } catch {
            Write-Host "$($indent)      Failed: $($error[0])"

            try { WatchCatchableExitSignal } catch { }

            try {
                Write-Host "$($indent)    Silent via Integrated Windows Authentication with login hint"
                # Required, because IWA without login hint may fail when account enumeration is blocked at OS level

                $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                if (-not $EXO) {
                    $script:GraphUser = ($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username
                }

                Write-Host "$($indent)      Login hint: '$($script:GraphUser)'"

                if (-not ([string]::IsNullOrWhiteSpace($script:GraphUser))) {
                    $auth = $script:msalClientApp | Get-MsalToken -IntegratedWindowsAuth -LoginHint $script:GraphUser -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 1)
                } else {
                    throw 'No login hint found before'
                }

                Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
            } catch {
                Write-Host "$($indent)      Failed: $($error[0])"

                try { WatchCatchableExitSignal } catch { }

                try {
                    Write-Host "$($indent)    Silent via Authentication Broker without login hint"

                    if ($EXO) {
                        throw 'Ignoring because login hint is available.'
                    }

                    if ($IsLinux) {
                        $LinuxAuthBrokerMissingDependencies = @(ldd $(Join-Path -Path $script:MsalModulePath -ChildPath 'netstandard2.0/libmsalruntime.so') | grep 'not found')

                        if ($LinuxAuthBrokerMissingDependencies.Count -gt 0) {
                            throw 'Missing dependencies for authentication broker: ' + $(
                                @(
                                    $LinuxAuthBrokerMissingDependencies | ForEach-Object {
                                        $(($_ -ireplace '=> not found', '').Trim())
                                    }
                                ) -join ', '
                            )
                        }
                    }

                    $script:msalClientApp = New-MsalClientApplication -AuthenticationBroker -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                    $auth = $script:msalClientApp | Get-MsalToken -Silent -AuthenticationBroker -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)

                    Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
                } catch {
                    Write-Host "$($indent)      Failed: $($error[0])"

                    try { WatchCatchableExitSignal } catch { }

                    try {
                        Write-Host "$($indent)    Silent via Authentication Broker with login hint"

                        if ($IsLinux) {
                            $LinuxAuthBrokerMissingDependencies = @(ldd $(Join-Path -Path $script:MsalModulePath -ChildPath 'netstandard2.0/libmsalruntime.so') | grep 'not found')

                            if ($LinuxAuthBrokerMissingDependencies.Count -gt 0) {
                                throw 'Missing dependencies for authentication broker: ' + $(
                                    @(
                                        $LinuxAuthBrokerMissingDependencies | ForEach-Object {
                                            $(($_ -ireplace '=> not found', '').Trim())
                                        }
                                    ) -join ', '
                                )
                            }
                        }

                        $script:msalClientApp = New-MsalClientApplication -AuthenticationBroker -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                        if (-not $EXO) {
                            $script:GraphUser = ($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username
                        }

                        Write-Host "$($indent)      Login hint: '$($script:GraphUser)'"

                        if (-not ([string]::IsNullOrWhiteSpace($script:GraphUser))) {
                            $auth = $script:msalClientApp | Get-MsalToken -Silent -AuthenticationBroker -LoginHint $script:GraphUser -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)
                        } else {
                            throw 'No login hint found before'
                        }

                        Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
                    } catch {
                        Write-Host "$($indent)      Failed: $($error[0])"

                        try {
                            Write-Host "$($indent)    Silent via refresh token, with login hint"

                            $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId -RedirectUri 'http://localhost' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                            if (-not $EXO) {
                                $script:GraphUser = ($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username
                            }

                            Write-Host "$($indent)      Login hint: '$($script:GraphUser)'"

                            if (-not ([string]::IsNullOrWhiteSpace($script:GraphUser))) {
                                $auth = $script:msalClientApp | Get-MsalToken -Silent -LoginHint $script:GraphUser -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -ForceRefresh -Timeout (New-TimeSpan -Minutes 1)
                            } else {
                                throw 'No login hint found before'
                            }

                            Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
                        } catch {
                            Write-Host "$($indent)      Failed: $($error[0])"

                            try { WatchCatchableExitSignal } catch { }

                            # Interactive authentication methods
                            Write-Host "$($indent)    All silent authentication methods failed, switching to interactive authentication methods."

                            if ((-not $EXO) -and (-not [string]::IsNullOrWhitespace($GraphHtmlMessageboxText))) {
                                if ($IsWindows -and (-not (Test-Path -LiteralPath env:SSH_CLIENT))) {
                                    Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

                                    $window = New-Object System.Windows.Window -Property @{
                                        Width                 = 1
                                        Height                = 1
                                        WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
                                        ShowActivated         = $false
                                        Topmost               = $true
                                    }

                                    $window.Show()
                                    $window.Hide()

                                    $MessageBoxResult = [System.Windows.MessageBox]::Show($window, "$($GraphHtmlMessageboxText)", $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }), [System.Windows.MessageBoxButton]::OKCancel, [System.Windows.MessageBoxImage]::Information, [System.Windows.MessageBoxResult]::None)

                                    $window.Close()

                                    if ($MessageBoxResult -ieq 'Cancel') {
                                        return @{
                                            error             = 'Authentication cancelled by user. Exiting.'
                                            AccessToken       = $null
                                            authHeader        = $null
                                            AccessTokenExo    = $null
                                            authHeaderExo     = $null
                                            AppAccessToken    = $null
                                            AppAuthHeader     = $null
                                            AppAccessTokenExo = $null
                                            AppAuthHeaderExo  = $null
                                        }
                                    }
                                } elseif ($IsLinux -and ((Test-Path -LiteralPath env:DISPLAY))) {
                                    if ($(Get-Command -Name 'kdialog' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                                        $null = kdialog `
                                            --title $(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                                            --msgbox "$($GraphHtmlMessageboxText)"
                                    } elseif ($(Get-Command -Name 'zenity' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                                        $null = zenity `
                                            --info `
                                            --title=$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' }) `
                                            --text="$($GraphHtmlMessageboxText)"
                                    } else {
                                        Write-Host "$($indent)    Neither kdialog nor zenity found, so no message box could be shown: $($GraphHtmlMessageboxText)"
                                    }
                                } elseif ($IsMacOS -and ((Test-Path -LiteralPath env:DISPLAY))) {
                                    Write-Host $("display alert ""$(if ($BenefactorCircleLicenseFile) { 'Set-OutlookSignatures Benefactor Circle' } else { 'Set-OutlookSignatures' })"" message ""$($GraphHtmlMessageboxText)""  buttons { ""OK"" } default button 1" | osascript *>&1; '')
                                }

                                try { WatchCatchableExitSignal } catch { }
                            }

                            $MsalInteractiveParams = @{}

                            if (-not [string]::IsNullOrWhiteSpace($GraphBrowserRedirectSuccess)) {
                                $MsalInteractiveParams.BrowserRedirectSuccess = $GraphBrowserRedirectSuccess
                            }

                            if (-not [string]::IsNullOrWhiteSpace($GraphBrowserRedirectError)) {
                                $MsalInteractiveParams.BrowserRedirectError = $GraphBrowserRedirectError
                            }

                            if (-not [string]::IsNullOrWhiteSpace($GraphHtmlMessageSuccess)) {
                                $MsalInteractiveParams.HtmlMessageSuccess = $GraphHtmlMessageSuccess
                            }

                            if (-not [string]::IsNullOrWhiteSpace($GraphHtmlMessageError)) {
                                $MsalInteractiveParams.HtmlMessageError = $GraphHtmlMessageError
                            }

                            try { WatchCatchableExitSignal } catch { }

                            try {
                                Write-Host "$($indent)    Interactive via Authentication Broker"

                                if ($IsLinux) {
                                    $LinuxAuthBrokerMissingDependencies = @(ldd $(Join-Path -Path $script:MsalModulePath -ChildPath 'netstandard2.0/libmsalruntime.so') | grep 'not found')

                                    if ($LinuxAuthBrokerMissingDependencies.Count -gt 0) {
                                        throw 'Missing dependencies for authentication broker: ' + $(
                                            @(
                                                $LinuxAuthBrokerMissingDependencies | ForEach-Object {
                                                    $(($_ -ireplace '=> not found', '').Trim())
                                                }
                                            ) -join ', '
                                        )
                                    }
                                }

                                $script:msalClientApp = New-MsalClientApplication -AuthenticationBroker -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                                if (-not $EXO) {
                                    $script:GraphUser = ($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username
                                }

                                Write-Host "$($indent)      Login hint: '$($script:GraphUser)'"

                                Write-Host "$($indent)      Opening authentication broker window and waiting for you to authenticate. Stopping script execution after five minutes."
                                $auth = $script:msalClientApp | Get-MsalToken -Interactive -AuthenticationBroker -LoginHint $(if ($script:GraphUser) { $script:GraphUser } else { '' }) -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 5) -Prompt 'NoPrompt' -UseEmbeddedWebView:$false @MsalInteractiveParams

                                Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
                            } catch {
                                Write-Host "$($indent)      Failed: $($error[0])"

                                try {
                                    Write-Host "$($indent)    Interactive via browser"

                                    $script:msalClientApp = New-MsalClientApplication -ClientId $GraphClientID -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -TenantId $script:GraphTenantId -RedirectUri 'http://localhost' | Enable-MsalTokenCacheOnDisk -PassThru -WarningAction SilentlyContinue

                                    if (-not $EXO) {
                                        $script:GraphUser = ($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username
                                    }

                                    Write-Host "$($indent)      Login hint: '$($script:GraphUser)'"

                                    Write-Host "$($indent)      Opening new browser window and waiting for you to authenticate. Stopping script execution after five minutes."
                                    $auth = $script:msalClientApp | Get-MsalToken -Interactive -LoginHint $(if ($script:GraphUser) { $script:GraphUser } else { '' }) -AzureCloudInstance $script:CloudEnvironmentEnvironmentName -Scopes $(if (-not $EXO) { "$($script:CloudEnvironmentGraphApiEndpoint)/.default" } else { "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default" }) -Timeout (New-TimeSpan -Minutes 5) -Prompt 'NoPrompt' -UseEmbeddedWebView:$false @MsalInteractiveParams

                                    Write-Host "$($indent)      Success: '$(($script:msalClientApp | Get-MsalAccount | Select-Object -First 1).username)'"
                                } catch {
                                    Write-Host "$($indent)      Failed: $($error[0])"
                                    Write-Host '$($indent)    No authentication possible'

                                    $auth = $null

                                    return @{
                                        error             = (($error[0] | Out-String) + @"
No authentication possible.
1. Did you follow the Quick Start Guide and configure the Entra ID app correctly?
   https://set-outlooksignatures.com/quickstart
2. Run Set-OutlookSignatures with the "-Verbose" parameter and check for authentication messages
3. If the "Interactive" message is displayed:
   - When using an Authentication Broker (which is preferred on supported platforms):
     - Does the account picker window show up?
     - Check if authentication happens within five minutes
     - Check if your firewall or anti-malware software blocks Set-OutlookSignatures from creating a temporary listener port for localhost.
     - Check if the correct user account is selected/entered and if the authentication is successful
   - When not using an Authentication Broker (on a system without support for it, or when broker auth failed):
     - Does a browser (the system default browser, if configured) open and ask for authentication?
      - Yes:
       - Check if authentication happens within five minutes
       - Ensure that your browser does not block access to 'http://localhost', errors such as 'connection refused' point to this problem. ('https://localhost' is currently not technically feasible, see 'https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/using-web-browsers' and 'https://learn.microsoft.com/en-us/entra/msal/dotnet/acquiring-tokens/using-web-browsers' for details)
         This is typically due to enforced redirection to HTTPS being applied to localhost. If not configured via policies: edge://net-internals/#hsts or chrome://net-internals/#hsts, delete domain security policies for localhost.
       - Check if your firewall or anti-malware software blocks Set-OutlookSignatures from creating a temporary listener port for localhost.
       - Check if the correct user account is selected/entered and if the authentication is successful
     - No:
       - Check if a default browser is set and if the PowerShell command 'start https://set-outlooksignatures.com' opens it
       - Make sure that Set-OutlookSignatures is executed in the security context of the currently logged-in user
       - Run Set-OutlookSignatures in a new PowerShell session
       - Check your anti-malware configuration (errors such as 'error sending the request' or 'connection refused' point at a problem there)
       - Make sure that the current PowerShell session allows TLS 1.2+ (see https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues/85 for details)
4. Delete the Graph token cache: $($script:msalClientApp.cacheInfo).
"@)
                                        AccessToken       = $null
                                        AuthHeader        = $null
                                        AccessTokenExo    = $null
                                        AuthHeaderExo     = $null
                                        AppAccessToken    = $null
                                        AppAuthHeader     = $null
                                        AppAccessTokenExo = $null
                                        AppAuthHeaderExo  = $null
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if ($auth) {
            try {
                $script:GraphUser = $auth.account.username

                if (-not $EXO) {
                    $script:AuthorizationHeader = @{
                        Authorization = $auth.CreateAuthorizationHeader()
                    }

                    $script:AuthorizationToken = $auth.AccessToken
                } else {
                    $script:ExoAuthorizationHeader = @{
                        Authorization = $auth.CreateAuthorizationHeader()
                    }

                    $script:ExoAuthorizationToken = $auth.AccessToken
                }

                if (-not $EXO) {
                    $authExo = GraphGetToken -EXO -indent $indent

                    if ($authExo -and ($authExo.error -eq $false)) {
                        return @{
                            error             = $false
                            AccessToken       = $script:AuthorizationToken
                            AuthHeader        = $script:AuthorizationHeader
                            AccessTokenExo    = $script:ExoAuthorizationToken
                            AuthHeaderExo     = $script:ExoAuthorizationHeader
                            AppAccessToken    = $null
                            AppAuthHeader     = $null
                            AppAccessTokenExo = $null
                            AppAuthHeaderExo  = $null
                        }
                    } else {
                        if ($authExo -and ($authExo.error -ne $false)) {
                            throw "No Exchange Online token: $($authExo.error)"
                        } else {
                            throw 'No Exchange Online token'
                        }
                    }
                } else {
                    return @{
                        error             = $false
                        AccessToken       = $null
                        AuthHeader        = $null
                        AccessTokenExo    = $auth.AccessToken
                        AuthHeaderExo     = $script:ExoAuthorizationHeader
                        AppAccessToken    = $null
                        AppAuthHeader     = $null
                        AppAccessTokenExo = $null
                        AppAuthHeaderExo  = $null
                    }
                }
            } catch {
                Write-Host "$($indent)  Error: $($error[0])"

                return @{
                    error             = ($error[0] | Out-String)
                    AccessToken       = $null
                    authHeader        = $null
                    AccessTokenExo    = $null
                    authHeaderExo     = $null
                    AppAccessToken    = $null
                    AppAuthHeader     = $null
                    AppAccessTokenExo = $null
                    AppAuthHeaderExo  = $null
                }
            }
        }
    }
}


function GraphDomainToTenantID {
    param (
        [string]$domain = 'explicitconsulting.at',
        [uri]$SpecificGraphApiEndpointOnly = $null
    )

    if (-not $script:GraphDomainToTenantIDCache) {
        $script:GraphDomainToTenantIDCache = @{}
    }

    if (-not $script:GraphDomainToCloudInstanceCache) {
        $script:GraphDomainToCloudInstanceCache = @{}
    }

    $domain = $domain.Trim().ToLower()


    # If $domain is a mail address, extract the domain part
    try {
        try { WatchCatchableExitSignal } catch { }

        $tempDomain = [mailaddress]$domain

        if ($tempDomain.Host) {
            $domain = $tempDomain.Host
        }
    } catch {}

    # If $domain is a URL, extract the DNS safe host
    try {
        try { WatchCatchableExitSignal } catch { }

        $tempDomain = [uri]$domain
        if ($tempDomain.DnsSafeHost) {
            $domain = $tempDomain.DnsSafeHost
        }
    } catch {
        # Not a URI, do nothing
    }

    try { WatchCatchableExitSignal } catch { }

    foreach ($SharePointDomain in @('sharepoint.com', 'sharepoint.us', 'dps.mil', 'sharepoint-mil.us', 'sharepoint.cn')) {
        if ($domain.EndsWith("-my.$($SharePointDomain)")) {
            $domain = $domain -ireplace "-my.$($SharePointDomain)", '.onmicrosoft.com'

            break
        }
    }

    try { WatchCatchableExitSignal } catch { }

    if ([string]::IsNullOrWhitespace($domain)) {
        return
    }

    if ($script:GraphDomainToTenantIDCache.ContainsKey($domain)) {
        return $script:GraphDomainToTenantIDCache[$domain]
    }

    try {
        try { WatchCatchableExitSignal } catch { }

        $local:result = Invoke-RestMethod -UseBasicParsing -Uri "https://odc.officeapps.live.com/odc/v2.1/federationprovider?domain=$($domain)"

        try { WatchCatchableExitSignal } catch { }

        $script:GraphDomainToTenantIDCache[$domain] = $local:result.tenantId

        if ($null -eq $local:result.tenantId) {
            return
        }

        if (
            $(
                if ([string]::IsNullOrWhitespace($SpecificGraphApiEndpointOnly)) {
                    $true
                } else {
                    if ([uri]$local:result.graph -ieq [uri]$SpecificGraphApiEndpointOnly) {
                        $true
                    } else {
                        $false
                    }
                }
            )
        ) {
            $script:GraphDomainToCloudInstanceCache[$domain] = $script:GraphDomainToCloudInstanceCache[$local:result.tenantId] = switch ($local:result.authority_host) {
                'login.microsoftonline.com' { 'AzurePublic'; break }
                'login.chinacloudapi.cn' { 'AzureChina'; break }
                'login.microsoftonline.us' { 'AzureUsGovernment'; break }
                default { 'AzurePublic' }
            }

            return $local:result.tenantId
        } else {
            return
        }
    } catch {
        $script:GraphDomainToTenantIDCache[$domain] = $null

        return
    }
}


function GraphSwitchContext {
    param (
        $TenantID = $null
    )

    try { WatchCatchableExitSignal } catch { }

    try {
        if ($null -eq $TenantID -and $script:GraphUser) {
            $TenantID = GraphDomainToTenantID -domain ($script:GraphUser -split '@')[1]
        } else {
            $TenantID = GraphDomainToTenantID -domain $TenantID
        }

        try { WatchCatchableExitSignal } catch { }

        if ($TenantID -and $script:GraphTokenDictionary.ContainsKey($TenantID)) {
            $script:GraphToken = $script:GraphTokenDictionary[$TenantID]

            $script:AuthorizationToken = $script:GraphTokenDictionary[$TenantID].AccessToken
            $script:ExoAuthorizationToken = $script:GraphTokenDictionary[$TenantID].AccessTokenExo

            $script:AuthorizationHeader = $(
                if ($script:GraphTokenDictionary[$TenantID].AuthHeader -is [hashtable]) {
                    $script:GraphTokenDictionary[$TenantID].AuthHeader
                } else {
                    @{
                        Authorization = $script:GraphTokenDictionary[$TenantID].AuthHeader
                    }
                }
            )

            $script:ExoAuthorizationHeader = $(
                if ($script:GraphTokenDictionary[$TenantID].AuthHeaderExo -is [hashtable]) {
                    $script:GraphTokenDictionary[$TenantID].AuthHeaderExo
                } else {
                    @{
                        Authorization = $script:GraphTokenDictionary[$TenantID].AuthHeaderExo
                    }
                }
            )

            $script:AppAuthorizationToken = $script:GraphTokenDictionary[$TenantID].AppAccessToken
            $script:AppExoAuthorizationToken = $script:GraphTokenDictionary[$TenantID].AppAccessTokenExo

            $script:AppAuthorizationHeader = $(
                if ($script:GraphTokenDictionary[$TenantID].AppAuthHeader -is [hashtable]) {
                    $script:GraphTokenDictionary[$TenantID].AppAuthHeader
                } else {
                    @{
                        Authorization = $script:GraphTokenDictionary[$TenantID].AppAuthHeader
                    }
                }
            )
            $script:AppExoAuthorizationHeader = $(
                if ($script:GraphTokenDictionary[$TenantID].AppAuthHeaderExo -is [hashtable]) {
                    $script:GraphTokenDictionary[$TenantID].AppAuthHeaderExo
                } else {
                    @{
                        Authorization = $script:GraphTokenDictionary[$TenantID].AppAuthHeaderExo
                    }
                }
            )
        }

        if ($TenantID -and $script:GraphDomainToCloudInstanceCache.ContainsKey($TenantID)) {
            $CloudEnvironment = $script:GraphDomainToCloudInstanceCache[$TenantID]
        }

        # Endpoints from https://github.com/microsoft/CSS-Exchange/blob/main/Shared/AzureFunctions/Get-CloudServiceEndpoint.ps1
        # Environment names must match https://learn.microsoft.com/en-us/dotnet/api/microsoft.identity.client.azurecloudinstance?view=msal-dotnet-latest
        switch ($CloudEnvironment) {
            { $_ -iin @('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC') } {
                $script:CloudEnvironmentEnvironmentName = 'AzurePublic'
                $script:CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.com'
                $script:CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office.com'
                $script:CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.outlook.com'
                $script:CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.com'
                $script:CloudEnvironmentSharePointOnlineDomains = @('sharepoint.com')
                break
            }

            { $_ -iin @('AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4') } {
                $script:CloudEnvironmentEnvironmentName = 'AzureUSGovernment'
                $script:CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.us'
                $script:CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office365.us'
                $script:CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.office365.us'
                $script:CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.us'
                $script:CloudEnvironmentSharePointOnlineDomains = @('sharepoint.us')
                break
            }

            { $_ -iin @('AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5') } {
                $script:CloudEnvironmentEnvironmentName = 'AzureUSGovernment'
                $script:CloudEnvironmentGraphApiEndpoint = 'https://dod-graph.microsoft.us'
                $script:CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook-dod.office365.us'
                $script:CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s-dod.office365.us'
                $script:CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.us'
                $script:CloudEnvironmentSharePointOnlineDomains = @('dps.mil', 'sharepoint-mil.us')
                break
            }

            { $_ -iin @('China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud') } {
                $script:CloudEnvironmentEnvironmentName = 'AzureChina'
                $script:CloudEnvironmentGraphApiEndpoint = 'https://microsoftgraph.chinacloudapi.cn'
                $script:CloudEnvironmentExchangeOnlineEndpoint = 'https://partner.outlook.cn'
                $script:CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.partner.outlook.cn'
                $script:CloudEnvironmentAzureADEndpoint = 'https://login.partner.microsoftonline.cn'
                $script:CloudEnvironmentSharePointOnlineDomains = @('sharepoint.cn')
                break
            }

            default {
                $script:CloudEnvironmentEnvironmentName = 'AzurePublic'
                $script:CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.com'
                $script:CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office.com'
                $script:CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.outlook.com'
                $script:CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.com'
                $script:CloudEnvironmentSharePointOnlineDomains = @('sharepoint.com')
                break
            }
        }
    } catch {
        $script:GraphToken = $null
        $script:AuthorizationToken = $null
        $script:ExoAuthorizationToken = $null
        $script:AuthorizationHeader = $null
        $script:ExoAuthorizationHeader = $null
        $script:AppAuthorizationHeader = $null
        $script:AppExoAuthorizationHeader = $null
    }

    try { WatchCatchableExitSignal } catch { }
}


function GraphGetTokenWrapper {
    param (
        $indent = ''
    )

    if (-not ($GraphClientIDOriginal -is [string])) {
        $tempGraphClientIDOriginal = @()

        foreach ($item in $GraphClientIDOriginal) {
            $tempGraphClientIDOriginal += , @($item)
        }

        $GraphClientIDOriginal = $tempGraphClientIDOriginal
    }

    if ($GraphClientIDOriginal -is [array]) {
        $GraphClientIDOriginal = @(
            $GraphClientIDOriginal | Where-Object {
                $($_ -is [array]) -and
                $($_[0]) -and
                $($_[1]) -and
                $($null -ne (GraphDomainToTenantID -domain $_[0]))
            } | ForEach-Object {
                , @(
                    "$($_[0])".Trim().ToLower(),
                    "$($_[1])".Trim().ToLower()
                )
            }
        )
    }

    GraphSwitchContext -TenantID $null

    if (
        $($GraphClientIDOriginal -is [string])
    ) {
        $GraphClientID = $GraphClientIDOriginal
        $script:GraphTenantId = 'organizations'

        try {
            try { WatchCatchableExitSignal } catch { }

            $script:GraphToken = GraphGetToken -indent "$($indent)"

            try { WatchCatchableExitSignal } catch { }
        } catch {
            $script:GraphToken = $null
        }
    } elseif (
        $($GraphClientIDOriginal -is [array])
    ) {
        try { WatchCatchableExitSignal } catch { }

        if (-not $SimulateAndDeployGraphCredentialFile) {
            for ($i = 0; $i -lt $GraphClientIDOriginal.Count; $i++) {
                try { WatchCatchableExitSignal } catch { }

                $script:GraphTenantId = GraphDomainToTenantID -domain $GraphClientIDOriginal[$i][0]
                $GraphClientID = $GraphClientIDOriginal[$i][1]

                Write-Host "$($indent)Tenant $($GraphClientIDOriginal[$i][0]) ($($script:GraphTenantId))"

                try {
                    try { WatchCatchableExitSignal } catch { }


                    $script:GraphTokenDictionary[$script:GraphTenantId] = $script:GraphToken = GraphGetToken -indent "$($indent)  "

                    try { WatchCatchableExitSignal } catch { }
                } catch {
                    $script:GraphTokenDictionary[$script:GraphTenantId] = $script:GraphToken = $null
                }

                GraphSwitchContext -TenantID $null
            }
        } else {
            GraphGetToken -indent "$($indent)  "
        }
    }

    GraphSwitchContext -TenantID $null
}


function GraphGenericQuery {
    [CmdletBinding()]

    param (
        [Parameter(Mandatory = $true)]
        [string]$method,

        [Parameter(Mandatory = $true)]
        [uri]$uri,

        [Parameter(Mandatory = $false)]
        [AllowEmptyString()] [string]$body = $null,

        [Parameter(Mandatory = $false)]
        $authHeader,

        $GraphContext = $null
    )

    GraphSwitchContext -TenantID $GraphContext

    if (-not $authHeader) {
        $authHeader = $(if ($SimulateUser -and $SimulateAndDeployGraphCredentialFile) { $script:AppAuthorizationHeader } else { $script:AuthorizationHeader })
    }

    $error.clear()

    try {
        $requestBody = @{
            Method      = $method
            Uri         = $uri
            Headers     = $authHeader
            ContentType = 'application/json; charset=utf-8'
        }

        if ($body) {
            $requestBody['Body'] = $body
        }

        $requestBody['Headers']['x-overridetimestamp'] = 'true'

        $requestBody['Headers']['content-type'] = 'Application/Json; charset=utf-8'

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error  = $error[0] | Out-String
            result = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error  = $false
            result = $local:x
        }
    } else {
        return @{
            error  = $error[0] | Out-String
            result = $null
        }
    }
}


function GraphGetMe {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s)
    #   Delegated: User.Read.All
    #   Application: User.Read.All (/me is not supported in applications)

    GraphSwitchContext -TenantID $script:GraphUser

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/me?`$select=" + [System.Net.WebUtility]::UrlEncode(($GraphUserProperties -join ','))
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error = $error[0] | Out-String
            me    = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            me    = $local:x
        }
    } else {
        return @{
            error = $error[0] | Out-String
            me    = $null
        }
    }
}


function GraphGetUpnFromSmtp($user, $authHeader) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    GraphSwitchContext -TenantID $user

    if (-not $authHeader) {
        $authHeader = $script:AuthorizationHeader
    }


    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$filter=proxyAddresses/any(x:x eq 'smtp:$($user)')"
            Headers     = $authHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error      = $false
            properties = $local:x
        }
    } else {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }
}


function GraphGetUserProperties($user, $authHeader) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    GraphSwitchContext -TenantID $user

    if (-not $authHeader) {
        $authHeader = $script:AuthorizationHeader
    }

    try { WatchCatchableExitSignal } catch { }

    $user = GraphGetUpnFromSmtp -user $user -authHeader $authHeader

    if ($user.properties.value.userprincipalname) {
        try {
            $requestBody = @{
                Method      = 'Get'
                Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user.properties.value.userprincipalname)?`$select=" + [System.Net.WebUtility]::UrlEncode($(@($GraphUserProperties | Select-Object -Unique) -join ','))
                Headers     = $authHeader
                ContentType = 'Application/Json; charset=utf-8'
            }

            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $local:x = @()
            $local:uri = $null

            do {
                try { WatchCatchableExitSignal } catch { }

                if ($local:uri) {
                    $requestBody['Uri'] = $local:uri
                }

                $local:pagedResults = Invoke-RestMethod @requestBody
                $local:x += $local:pagedResults

                if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                    $local:uri = $null
                } else {
                    $local:uri = $local:pagedResults.'@odata.nextlink'
                }
            } until (!($local:uri))


            if (($user.properties.value.userprincipalname -ieq $script:GraphUser) -and ((-not $SimulateUser) -or ($SimulateUser -and $SimulateAndDeployGraphCredentialFile)) -and (($SetCurrentUserOOFMessage -eq $true) -or ($SetCurrentUserOutlookWebSignature -eq $true) -or ($MirrorCloudSignatures -ne $false))) {
                try {
                    $requestBody = @{
                        Method      = 'Get'
                        Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user.properties.value.userprincipalname)?`$select=mailboxsettings"
                        Headers     = $authHeader
                        ContentType = 'Application/Json; charset=utf-8'
                    }

                    $OldProgressPreference = $ProgressPreference
                    $ProgressPreference = 'SilentlyContinue'

                    $local:y = @()

                    do {
                        try { WatchCatchableExitSignal } catch { }

                        if ($local:uri) {
                            $requestBody['Uri'] = $local:uri
                        }

                        $local:pagedResults = Invoke-RestMethod @requestBody
                        $local:y += $local:pagedResults

                        if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                            $local:uri = $null
                        } else {
                            $local:uri = $local:pagedResults.'@odata.nextlink'
                        }
                    } until (!($local:uri))

                    $local:x | Add-Member -MemberType NoteProperty -Name 'mailboxSettings' -Value $local:y.mailboxSettings -Force
                } catch {
                    Write-Host $error[0]
                    Write-Host "      Problem getting mailboxSettings for '$($script:GraphUser)' from Microsoft Graph." -ForegroundColor Yellow
                    Write-Host '      This is a Microsoft Graph API problem, which can only be solved by Microsoft itself.' -ForegroundColor Yellow
                    Write-Host '      Disabling SetCurrentUserOutlookWebSignature and SetCurrentUserOOFMessage to be able to continue.' -ForegroundColor Yellow

                    $SetCurrentUserOutlookWebSignature = $false
                    $SetCurrentUserOOFMessage = $false
                }
            }

            $ProgressPreference = $OldProgressPreference
        } catch {
            return @{
                error      = $error[0] | Out-String
                properties = $null
            }
        }

        try { WatchCatchableExitSignal } catch { }

        if (($user.properties.value.userprincipalname -ieq $script:GraphUser) -and ($SimulateUser -and $SimulateAndDeployGraphCredentialFile -and ($authHeader -eq $script:AuthorizationHeader))) {
            $temp = GraphGetUserProperties -user $($user.properties.value.userprincipalname) -authHeader $script:AppAuthorizationHeader

            if ($temp.error -eq $false) {
                $local:x = $temp.properties
            } else {
            }
        }

        if ($null -ne $local:x) {
            return @{
                error      = $false
                properties = $local:x
            }
        } else {
            return @{
                error      = $error[0] | Out-String
                properties = $null
            }
        }
    } else {
        return @{
            error      = $user.error
            properties = $null
        }
    }
}


function GraphGetUserManager($user) {
    # Current mailbox manager
    # https://docs.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    GraphSwitchContext -TenantID $user

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/manager"
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error      = $false
            properties = $local:x
        }
    } else {
        return @{
            error      = $error[0] | Out-String
            properties = $null
        }
    }

}


function GraphGetUserTransitiveMemberOf($user) {
    # https://learn.microsoft.com/en-us/graph/api/user-list-transitivememberof?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    GraphSwitchContext -TenantID $user

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/transitiveMemberOf"
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error    = $error[0] | Out-String
            memberof = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error    = $false
            memberof = $local:x
        }
    } else {
        return @{
            error    = $error[0] | Out-String
            memberof = $null
        }
    }
}


function GraphGetUserPhoto($user) {
    # https://docs.microsoft.com/en-us/graph/api/profilephoto-get?view=graph-rest-1.0
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    GraphSwitchContext -TenantID $user

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/photo/`$value"
            Headers     = $script:AuthorizationHeader
            ContentType = 'image/jpg'
        }

        $local:tempFile = (Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ((New-Guid).Guid))
        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $null = Invoke-RestMethod @requestBody -OutFile $local:tempFile

        $ProgressPreference = $OldProgressPreference

        try { WatchCatchableExitSignal } catch { }

        $local:x = [System.IO.File]::ReadAllBytes($local:tempFile)

        Remove-Item -LiteralPath $local:tempFile -Force -ErrorAction SilentlyContinue
    } catch {
        return @{
            error = $error[0] | Out-String
            photo = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            photo = $local:x
        }
    } else {
        return @{
            error = $error[0] | Out-String
            photo = $null
        }
    }
}


function GraphPatchUserMailboxsettings($user, $OOFInternal, $OOFExternal, $authHeader) {
    # https://learn.microsoft.com/en-us/graph/api/user-updatemailboxsettings?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: Mailboxsettings.ReadWrite
    #   Application: Mailboxsettings.ReadWrite

    GraphSwitchContext -TenantID $user

    if (-not $authHeader) {
        $authHeader = $(if ($SimulateUser -and $SimulateAndDeploy -and $SimulateAndDeployGraphCredentialFile) { $script:AppAuthorizationHeader } else { $script:AuthorizationHeader })
    }

    try { WatchCatchableExitSignal } catch { }

    try {
        if ($OOFInternal -or $OOFExternal) {
            $body = @{}
            $body.add('automaticRepliesSetting', @{})

            if ($OOFInternal) { $Body.'automaticRepliesSetting'.add('internalReplyMessage', $OOFInternal) }

            if ($OOFExternal) { $Body.'automaticRepliesSetting'.add('externalReplyMessage', $OOFExternal) }

            $body = ConvertTo-Json -InputObject $body

            $requestBody = @{
                Method      = 'Patch'
                Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users/$($user)/mailboxsettings"
                Headers     = $authHeader
                ContentType = 'Application/Json; charset=utf-8'
                Body        = $body
            }

            $OldProgressPreference = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'

            $null = Invoke-RestMethod @requestBody

            $ProgressPreference = $OldProgressPreference
        }

        return @{
            error = $false
        }
    } catch {
        return @{
            error = $error[0] | Out-String
        }
    }
}


function GraphFilterGroups($filter, $GraphContext = $null) {
    # https://docs.microsoft.com/en-us/graph/api/group-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: GroupMember.Read.All
    #   Application: GroupMember.Read.All

    GraphSwitchContext -TenantID $GraphContext

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/groups?`$select=securityidentifier&`$filter=" + [System.Net.WebUtility]::UrlEncode($filter)
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error  = $error[0] | Out-String
            groups = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error  = $false
            groups = $local:x
        }
    } else {
        return @{
            error  = $error[0] | Out-String
            groups = $null
        }
    }
}


function GraphFilterUsers($filter, $GraphContext = $null) {
    # https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
    # Required permission(s):
    #   Delegated: User.Read.All
    #   Application: User.Read.All

    GraphSwitchContext -TenantID $GraphContext

    try { WatchCatchableExitSignal } catch { }

    try {
        $requestBody = @{
            Method      = 'Get'
            Uri         = "$($script:CloudEnvironmentGraphApiEndpoint)/$($GraphEndpointVersion)/users?`$select=securityidentifier&`$filter=" + [System.Net.WebUtility]::UrlEncode($filter)
            Headers     = $script:AuthorizationHeader
            ContentType = 'Application/Json; charset=utf-8'
        }

        $OldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'

        $local:x = @()
        $local:uri = $null

        do {
            try { WatchCatchableExitSignal } catch { }

            if ($local:uri) {
                $requestBody['Uri'] = $local:uri
            }

            $local:pagedResults = Invoke-RestMethod @requestBody
            $local:x += $local:pagedResults

            if ([string]::IsNullOrWhiteSpace($local:pagedResults.'@odata.nextlink')) {
                $local:uri = $null
            } else {
                $local:uri = $local:pagedResults.'@odata.nextlink'
            }
        } until (!($local:uri))

        $ProgressPreference = $OldProgressPreference
    } catch {
        return @{
            error = $error[0] | Out-String
            users = $null
        }
    }

    if ($null -ne $local:x) {
        return @{
            error = $false
            users = $local:x
        }
    } else {
        return @{
            error = $error[0] | Out-String
            users = $null
        }
    }
}


function GetIniContent ($filePath, $additionalLines) {
    try { WatchCatchableExitSignal } catch { }

    $local:ini = [ordered]@{}
    $local:SectionIndex = -1

    if ($filePath -ne '') {
        try {
            Write-Verbose '    Original INI content'

            foreach ($line in @(@((ConvertEncoding -InFile $FilePath -InIsHtml $false) -split '\r?\n') + @($additionalLines -split '\r?\n'))) {
                Write-Verbose "      $line"
                switch -regex ($line) {
                    # Comments starting with ; or # or //, or empty line, whitespace(s) before are ignored
                    '(^\s*(;|#|//))|(^\s*$)' { continue }

                    # Section in square brackets, whitespace(s) before and after brackets are ignored
                    '^\s*\[(.+)\]\s*' {
                        $local:section = ($matches[1]).trim().trim('"').trim('''')
                        if ($null -ne $local:section) {
                            $local:SectionIndex++
                            $local:ini["$($local:SectionIndex)"] = [ordered]@{ '<Set-OutlookSignatures template>' = $local:section }
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
            Write-Host $error[0]
            Write-Host "Error accessing '$FilePath'. Exit." -ForegroundColor red
            $script:ExitCode = 31
            $script:ExitCodeDescription = "Error accessing '$FilePath'."
            exit
        }
    }

    try { WatchCatchableExitSignal } catch { }

    # default values for <Set-OutlookSignatures configuration>
    if (
        $(
            try {
                ((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))
            } catch {
                $false
            }
        )
    ) {

        if (
            -not $(
                try {
                    $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortOrder']
                } catch {
                    $false
                }
            )
        ) {
            $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortOrder'] = 'AsInThisFile'
        }

        if (
            -not $(
                try {
                    $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortCulture']
                } catch {
                    $false
                }
            )
        ) {
            $((@($local:ini[($local:ini.GetEnumerator().name)] | Where-Object { $_['<Set-OutlookSignatures template>'] -ieq '<Set-OutlookSignatures configuration>' }) | Select-Object -Last 1))['SortCulture'] = '127'
        }
    } else {
        $local:ini["$($local:ini.Count)"] = [ordered]@{
            '<Set-OutlookSignatures template>' = '<Set-OutlookSignatures configuration>'
            'SortOrder'                        = 'AsInThisFile'
            'SortCulture'                      = '127'
        }
    }

    try { WatchCatchableExitSignal } catch { }

    return $local:ini
}


function ConvertPath ([ref]$path) {
    try { WatchCatchableExitSignal } catch { }

    if ($path) {
        if (($path.value.StartsWith('https://', 'CurrentCultureIgnoreCase')) -or ($path.value -ilike '*@SSL\*')) {
            if (-not [System.Uri]::IsWellFormedUriString($path.value, [System.UriKind]::Absolute)) {
                $path.value = ([uri]($path.value -ireplace '@SSL\\', '/' -ireplace '^\\\\', 'https://' -ireplace '\\', '/')).AbsoluteUri
            }
            $path.value = ([uri]$path.value).GetLeftPart([System.UriPartial]::Path) -ireplace "$(([uri]$path.value).GetLeftPart([System.UriPartial]::Authority))/:\S:/\S", $(([uri]$path.value).GetLeftPart([System.UriPartial]::Authority))
            $path.value = ([uri]::UnescapeDataString($path.value) -ireplace ('https://', '\\'))
            $path.value = ([System.URI]$path.value).AbsoluteURI -ireplace 'file:\/\/(.*?)\/(.*)', '\\${1}@SSL\${2}' -ireplace '/', '\'
            $path.value = [uri]::UnescapeDataString($path.value)
        } else {
            $path.value = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path.value)

            if ($IsWindows) {
                $path.value = ([System.URI]$path.value).absoluteuri -ireplace 'file:///', '' -ireplace 'file://', '\\' -ireplace '/', '\'
                $path.value = [uri]::UnescapeDataString($path.value)
            }
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function RemoveItemAlternativeRecurse {
    # Function to avoid problems with OneDrive throwing "Access to the cloud file is denied"

    param(
        [alias('LiteralPath')][string] $Path,
        [switch] $SkipFolder # when $Path is a folder, do not delete $path, only it's content
    )

    try { WatchCatchableExitSignal } catch { }

    $local:ToDelete = @()

    if (Test-Path -LiteralPath $path) {
        foreach ($SinglePath in @(Get-Item -LiteralPath $Path)) {
            try { WatchCatchableExitSignal } catch { }

            if (Test-Path -LiteralPath $SinglePath -PathType Container) {
                if (-not $SkipFolder) {
                    $local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Culture 127 -Property PSIsContainer, @{expression = { $_.FullName.split([IO.Path]::DirectorySeparatorChar).count }; descending = $true }, fullname)
                    $local:ToDelete += @(Get-Item -LiteralPath $SinglePath -Force)
                } else {
                    $local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Culture 127 -Property PSIsContainer, @{expression = { $_.FullName.split([IO.Path]::DirectorySeparatorChar).count }; descending = $true }, fullname)
                }
            } elseif (Test-Path -LiteralPath $SinglePath -PathType Leaf) {
                $local:ToDelete += (Get-Item -LiteralPath $SinglePath -Force)
            }
        }
    } else {
        # Item to delete does not exist, nothing to do
    }

    foreach ($SingleItemToDelete in $local:ToDelete) {
        try { WatchCatchableExitSignal } catch { }

        try {
            if ((Test-Path -LiteralPath $SingleItemToDelete.FullName) -eq $true) {
                Remove-Item -LiteralPath $SingleItemToDelete.FullName -Force -Recurse
            }
        } catch {
            Write-Verbose "Could not delete $($SingleItemToDelete.FullName), error: $($_.Exception.Message)"
            Write-Verbose $_
        }
    }

    try { WatchCatchableExitSignal } catch { }
}


function ParseJwtToken {
    # Idea for this code: https://www.michev.info/blog/post/2140/decode-jwt-access-and-id-tokens-via-powershell

    [cmdletbinding()]
    param([Parameter(Mandatory = $true)][string]$token)

    try { WatchCatchableExitSignal } catch { }

    # Validate as per https://tools.ietf.org/html/rfc7519
    # Access and ID tokens are fine, Refresh tokens will not work
    if (!$token.Contains('.') -or !$token.StartsWith('eyJ')) {
        return @{
            error   = 'Invalid token'
            header  = $null
            payload = $null
        }
    } else {
        # Header
        $tokenheader = $token.Split('.')[0].Replace('-', '+').Replace('_', '/')

        # Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
        while ($tokenheader.Length % 4) { $tokenheader += '=' }

        # Convert from Base64 encoded string to PSObject all at once
        $tokenHeader = [System.Text.Encoding]::UTF8.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json

        # Payload
        $tokenPayload = $token.Split('.')[1].Replace('-', '+').Replace('_', '/')

        # Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
        while ($tokenPayload.Length % 4) { $tokenPayload += '=' }

        # Convert to Byte array
        $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)

        # Convert to string array
        $tokenArray = [System.Text.Encoding]::UTF8.GetString($tokenByteArray)

        # Convert from JSON to PSObject
        $tokenPayload = $tokenArray | ConvertFrom-Json

        return @{
            error   = $false
            header  = $tokenHeader
            payload = $tokenPayload
        }
    }
}


### ▼▼▼ WatchCatchableExitSignal initiation code below ▼▼▼
##
#
# Place this code in your main script, as early in the code as possible
#
# Call WatchCatchableExitSignal wherever you want to gracefully exit in case of
#   - a Logoff/Reboot/Shutdown message on Windows
#   - a catchable POSIX signal on Linux and macOS
#
# If $WatchCatchableExitSignalNonExitScriptBlock is of type [scriptblock],
#   it is executed when no catchable exit signal is detected.
#
# Clean-up is triggered by WatchCatchableExitSignal running the "$script:ExitCode = 1; $script:ExitCodeDescription = ''; exit" command
#   This triggers the Finally part of a Try/Catch/Finally block
#
# Place the following two lines of code at the end of your clean-up routine
#   WatchCatchableExitSignal -CleanupDone
#

$global:WatchCatchableExitSignalStatus = [hashtable]::Synchronized(@{})
$global:WatchCatchableExitSignalStatus[0] = 'Nothing detected yet'
# Possible values for $global:WatchCatchableExitSignalStatus[0]
#   "Nothing detected yet" when no catchable exit signal has been found until now
#   "Detected '<description>', initiate clean-up and exit" when a catchable exit signal has been found
#   "Clean-up done" after clean-up is done

$WatchCatchableExitSignalRunspace = [runspacefactory]::CreateRunspace()
$WatchCatchableExitSignalRunspace.Open()
$WatchCatchableExitSignalRunspace.SessionStateProxy.SetVariable('WatchCatchableExitSignalStatus', $global:WatchCatchableExitSignalStatus)
$WatchCatchableExitSignalPowershell = [powershell]::Create()
$WatchCatchableExitSignalPowershell.Runspace = $WatchCatchableExitSignalRunspace

if ($IsWindows -or (-not (Test-Path -LiteralPath 'variable:IsWindows'))) {
    $WatchCatchableExitSignalForm = $null

    $WatchCatchableExitSignalRunspace.SessionStateProxy.SetVariable('WatchCatchableExitSignalForm', [ref]$WatchCatchableExitSignalForm)

    $null = $WatchCatchableExitSignalPowershell.AddScript(
        {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -TypeDefinition @'
            using System;
            using System.Windows.Forms;
            using System.Management.Automation;
            using System.Management.Automation.Runspaces;
            using System.Collections.ObjectModel;

            public class CustomForm : Form {
                public event Action<Message> EndSessionInitiateCleanup;

                protected override CreateParams CreateParams {
                    // Hide the window from Alt-Tab
                    get {
                        CreateParams cp = base.CreateParams;
                        cp.ExStyle |= 0x80;  // WS_EX_TOOLWINDOW
                        return cp;
                    }
                }

                protected override void WndProc(ref Message m) {
                    if (EndSessionInitiateCleanup != null && !this.IsDisposed) {
                        try {
                            EndSessionInitiateCleanup.Invoke(m);
                        } catch {
                            // Do nothing
                        }
                    }

                    base.WndProc(ref m);
                }
            }
'@ -ReferencedAssemblies $(
                if ($($PSVersionTable.PSEdition) -ieq 'Core') {
                    $(@('System.Windows.Forms', 'System.ComponentModel.Primitives', 'System.Management.Automation', 'System.Windows.Forms.Primitives'))
                } else {
                    $(@('System.Windows.Forms', 'System.ComponentModel.Primitives', 'System.Management.Automation'))
                }
            )

            $formRef = $ExecutionContext.SessionState.PSVariable.GetValue('WatchCatchableExitSignalForm')
            $formRef.Value = [CustomForm]::new()
            $formRef.Value.Text = 'Set-OutlookSignatures non-blocking window for WM_* detection'
            $formRef.Value.Width = 300
            $formRef.Value.Height = 300
            $formRef.Value.ShowInTaskbar = $false
            $formRef.Value.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
            $formRef.Value.Opacity = 0
            $formRef.Value.WindowState = [System.Windows.Forms.FormWindowState]::Minimized

            # Event handler
            $formRef.Value.add_EndSessionInitiateCleanup(
                {
                    param($message)

                    try {
                        $WindowsMessagesByDecimal = @{
                            0 = 'WM_NULL'; 2 = 'WM_DESTROY'; 3 = 'WM_MOVE'; 5 = 'WM_SIZE'; 6 = 'WM_ACTIVATE'; 7 = 'WM_SETFOCUS'; 8 = 'WM_KILLFOCUS'; 10 = 'WM_ENABLE'; 11 = 'WM_SETREDRAW'; 12 = 'WM_SETTEXT'; 13 = 'WM_GETTEXT'; 14 = 'WM_GETTEXTLENGTH'; 15 = 'WM_PAINT'; 16 = 'WM_CLOSE'; 17 = 'WM_QUERYENDSESSION'; 18 = 'WM_QUIT'; 19 = 'WM_QUERYOPEN'; 20 = 'WM_ERASEBKGND'; 21 = 'WM_SYSCOLORCHANGE'; 22 = 'WM_ENDSESSION'; 23 = 'WM_SYSTEMERROR'; 24 = 'WM_SHOWWINDOW'; 25 = 'WM_CTLCOLOR'; 26 = 'WM_SETTINGCHANGE'; 27 = 'WM_DEVMODECHANGE'; 28 = 'WM_ACTIVATEAPP'; 29 = 'WM_FONTCHANGE'; 30 = 'WM_TIMECHANGE'; 31 = 'WM_CANCELMODE'; 32 = 'WM_SETCURSOR'; 33 = 'WM_MOUSEACTIVATE'; 34 = 'WM_CHILDACTIVATE'; 35 = 'WM_QUEUESYNC'; 36 = 'WM_GETMINMAXINFO'; 38 = 'WM_PAINTICON'; 39 = 'WM_ICONERASEBKGND'; 40 = 'WM_NEXTDLGCTL'; 42 = 'WM_SPOOLERSTATUS'; 43 = 'WM_DRAWITEM'; 44 = 'WM_MEASUREITEM'; 45 = 'WM_DELETEITEM'; 46 = 'WM_VKEYTOITEM'; 47 = 'WM_CHARTOITEM'; 48 = 'WM_SETFONT'; 49 = 'WM_GETFONT'; 50 = 'WM_SETHOTKEY'; 51 = 'WM_GETHOTKEY'; 55 = 'WM_QUERYDRAGICON'; 57 = 'WM_COMPAREITEM'; 65 = 'WM_COMPACTING'; 70 = 'WM_WINDOWPOSCHANGING'; 71 = 'WM_WINDOWPOSCHANGED'; 72 = 'WM_POWER'; 74 = 'WM_COPYDATA'; 75 = 'WM_CANCELJOURNAL'; 78 = 'WM_NOTIFY'; 80 = 'WM_INPUTLANGCHANGEREQUEST'; 81 = 'WM_INPUTLANGCHANGE'; 82 = 'WM_TCARD'; 83 = 'WM_HELP'; 84 = 'WM_USERCHANGED'; 85 = 'WM_NOTIFYFORMAT'; 123 = 'WM_CONTEXTMENU'; 124 = 'WM_STYLECHANGING'; 125 = 'WM_STYLECHANGED'; 126 = 'WM_DISPLAYCHANGE'; 127 = 'WM_GETICON'; 128 = 'WM_SETICON'; 129 = 'WM_NCCREATE'; 130 = 'WM_NCDESTROY'; 131 = 'WM_NCCALCSIZE'; 132 = 'WM_NCHITTEST'; 133 = 'WM_NCPAINT'; 134 = 'WM_NCACTIVATE'; 135 = 'WM_GETDLGCODE'; 160 = 'WM_NCMOUSEMOVE'; 161 = 'WM_NCLBUTTONDOWN'; 162 = 'WM_NCLBUTTONUP'; 163 = 'WM_NCLBUTTONDBLCLK'; 164 = 'WM_NCRBUTTONDOWN'; 165 = 'WM_NCRBUTTONUP'; 166 = 'WM_NCRBUTTONDBLCLK'; 167 = 'WM_NCMBUTTONDOWN'; 168 = 'WM_NCMBUTTONUP'; 169 = 'WM_NCMBUTTONDBLCLK'; 256 = 'WM_KEYDOWN'; 257 = 'WM_KEYUP'; 258 = 'WM_CHAR'; 259 = 'WM_DEADCHAR'; 260 = 'WM_SYSKEYDOWN'; 261 = 'WM_SYSKEYUP'; 262 = 'WM_SYSCHAR'; 263 = 'WM_SYSDEADCHAR'; 264 = 'WM_KEYLAST'; 269 = 'WM_IME_STARTCOMPOSITION'; 270 = 'WM_IME_ENDCOMPOSITION'; 271 = 'WM_IME_COMPOSITION'; 272 = 'WM_INITDIALOG'; 273 = 'WM_COMMAND'; 274 = 'WM_SYSCOMMAND'; 275 = 'WM_TIMER'; 276 = 'WM_HSCROLL'; 277 = 'WM_VSCROLL'; 278 = 'WM_INITMENU'; 279 = 'WM_INITMENUPOPUP'; 287 = 'WM_MENUSELECT'; 288 = 'WM_MENUCHAR'; 289 = 'WM_ENTERIDLE'; 306 = 'WM_CTLCOLORMSGBOX'; 307 = 'WM_CTLCOLOREDIT'; 308 = 'WM_CTLCOLORLISTBOX'; 309 = 'WM_CTLCOLORBTN'; 310 = 'WM_CTLCOLORDLG'; 311 = 'WM_CTLCOLORSCROLLBAR'; 312 = 'WM_CTLCOLORSTATIC'; 512 = 'WM_MOUSEMOVE'; 513 = 'WM_LBUTTONDOWN'; 514 = 'WM_LBUTTONUP'; 515 = 'WM_LBUTTONDBLCLK'; 516 = 'WM_RBUTTONDOWN'; 517 = 'WM_RBUTTONUP'; 518 = 'WM_RBUTTONDBLCLK'; 519 = 'WM_MBUTTONDOWN'; 520 = 'WM_MBUTTONUP'; 521 = 'WM_MBUTTONDBLCLK'; 522 = 'WM_MOUSEWHEEL'; 526 = 'WM_MOUSEHWHEEL'; 528 = 'WM_PARENTNOTIFY'; 529 = 'WM_ENTERMENULOOP'; 530 = 'WM_EXITMENULOOP'; 531 = 'WM_NEXTMENU'; 532 = 'WM_SIZING'; 533 = 'WM_CAPTURECHANGED'; 534 = 'WM_MOVING'; 536 = 'WM_POWERBROADCAST'; 537 = 'WM_DEVICECHANGE'; 544 = 'WM_MDICREATE'; 545 = 'WM_MDIDESTROY'; 546 = 'WM_MDIACTIVATE'; 547 = 'WM_MDIRESTORE'; 548 = 'WM_MDINEXT'; 549 = 'WM_MDIMAXIMIZE'; 550 = 'WM_MDITILE'; 551 = 'WM_MDICASCADE'; 552 = 'WM_MDIICONARRANGE'; 553 = 'WM_MDIGETACTIVE'; 560 = 'WM_MDISETMENU'; 561 = 'WM_ENTERSIZEMOVE'; 562 = 'WM_EXITSIZEMOVE'; 563 = 'WM_DROPFILES'; 564 = 'WM_MDIREFRESHMENU'; 641 = 'WM_IME_SETCONTEXT'; 642 = 'WM_IME_NOTIFY'; 643 = 'WM_IME_CONTROL'; 644 = 'WM_IME_COMPOSITIONFULL'; 645 = 'WM_IME_SELECT'; 646 = 'WM_IME_CHAR'; 656 = 'WM_IME_KEYDOWN'; 657 = 'WM_IME_KEYUP'; 673 = 'WM_MOUSEHOVER'; 674 = 'WM_NCMOUSELEAVE'; 675 = 'WM_MOUSELEAVE'; 768 = 'WM_CUT'; 769 = 'WM_COPY'; 770 = 'WM_PASTE'; 771 = 'WM_CLEAR'; 772 = 'WM_UNDO'; 773 = 'WM_RENDERFORMAT'; 774 = 'WM_RENDERALLFORMATS'; 775 = 'WM_DESTROYCLIPBOARD'; 776 = 'WM_DRAWCLIPBOARD'; 777 = 'WM_PAINTCLIPBOARD'; 778 = 'WM_VSCROLLCLIPBOARD'; 779 = 'WM_SIZECLIPBOARD'; 780 = 'WM_ASKCBFORMATNAME'; 781 = 'WM_CHANGECBCHAIN'; 782 = 'WM_HSCROLLCLIPBOARD'; 783 = 'WM_QUERYNEWPALETTE'; 784 = 'WM_PALETTEISCHANGING'; 785 = 'WM_PALETTECHANGED'; 786 = 'WM_HOTKEY'; 791 = 'WM_PRINT'; 792 = 'WM_PRINTCLIENT'; 856 = 'WM_HANDHELDFIRST'; 863 = 'WM_HANDHELDLAST'; 896 = 'WM_PENWINFIRST'; 911 = 'WM_PENWINLAST'; 912 = 'WM_COALESCE_FIRST'; 927 = 'WM_COALESCE_LAST'; 992 = 'WM_DDE_INITIATE'; 993 = 'WM_DDE_TERMINATE'; 994 = 'WM_DDE_ADVISE'; 995 = 'WM_DDE_UNADVISE'; 996 = 'WM_DDE_ACK'; 997 = 'WM_DDE_DATA'; 998 = 'WM_DDE_REQUEST'; 999 = 'WM_DDE_POKE'; 1000 = 'WM_DDE_EXECUTE'
                        }

                        if (
                            $(
                                $($WindowsMessagesByDecimal[$($message.Msg)] -ieq 'WM_ENDSESSION') -and
                                $($message.WParam -ne [IntPtr]::Zero)
                            ) -or
                            $($WindowsMessagesByDecimal[$($message.Msg)] -ieq 'WM_QUERYENDSESSION')
                        ) {
                            # Logoff/Reboot/Shutdown will happen.
                            # Set status, wait for clean-up and then return 0.
                            $global:WatchCatchableExitSignalStatus[0] = "Detected '$(@(@($($message.Msg), $($WindowsMessagesByDecimal[$($message.Msg)]), $($message.WParam), $($message.LParam)) | Where-Object {$_})-join ', ')', initiate clean-up and exit"

                            until (
                                $($global:WatchCatchableExitSignalStatus[0] -eq 'Clean-up done')
                            ) {
                                Start-Sleep -Milliseconds 100
                            }

                            $message.Result = [IntPtr]::Zero

                            $formRef.Value.Close()
                        }
                    } catch {
                    }
                }
            )

            $formRef.Value.ShowDialog()
        }
    )
} elseif ($IsLinux -or $IsMacOS) {
    $null = $WatchCatchableExitSignalPowershell.AddScript(
        {
            # Use trap instead of try/catch, because trap reacts to catchable POSIX signals
            trap {
                $global:WatchCatchableExitSignalStatus[0] = "Detected '$($_)', initiate clean-up and exit"

                until ($global:WatchCatchableExitSignalStatus[0] -eq 'Clean-up done') {
                    Start-Sleep -Milliseconds 100
                }
            }

            while ($true) {
                Start-Sleep -Milliseconds 100
            }
        }
    )
}

$null = $WatchCatchableExitSignalPowershell.BeginInvoke()


function global:WatchCatchableExitSignal {
    param (
        [ScriptBlock]$NonExitScriptBlock = $WatchCatchableExitSignalNonExitScriptBlock,
        [switch]$CleanupDone
    )

    if ($CleanupDone) {
        if ($WatchCatchableExitSignalForm) {
            try {
                $WatchCatchableExitSignalForm.Close()
            } catch {
                # Do nothing
            }
        }

        $global:WatchCatchableExitSignalStatus[0] = 'Clean-up done'
    } elseif (
        $($global:WatchCatchableExitSignalStatus[0].StartsWith('Detected ''')) -and
        $($global:WatchCatchableExitSignalStatus[0].EndsWith(''', initiate clean-up and exit'))
    ) {
        Write-Host
        Write-Host "WatchCatchableExitSignal: $($global:WatchCatchableExitSignalStatus[0])" -ForegroundColor Yellow

        if ($WatchCatchableExitSignalForm) {
            try {
                $WatchCatchableExitSignalForm.Close()
            } catch {
                # Do nothing
            }
        }

        $script:ExitCode = 1
        $script:ExitCodeDescription = 'Detected catchable exit signal.'
        exit
    } else {
        . $NonExitScriptBlock
    }
}
#
##
### ▲▲▲ WatchCatchableExitSignal initiation code above ▲▲▲


$WatchCatchableExitSignalNonExitScriptBlock = {
    if ($script:COMWord) {
        try {
            $script:COMWord.Visible = $false
        } catch {
        }
    }

    if ($script:COMWordDummy) {
        try {
            $script:COMWordDummy.Visible = $false
        } catch {
        }
    }
}


#
# All functions have been defined above
# Initially executed code starts here
#


$script:ExitCode = 255
$script:ExitCodeDescription = 'Generic exit code, no details available. Could be because of Ctrl+C.'


try {
    try {
        $TranscriptFullName = Join-Path -Path $(Join-Path -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) -ChildPath '\Set-OutlookSignatures\Logs') -ChildPath $("Set-OutlookSignatures_Log_$(Get-Date $([DateTime]::UtcNow) -Format FileDateTimeUniversal).txt")
        $TranscriptFullName = (Start-Transcript -LiteralPath $TranscriptFullName -Force).Path

        "This folder contains log files generated by Set-OutlookSignatures.$([Environment]::NewLine)$([Environment]::NewLine)Each file is named according to the pattern 'Set-OutlookSignatures_Log_yyyyMMddTHHmmssffffZ.txt'.$([Environment]::NewLine)$([Environment]::NewLine)Files older than 14 days are automatically deleted with each execution of Set-OutlookSignatures.$([Environment]::NewLine)$([Environment]::NewLine)Ignore log lines starting with 'PS>TerminatingError' or '>> TerminatingError' unless instructed otherwise." | Out-File -LiteralPath $(Join-Path -Path (Split-Path -LiteralPath $TranscriptFullName) -ChildPath '_README.txt') -Encoding utf8 -Force
    } catch {
        $TranscriptFullName = $null
    }


    Write-Host
    Write-Host "Start Set-OutlookSignatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($TranscriptFullName) {
        Write-Host "  Log file: '$TranscriptFullName'"
        Write-Host "    Ignore log lines starting with 'PS>TerminatingError' or '>> TerminatingError' unless instructed otherwise."

        try {
            Get-ChildItem -LiteralPath $(Split-Path -LiteralPath $TranscriptFullName) -File -Force | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-14) } | ForEach-Object {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        } catch {
        }
    }

    if ($psISE) {
        Write-Host '  PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
        Write-Host '  Required features are not available in ISE. Exit.' -ForegroundColor Red
        $script:ExitCode = 2
        $script:ExitCodeDescription = 'PowerShell ISE detected.'
        exit
    }

    if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
        {
            Write-Host '' This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode."" -ForegroundColor Red
            Write-Host '  Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
            $script:ExitCode = 32
            $script:ExitCodeDescription = 'Not running in FullLanguage mode.'
            exit
        }
    }

    if ($global:SetOutlookSignaturesLastRunGuid) {
        Write-Host '  Set-OutlookSignatures has already been run before in this PowerShell session.' -ForegroundColor Yellow
        Write-Host '    Set-OutlookSignatures is allowed to run only once per session, ideally in a fresh one.' -ForegroundColor Yellow
        Write-Host '    This is the only way to avoid problems caused by .Net caching DLL files in memory.' -ForegroundColor Yellow

        $script:ExitCode = 3
        $script:ExitCodeDescription = 'Set-OutlookSignatures has already been run in this PowerShell session, is only supported once.'
        exit
    } else {
        $global:SetOutlookSignaturesLastRunGuid = (New-Guid).Guid
    }

    if (-not (Test-Path -LiteralPath 'variable:IsWindows')) {
        $script:IsWindows = $true
        $script:IsLinux = $false
        $script:IsMacOS = $false
    }

    BlockSleep

    try { WatchCatchableExitSignal } catch { }

    $OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

    if ($PSScriptRoot) {
        Set-Location -LiteralPath $PSScriptRoot
    } else {
        Write-Host 'Could not determine the script path, which is essential for this script to work.' -ForegroundColor Red
        Write-Host 'Make sure to run this script as a file from a PowerShell console, and not just as a text selection in a code editor.' -ForegroundColor Red
        Write-Host 'Exit.' -ForegroundColor Red

        $script:ExitCode = 41
        $script:ExitCodeDescription = 'Set-OutlookSignatures needs to be run as a file from a PowerShell console, and not just as a text selection in a code editor.'
        exit
    }

    $ScriptInvocation = $MyInvocation

    $script:tempDir = (New-Item -Path ([System.IO.Path]::GetTempPath()) -Name (New-Guid).Guid -ItemType Directory).FullName
    $script:ScriptRunGuid = Split-Path -Path $script:tempDir -Leaf

    $script:SetOutlookSignaturesCommonDllFilePath = (Join-Path -Path $script:tempDir -ChildPath (((New-Guid).guid) + '.dll'))
    Copy-Item -LiteralPath ((Join-Path -Path '.' -ChildPath 'bin\Set-OutlookSignatures\Set-OutlookSignatures.Common.dll')) -Destination $script:SetOutlookSignaturesCommonDllFilePath
    if (-not $IsLinux) {
        Unblock-File -LiteralPath $script:SetOutlookSignaturesCommonDllFilePath
    }

    try {
        Import-Module -Name $script:SetOutlookSignaturesCommonDllFilePath -Force -ErrorAction Stop
    } catch {
        Write-Host $error[0]
        Write-Host '    Problem importing Set-OutlookSignatures.Common.dll. Exit.' -ForegroundColor Red
        $script:ExitCode = 4
        $script:ExitCodeDescription = 'Problem importing Set-OutlookSignatures.Common.dll.'
        exit
    }

    try { WatchCatchableExitSignal } catch { }

    main

    $script:ExitCode = 0
    $script:ExitCodeDescription = 'Success.'
} catch {
    Write-Host ($error[0] | Format-List * | Out-String)
    Write-Host
    Write-Host 'Unexpected error. Exit.' -ForegroundColor red
} finally {
    Write-Host
    Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    # Restore original Word AlertIfNotDefault setting
    Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Options" -Name 'AlertIfNotDefault' -Value $script:WordAlertIfNotDefaultOriginal -ErrorAction SilentlyContinue | Out-Null

    # Restore original Word security setting
    Set-ItemProperty -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction SilentlyContinue | Out-Null

    if ($script:COMWordDummy) {
        if ($script:COMWordDummy.ActiveDocument) {
            if ($null -ne $script:COMWordShowFieldCodesOriginal) {
                try {
                    $script:COMWordDummy.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
                } catch {
                }
            }

            # Restore original WebOptions
            try {
                if ($null -ne $script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        try {
                            $script:COMWordDummy.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                        } catch {
                        }
                    }
                }
            } catch {}

            # Restore original TextEncoding
            if ($null -ne $script:WordTextEncoding) {
                try {
                    $script:COMWordDummy.ActiveDocument.TextEndocing = $script:WordTextEncoding
                } catch {
                }
            }
        }

        try {
            $script:COMWordDummy.Quit([ref]$false)
        } catch {}

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWordDummy) | Out-Null

        Remove-Variable -Name 'COMWordDummy' -Scope 'script'
    }

    if ($script:COMWord) {
        if ($script:COMWord.ActiveDocument) {
            if ($null -ne $script:COMWordShowFieldCodesOriginal) {
                try {
                    $script:COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = $script:COMWordShowFieldCodesOriginal
                } catch {
                }
            }

            if ($null -ne $script:COMWordViewTypeOriginal) {
                try {
                    $script:COMWord.ActiveDocument.ActiveWindow.View.Type = $script:COMWordViewTypeOriginal
                } catch {

                }
            }

            # Restore original WebOptions
            try {
                if ($null -ne $script:WordWebOptions) {
                    foreach ($property in @('TargetBrowser', 'BrowserLevel', 'AllowPNG', 'OptimizeForBrowser', 'RelyOnCSS', 'RelyOnVML', 'Encoding', 'OrganizeInFolder', 'PixelsPerInch', 'ScreenSize', 'UseLongFileNames')) {
                        if ($script:COMWord.ActiveDocument.WebOptions.$property -ne $script:WordWebOptions.$property) {
                            $script:COMWord.ActiveDocument.WebOptions.$property = $script:WordWebOptions.$property
                        }
                    }
                }
            } catch {}

            # Restore original TextEncoding
            if ($null -ne $script:WordTextEncoding) {
                try {
                    $script:COMWord.ActiveDocument.TextEndocing = $script:WordTextEncoding
                } catch {
                }
            }
        }

        try {
            $script:COMWord.Quit([ref]$false)
        } catch {}

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:COMWord) | Out-Null

        Remove-Variable -Name 'COMWord' -Scope 'script'
    }

    if ($script:SetOutlookSignaturesCommonDllFilePath) {
        Remove-Module -Name $([System.IO.Path]::GetFileNameWithoutExtension($script:SetOutlookSignaturesCommonDllFilePath)) -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:SetOutlookSignaturesCommonDllFilePath -Force -ErrorAction SilentlyContinue
    }

    if ($script:BenefactorCircleLicenseFilePath) {
        Remove-Module -Name $([System.IO.Path]::GetFileNameWithoutExtension($script:BenefactorCircleLicenseFilePath)) -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:BenefactorCircleLicenseFilePath -Force -ErrorAction SilentlyContinue
    }

    if ($script:WebServicesDllPath) {
        Remove-Module -Name $([System.IO.Path]::GetFileNameWithoutExtension($script:WebServicesDllPath)) -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:WebServicesDllPath -Force -ErrorAction SilentlyContinue
    }

    if ($script:MsalModulePath) {
        Remove-Module -Name MSAL.PS -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:MsalModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:PreMailerNetModulePath) {
        Remove-Item -LiteralPath $script:PreMailerNetModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:UtfUnknownModulePath) {
        Remove-Module -Name UtfUnknown -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:UtfUnknownModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:HtmlAgilityPackModulePath) {
        Remove-Module -Name HtmlAgilityPack -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:HtmlAgilityPackModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:QRCoderModulePath) {
        Remove-Module -Name QRCoder -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:QRCoderModulePath -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($script:ScriptProcessPriorityOriginal) {
        try {
            $((Get-Process -PID $PID).PriorityClass = $script:ScriptProcessPriorityOriginal)
        } catch {
        }
    }

    if ($script:SystemNetServicePointManagerSecurityProtocolOld) {
        [System.Net.ServicePointManager]::SecurityProtocol = $script:SystemNetServicePointManagerSecurityProtocolOld
    }

    if ($script:tempDir) {
        Remove-Item -LiteralPath $script:tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    if ($PSDefaultParameterValuesOriginal) {
        $PSDefaultParameterValues = $PSDefaultParameterValuesOriginal.Clone()
    }

    if ($TranscriptFullName) {
        Write-Host
        Write-Host 'Log file'
        Write-Host "  '$TranscriptFullName'"
    }

    Write-Host

    if ($script:ExitCode -eq 0) {
        Write-Host 'Exit code'
        Write-Host "  Code: $($script:ExitCode)"
        Write-Host "  Description: $($script:ExitCodeDescription)"
    } else {
        Write-Host 'Exit code' -ForegroundColor Yellow
        Write-Host "  Code: $($script:ExitCode)" -ForegroundColor Yellow
        Write-Host "  Description: $($script:ExitCodeDescription)" -ForegroundColor Yellow

        Write-Host '  Check for existing issues at https://github.com/Set-OutlookSignatures/Set-OutlookSignatures/issues?q=' -ForegroundColor Yellow
        Write-Host '  or get fee-based support from ExplicIT Consulting at https://set-outlooksignatures.com/support.' -ForegroundColor Yellow
    }

    Write-Host
    Write-Host "End Set-OutlookSignatures @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

    if ($TranscriptFullName) {
        try { Stop-Transcript | Out-Null } catch { }
    }

    # Allow sleep
    try { BlockSleep -AllowSleep } catch { }

    # Stop watching for catchable exit signals
    try { WatchCatchableExitSignal -CleanupDone } catch { }

    # End script with exit 0 or whatever is defined in $script:ExitCode
    exit $script:ExitCode
}
