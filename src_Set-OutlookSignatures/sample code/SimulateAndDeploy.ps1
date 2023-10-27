<#
This script shows how the simulation mode of Set-OutlookSignatures can be used to deploy Outlook text signatures without client involvement.

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Looking for support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support.

Features
- Automate simulation mode for all given mailboxes
- A configurable number of Set-OutlookSignatures instances run in parallel for better performance
- Set default signature in Outlook Web, no matter if classic signature or roaming signatures
- Set internal and external out of office (OOF) message
- Supports on-prem, hybrid and cloud-only environments

Requirements
- On-prem
  - the software needs to be run with an account that has a mailbox which is granted full access to all simulated mailboxes
  - If you do not want to simulate cloud mailboxes, set $ConnectOnpremInsteadOfCloud to $true
- Cloud
  - the software needs to be run with an account that has a mailbox which is granted full access to all simulated mailboxes
    - MFA is not yet supported, but script can be adapted accordingly
	  - MFA would require interactivity, breaking the possibility for complete automation
	  - Better configure a Conditional Access Policy that only allows logon from a controlled network and does not require MFA
	- Service Principals are not supported by the API
  - Create a new app registration in Entra ID/Azure AD with the following non-default properties:
    - Application permissions with admin consent
      - Microsoft Graph
        - MailboxSettings.ReadWrite # To set out of office replies for the simulated mailboxes
      - Office 365 Exchange Online
        - full_access_as_app # Required for Exchange Web Services access (read Outlook Web configuration, set classic signature and roaming signatures)
    - Delegated permissions with admin consent
      - Microsoft Graph
	    - email # Allows the app to read your users' primary email address
        - EWS.AccessAsUser.All # Allows the app to have the same access to mailboxes as the signed-in user via Exchange Web Services.
        - Group.Read.All # Allows the app to list groups, and to read their properties and all group memberships on behalf of the signed-in user. Also allows the app to read calendar, conversations, files, and other group content for all groups the signed-in user can access.
        - MailboxSettings.readwrite # Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
        - offline_access # Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.
        - openid # Allows users to sign in to the app with their work or school accounts and allows the app to see basic user profile information.
        - profile # Allows the app to see your users' basic profile (e.g., name, picture, user name, email address)
        - User.Read.All # Allows the app to read the full set of profile properties, reports, and managers of users in your organization, on behalf of the signed-in user.
    - Define a client secret (and set a reminder to update it, because it will expire)
	  - The code can easily be adapted for certificate authentication at application level (which is not possible for user authentication)
    - Set supported account types to "Accounts in this organizational directory only" (for security reasons)
  - You can limit the access of the app to specific mailboxes. This is recommended because of the "MailboxSettings.ReadWrite" and "full_access_as-app" permission required at application level.
    - Use the New-ApplicationAccessPolicy cmdlet to limit access or to deny access to specific mailboxes or to mailboxes organized in groups.
    - See https://learn.microsoft.com/en-us/powershell/module/exchange/new-applicationaccesspolicy for details
- Microsoft Word (see 'Limitations' for a scenario that does not require Word)
- File paths can get very long and be longer than the default OS limit. Make sure you allow long file paths.
  - https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation

Limitations
- Despitze parallelization, the software runtime can be unsuited for a higher number of users. The reason usually is the Word background process.
  - If you require signatures in RTF and/or TXT format, Word is needed for document conversion and you can only shorten runtime by adding hardware (scale up or scale out)
  - If you do need HTML signatures only, you can use the following workaround to avoid starting Word:
    - Use HTM templates instead of DOCX templates (parameter '-UseHtmTemplates true')
	    There are features in DOCX templates that can not be replicated HTM templates, such as applying Word specific image and text filters
    - Do not create signatures in RTF format (parameter '-CreateRtfSignatures false')
    - Do not create signatures in TXT format (parameter '-CreateTxtSignatures false')
- Roaming signatures can currently not be deployed for shared mailboxes, as the API does not support this scenario.
  - Roaming signatures for shared mailboxes pose a general problem, as only signatures with replacement variables from the $CurrentMailbox[...]$ namespace would make sense anyhow
#>


# Variables
param (
	$ConnectOnpremInsteadOfCloud = $false,
	[pscredential]$GraphUserCredential = (@(, @('SimulateAndDeployUser@example.com', 'P@ssw0rd!')) | ForEach-Object { New-Object System.Management.Automation.PSCredential ($_[0], $(ConvertTo-SecureString $_[1] -AsPlainText -Force)) }), # Use Get-Credential for interactive mode (MFA is not supported in any case)

	$GraphClientId = 'The Client Id of the Entra ID/Azure AD application for SimulateAndDeploy', # not the same id as defined in 'default graph config.ps1' or a custom Graph config file
	$GraphClientSecret = 'The Client Secret of the Entra ID/Azure AD application for SimulateAndDeploy',

	$SetOutlookSignaturesScriptPath = '..\Set-OutlookSignatures.ps1',
	$SetOutlookSignaturesScriptParameters = @{
		# Do not use: SimulateUser, SimulateMailboxes, AdditionalSignaturePath, SimulateAndDeployGraphCredentialFile
		#BenefactorCircleLicenseFile = "'\\server\share\folder\license.dll'"
		#BenefactorCircleId = '<BenefactorCircleId>'
		SimulateAndDeploy                             = $false # $false simulates but does not deploy, $true simulates and deploys
		UseHtmTemplates                               = $false
		SignatureTemplatePath                         = "'.\sample templates\Signatures DOCX'"
		SignatureIniPath                              = "'.\sample templates\Signatures DOCX\_Signatures.ini'"
		OOFTemplatePath                               = "'.\sample templates\Out of Office DOCX'"
		OOFIniPath                                    = "'.\sample templates\Out of Office DOCX\_OOF.ini'"
		ReplacementVariableConfigFile                 = "'.\config\default replacement variables.ps1'"
		GraphConfigFile                               = "'.\config\default graph config.ps1'"
		SignaturesForAutomappedAndAdditionalMailboxes = $true
		DeleteUserCreatedSignatures                   = $false
		DeleteScriptCreatedSignaturesWithoutTemplate  = $true
		SetCurrentUserOutlookWebSignature             = $true
		SetCurrentUserOOFMessage                      = $true
		MirrorLocalSignaturesToCloud                  = $true #Set to $false if you do not want to use this feature
		CreateRtfSignatures                           = $false
		CreateTxtSignatures                           = $true
		DocxHighResImageConversion                    = $true
		MoveCSSInline                                 = $true
		EmbedImagesInHtml                             = $false
		EmbedImagesInHtmlAdditionalSignaturePath      = $true
		GraphOnly                                     = $false
		TrustsToCheckForGroups                        = @('*')
		IncludeMailboxForestDomainLocalGroups         = $false
		WordProcessPriority                           = "'Normal'"
		Verbose                                       = '$true'
	},

	$SimulateResultPath = 'c:\test\SimulateAndDeploy',
	$JobsConcurrent = 2,
	$JobTimeout = [timespan]::FromMinutes(10),

	# List of users and mailboxes to simulate
	#   SimulateUser: UPN, or NT4 style NetBIOS domain name and logon name
	#   SimulateMailboxes: Separate multiple mailboxes by spaces or commas. Leave empty to get mailboxes from Outlook Web (recommended).
	#   Examples:
	#     ExampleDomain\ExampleUser;
	#     a@example.com;
	#     b@example.com;b@example.com
	#     c@example.com;c@example.com,b@example.com
	$SimulateList = (@'
SimulateUser;SimulateMailboxes
alex.alien@example.com;
bobby.busy@example.com;bobby.busy@example.com
fenix.fish@example.com;fenix.fish@example.com,nat.nuts@example.com
'@ | ConvertFrom-Csv -Delimiter ';')
)

# Functions
function CreateUpdateSimulateAndDeployGraphCredentialFile {
	# auth with user and app with delegated permissions
	$GraphClientTenantId = ($GraphUserCredential.username -split '@')[1]

	try {
		# User authentication
		$auth = get-msaltoken -UserCredential $GraphUserCredential -ClientId $GraphClientID -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://graph.microsoft.com/.default'
		$authExo = get-msaltoken -UserCredential $GraphUserCredential -ClientId $GraphClientID -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://outlook.office.com/.default'

		# App authentication
		$Appauth = get-msaltoken -ClientId $GraphClientID -ClientSecret ($GraphClientSecret | ConvertTo-SecureString -AsPlainText -Force) -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://graph.microsoft.com/.default'
		$AppauthExo = get-msaltoken -ClientId $GraphClientID -ClientSecret ($GraphClientSecret | ConvertTo-SecureString -AsPlainText -Force) -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://outlook.office.com/.default'

		$null = @{
			'AccessToken'       = $auth.accessToken
			'AuthHeader'        = $auth.createauthorizationheader()
			'AccessTokenExo'    = $authExo.accessToken
			'AuthHeaderExo'     = $authExo.createauthorizationheader()
			'AppAccessToken'    = $Appauth.accessToken
			'AppAuthHeader'     = $Appauth.createauthorizationheader()
			'AppAccessTokenExo' = $AppauthExo.accessToken
			'AppAuthHeaderExo'  = $AppauthExo.createauthorizationheader()
		} | Export-Clixml -Path $SimulateAndDeployGraphCredentialFile

		return @{
			'error' = $false
		}
	} catch {
		return @{
			'error' = $error[0] | Out-String
		}
	}
}


function RemoveItemAlternativeRecurse {
	# Function to avoid problems with OneDrive throwing "Access to the cloud file is denied"

	param(
		[alias('LiteralPath')][string] $Path,
		[switch] $SkipFolder # when $Path is a folder, do not delete $path, only it's content
	)

	$local:ToDelete = @()

	if (Test-Path -LiteralPath $path) {
		foreach ($SinglePath in @(Get-Item -LiteralPath $Path)) {
			if (Test-Path -LiteralPath $SinglePath -PathType Container) {
				if (-not $SkipFolder) {
					$local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Property PSIsContainer, @{expression = { $_.FullName.split('\').count }; descending = $true }, fullname)
					$local:ToDelete += @(Get-Item -LiteralPath $SinglePath -Force)
				} else {
					$local:ToDelete += @(Get-ChildItem -LiteralPath $SinglePath -Recurse -Force | Sort-Object -Property PSIsContainer, @{expression = { $_.FullName.split('\').count }; descending = $true }, fullname)
				}
			} elseif (Test-Path -LiteralPath $SinglePath -PathType Leaf) {
				$local:ToDelete += (Get-Item -LiteralPath $SinglePath -Force)
			}
		}
	} else {
		# Item to delete does not exist, nothing to do
	}

	foreach ($SingleItemToDelete in $local:ToDelete) {
		try {
			Remove-Item $SingleItemToDelete.FullName -Force -Recurse
		} catch {
			Write-Verbose "Could not delete $($SingleItemToDelete.FullName), error: $($_.Exception.Message)"
			Write-Verbose $_
		}
	}
}


# Start script
Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


# Folders and objects
Set-Location $PSScriptRoot | Out-Null

foreach ($VariableName in ('SimulateResultPath', 'SetOutlookSignaturesScriptPath')) {
	Set-Variable -Name $VariableName -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Get-Variable -Name $VariableName).Value).trimend('\')
}

if (-not (Test-Path $SimulateResultPath)) {
	New-Item -ItemType Directory $SimulateResultPath | Out-Null
} else {
	RemoveItemAlternativeRecurse -Path $SimulateResultPath -SkipFolder
}


# Connect to Graph
if (-not $ConnectOnpremInsteadOfCloud) {
	Write-Host "Conncect to Graph @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	Write-Host '  Microsoft Graph'
	$SimulateAndDeployGraphCredentialFile = Join-Path -Path $env:temp -ChildPath "$((New-Guid).guid).xml"

	$SetOutlookSignaturesScriptParameters['SimulateAndDeployGraphCredentialFile'] = "'$($SimulateAndDeployGraphCredentialFile)'"

	Import-Module $(Join-Path -Path (Split-Path $SetOutlookSignaturesScriptPath -Parent) -ChildPath '\bin\msal.ps') -Force

	$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

	if ($GraphConnectResult.error) {
		Start-Sleep -Seconds 10

		$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

		if ($GraphConnectResult.error) {
			Write-Host '    Exiting because of repeated Graph connection error' -ForegroundColor Red
			Write-Host "    $($GraphConnectResult.error)" -ForegroundColor Red
			exit 1
		}
	}
}


# Load and check SimulateList
Write-Host "Load info about mailboxes to simulate @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
Write-Host "  $(($SimulateList | Measure-Object).count) entries found"

$SimulateListCheckPositive = $true

$SimulateListUserDuplicate = @(@(($SimulateList | Group-Object -Property 'SimulateUser' | Where-Object { $_.count -ge 2 }).Group.SimulateUser) | Select-Object -Unique)

if ($SimulateListUserDuplicate) {
	$SimulateListCheckPositive = $false

	Write-Host '  Duplicate SimulateUser entries:' -ForegroundColor Red

	$SimulateListUserDuplicate | ForEach-Object {
		Write-Host "   $($_)" -ForegroundColor Red
	}
}

foreach ($SimulateEntry in $SimulateList) {
	if ($SimulateEntry.SimulateUser -inotmatch '^\S+@\S+$|^\S+\\\S+$') {
		$SimulateListCheckPositive = $false
		Write-Host "  Wrong format for SimulateUser: $($SimulateEntry.SimulateUser)" -ForegroundColor Red
	}

	if ($SimulateEntry.SimulateMailboxes) {
		try {
			[mailaddress[]] $tempSimulateMailboxes = @(@(($SimulateEntry.SimulateMailboxes -replace '\s+', ',' -replace ';+', ',' -replace ',+', ',') -split ',') | Where-Object { $_ })
			$SimulateEntry.SimulateMailboxes = "$($tempSimulateMailboxes -join ', ')"
		} catch {
			$SimulateListCheckPositive = $false
			Write-Host "  Wrong format for SimulateMailboxes: $($SimulateEntry.SimulateMailboxes)"
		}
	} else {
		$SimulateEntry.SimulateMailboxes = '$null'
	}
}

if (-not $SimulateListCheckPositive) {
	Write-Host

	Write-Host 'Errors found, see details above. Exiting.' -ForegroundColor Red
	exit 1
}


# Overcome Word security warning when export contains embedded pictures
# Set-OutlookSignatures handles this itself very well, but multiple instances running in the same user account may lead to problems
# As a workaround, we define the setting before running the jobs
Write-Host "Export Word security setting and disable it @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
$WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Word.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
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

$WordDisableWarningOnIncludeFieldsUpdate = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore

if (($null -eq $WordDisableWarningOnIncludeFieldsUpdate) -or ($WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
	New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -PropertyType DWord -Value 1 -ErrorAction Ignore | Out-Null
	Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value 1 -ErrorAction Ignore | Out-Null
}


# Run simulation mode for each user
Write-Host "Run simulation mode for each user and its mailbox(es) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

Write-Host '  Remove old jobs'
Get-Job | Remove-Job -Force

$JobsToStartTotal = ($SimulateList | Measure-Object).count
$JobsToStartOpen = ($SimulateList | Measure-Object).count
$JobsStarted = 0
$JobsCompleted = 0

Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"

do {
	while ((($JobsToStartOpen -gt 0) -and ((Get-Job).count -lt $JobsConcurrent))) {
		$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$Jobsstarted].SimulateUser)) -ChildPath '_log.txt'

		if ((Test-Path (Split-Path $LogFilePath -Parent)) -eq $false) {
			New-Item -ItemType Directory -Path (Split-Path $LogFilePath -Parent) | Out-Null
		}

		# Update Graph credential file before starting a job
		#   this makes sure that the token is still valid when the software runs longer than token lifetime
		if (-not $ConnectOnpremInsteadOfCloud) {
			$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

			if ($GraphConnectResult.error) {
				Start-Sleep -Seconds 10

				$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

				if ($GraphConnectResult.error) {
					$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

					if ($GraphConnectResult.error) {
						Start-Sleep -Seconds 30

						$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

						if ($GraphConnectResult.error) {
							Write-Host '    Exiting because of repeated Graph connection error' -ForegroundColor Red
							Write-Host "    $($GraphConnectResult.error)" -ForegroundColor Red
							exit 1
						}
					}
				}
			}
		}

		Start-Job {
			Param (
				$PowershellPath,
				$SetOutlookSignaturesScriptPath,
				[string]$SimulateUser,
				[string]$SimulateMailboxes,
				$SimulateResultPath,
				$LogFilePath,
				$SetOutlookSignaturesScriptParameters,
				$SimulateAndDeployGraphCredentialFile
			)

			Start-Transcript -LiteralPath $LogFilePath -Force

			try {
				Write-Host 'CREATE SIGNATURE FILES BY USING SIMULATON MODE OF SET-OUTLOOKSIGNATURES'

				$SetOutlookSignaturesScriptParameters['SimulateUser'] = $SimulateUser
				$SetOutlookSignaturesScriptParameters['SimulateMailboxes'] = $SimulateMailboxes
				$SetOutlookSignaturesScriptParameters['AdditionalSignaturePath'] = "'$(Join-Path -Path $SimulateResultPath -ChildPath $SimulateUser)'"

				$InvokeCode = "& '$SetOutlookSignaturesScriptPath' $([string]::Join(' ', ($SetOutlookSignaturesScriptParameters.GetEnumerator() | ForEach-Object { "-$($_.Key):$($_.Value)" })))"

				Write-Host $InvokeCode

				Invoke-Expression $InvokeCode

				if ($?) {
					Write-Host 'xxxSimulateAndDeployExitCode0xxx'
				} else {
					Write-Host 'xxxSimulateAndDeployExitCode999xxx'
				}
			} catch {
				$error[0]
				Write-Host 'xxxSimulateAndDeployExitCode999xxx'
			}

			Stop-Transcript
		} -Name ("$($Jobsstarted)_Job") -ArgumentList (Get-Process -Id $pid).Path,
		$SetOutlookSignaturesScriptPath,
		$($SimulateList[$Jobsstarted].SimulateUser),
		$($SimulateList[$Jobsstarted].SimulateMailboxes),
		$SimulateResultPath,
		$LogFilePath,
		$SetOutlookSignaturesScriptParameters,
		$SimulateAndDeployGraphCredentialFile | Out-Null

		Write-Host "    User $($SimulateList[$Jobsstarted].SimulateUser) started @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

		$JobsToStartOpen--
		$JobsStarted++

		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"
	}

	foreach ($x in (Get-Job | Where-Object { $_.State -ieq 'Running' -and (((Get-Date) - $_.PSBeginTime) -gt $JobTimeout) })) {
		Write-Host "    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) canceled due to timeout @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" -ForegroundColor Red

		$x | Remove-Job -Force

		$JobsCompleted++

		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"
	}

	foreach ($x in (Get-Job -State Completed)) {
		$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser)) -ChildPath '_log.txt'

		if (-not (Get-Content -Path $LogFilePath -Encoding UTF8 -Raw).trim().Contains('xxxSimulateAndDeployExitCode0xxx')) {
			Write-Host "    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) ended @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
			Write-Host "      User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser): Error creating signatures, please check log." -ForegroundColor Red
		}

		$x | Remove-Job -Force

		$JobsCompleted++

		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"
	}

} until (($JobsToStartTotal -eq $JobsStarted) -and ($JobsCompleted -eq $JobsToStartTotal))


# Restore Word security setting for embedded images
Write-Host "Restore original Word security setting @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if ($null -eq $WordDisableWarningOnIncludeFieldsUpdate) {
	Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
} else {
	Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
}


Write-Host "Cleanup @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
if (-not $ConnectOnpremInsteadOfCloud) {
	Remove-Module msal.ps
	Remove-Item -Force $SimulateAndDeployGraphCredentialFile -ErrorAction SilentlyContinue
}


Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
