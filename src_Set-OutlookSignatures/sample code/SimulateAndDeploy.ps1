<#
This sample code shows how to achieve two thing:
	- Running simulation mode for multiple users
  	- How to use simulation mode together with the Benefactor Circle add-on to push signatures and out-of-office replies into mailboxes,
	  without involving end users or their devices

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.


Features
	- Automate simulation mode for all given mailboxes
	- A configurable number of Set-OutlookSignatures instances run in parallel for better performance
	- Set default signature in Outlook Web, no matter if classic signature or roaming signatures (requires the Benefacot Circle add-on)
	- Set internal and external out of office (OOF) message (requires the Benefacot Circle add-on)
	- Supports on-prem, hybrid and cloud-only environments


Requirements
  	Follow the requirements exactly and in full. SimulateAndDeploy will not work correctly when even one requirement is not met.
	Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.

	- For mailboxes on-prem
		- The software needs to be run with an account that
			- has a mailbox
			- and is granted "full access" to all simulated mailboxes
		- If you do not want to simulate cloud mailboxes, set $ConnectOnpremInsteadOfCloud to $true
	- For mailboxes in Exchange Online
		- The software needs to be run with an account that
			- has a mailbox
			- and is granted "full access" to all simulated mailboxes
		- MFA is not yet supported, but script can be adapted accordingly
			- MFA would require interactivity, breaking the possibility for complete automation
			- Better configure a Conditional Access Policy that only allows logon from a controlled network and does not require MFA
		- Service Principals are not supported by the API
		- Create a new app registration in Entra ID/Azure AD
			- Option A: Create the app automatically by using the script '.\sample code\Create-EntraApp.ps1'
				The sample code creates the app with all required settings automatically, only providing admin consent is a manual task
			- Option B: Create the Entra app manually, with the following properties:
				- Application (!) permissions with admin consent
					- Microsoft Graph
						- GroupMember.Read.All
							Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
							Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
						- MailboxSettings.ReadWrite
							Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
							Required to set out-of-office replies for the simulated mailboxes
						- User.Read.All
							Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
							Required for $CurrentUser[...]$ and $CurrentMailbox[...]$ replacement variables, and for simulation mode.
					- Office 365 Exchange Online
						- full_access_as_app
							Allows the app to have full access via Exchange Web Services to all mailboxes without a signed-in user.
							Required for Exchange Web Services access (read Outlook Web configuration, set classic signature and roaming signatures)
				- Delegated (!) permissions with admin consent
					These permissions equal those mentioned in '.\config\default graph config.ps1'
					- Microsoft Graph
						- email
							Allows the app to read your users' primary email address.
							Required to log on the current user.
						- EWS.AccessAsUser.All
							Allows the app to have the same access to mailboxes as the signed-in user via Exchange Web Services.
							Required to connect to Outlook Web and to set Outlook Web signature (classic and roaming).
						- GroupMember.Read.All
							Allows the app to list groups, read basic group properties and read membership of all groups the signed-in user has access to.
							Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
						- MailboxSettings.ReadWrite
							Allows the app to create, read, update, and delete user's mailbox settings. Does not include permission to send mail.
							Required to detect the state of the out of office assistant and to set out-of-office replies.
						- offline_access
							Allows the app to see and update the data you gave it access to, even when users are not currently using the app. This does not give the app any additional permissions.
							Required to get a refresh token from Graph.
						- openid
							Allows users to sign in to the app with their work or school accounts and allows the app to see basic user profile information.
							Required to log on the current user.
						- profile
							Allows the app to see your users' basic profile (e.g., name, picture, user name, email address).
							Required to log on the current user, to access the '/me' Graph API, to get basic properties of the current user.
						- User.Read.All
							Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user.
							Required for $CurrentUser[...]$ and $CurrentMailbox[...]$ replacement variables, and for simulation mode.
					- Define a client secret (and set a reminder to update it, because it will expire)
						The code can easily be adapted for certificate authentication at application level (which is not possible for user authentication)
					- Set supported account types to "Accounts in this organizational directory only" (for security reasons)
		- You can limit the access of the app to specific mailboxes. This is recommended because of the "MailboxSettings.ReadWrite" and "full_access_as-app" permission required at application level.
			- Use the New-ApplicationAccessPolicy cmdlet to limit access or to deny access to specific mailboxes or to mailboxes organized in groups.
			- See https://learn.microsoft.com/en-us/powershell/module/exchange/new-applicationaccesspolicy for details
	- Microsoft Word (see 'Limitations' for a scenario that does not require Word)
	- File paths can get very long and be longer than the default OS limit. Make sure you allow long file paths.
		- https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation
	- Do not forget to adapt the "Variables" section of this script according to your needs and your configuration


Limitations
	- Despitze parallelization, the execution time can be too long for a higher number of users. The reason usually is the Word background process.
		- If you use DOCX templates and/or require signatures in RTF format, Word is needed for document conversion and you can only shorten runtime by adding hardware (scale up or scale out)
		- If you do need HTML signatures only, you can use the following workaround to avoid starting Word:
			- Use HTM templates instead of DOCX templates (parameter '-UseHtmTemplates true')
				There are features in DOCX templates that can not be replicated HTM templates, such as applying Word specific image and text filters
			- Do not create signatures in RTF format (parameter '-CreateRtfSignatures false')
	- Roaming signatures can currently not be deployed for shared mailboxes, as the API does not support this scenario.
		- Roaming signatures for shared mailboxes pose a general problem, as only signatures with replacement variables from the $CurrentMailbox[...]$ namespace would make sense anyhow
#>

[CmdletBinding()]

# Variables
param (
	$ConnectOnpremInsteadOfCloud = $false,
	[pscredential]$GraphUserCredential = (@(, @('SimulateAndDeployUser@example.com', 'P@ssw0rd!')) | ForEach-Object { New-Object System.Management.Automation.PSCredential ($_[0], $(ConvertTo-SecureString $_[1] -AsPlainText -Force)) }), # Use Get-Credential for interactive mode (MFA is not supported in any case)

	$GraphClientId = 'The Client Id of the Entra ID/Azure AD application for SimulateAndDeploy', # not the same id as defined in 'default graph config.ps1' or a custom Graph config file
	$GraphClientSecret = 'The Client Secret of the Entra ID/Azure AD application for SimulateAndDeploy',

	[ValidateSet('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC', 'AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4', 'AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5', 'China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud')]
	[string]$CloudEnvironment = 'Public',

	$SetOutlookSignaturesScriptPath = '..\Set-OutlookSignatures.ps1',
	$SetOutlookSignaturesScriptParameters = @{
		# Do not use: SimulateUser, SimulateMailboxes, AdditionalSignaturePath, SimulateAndDeployGraphCredentialFile
		#
		# ▼▼▼▼ The "Deploy" part of "SimulateAndDeploy" requires a Benefactor Circle license ▼▼▼▼
		# ▼▼▼▼ Without the license, signatures can not be read from and written to mailboxes ▼▼▼▼
		BenefactorCircleLicenseFile                   = "'\\server\share\folder\license.dll'"
		BenefactorCircleId                            = '<BenefactorCircleId>'
		# ▲▲▲▲ The "Deploy" part of "SimulateAndDeploy" requires a Benefactor Circle license ▲▲▲▲
		# ▲▲▲▲ Without the license, signatures can not be read from and written to mailboxes ▲▲▲▲
		#
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
		MirrorCloudSignatures                  = $true #Set to $false if you do not want to use this feature
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
	},

	$SimulateResultPath = 'c:\test\SimulateAndDeploy',
	$JobsConcurrent = 2,
	$JobTimeout = [timespan]::FromMinutes(10),

	$UpdateInterval = [timespan]::FromMinutes(1),

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
		$auth = get-msaltoken -AzureCloudInstance $CloudEnvironmentEnvironmentName -UserCredential $GraphUserCredential -ClientId $GraphClientID -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://graph.microsoft.com/.default'
		$authExo = get-msaltoken -AzureCloudInstance $CloudEnvironmentEnvironmentName -UserCredential $GraphUserCredential -ClientId $GraphClientID -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://outlook.office.com/.default'

		# App authentication
		$Appauth = get-msaltoken -AzureCloudInstance $CloudEnvironmentEnvironmentName -ClientId $GraphClientID -ClientSecret ($GraphClientSecret | ConvertTo-SecureString -AsPlainText -Force) -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://graph.microsoft.com/.default'
		$AppauthExo = get-msaltoken -AzureCloudInstance $CloudEnvironmentEnvironmentName -ClientId $GraphClientID -ClientSecret ($GraphClientSecret | ConvertTo-SecureString -AsPlainText -Force) -TenantId $GraphClientTenantId -RedirectUri 'http://localhost' -Scopes 'https://outlook.office.com/.default'

		$null = @{
			'AccessToken'       = $auth.AccessToken
			'AuthHeader'        = $auth.createauthorizationheader()
			'AccessTokenExo'    = $authExo.AccessToken
			'AuthHeaderExo'     = $authExo.createauthorizationheader()
			'AppAccessToken'    = $Appauth.AccessToken
			'AppAuthHeader'     = $Appauth.createauthorizationheader()
			'AppAccessTokenExo' = $AppauthExo.AccessToken
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
$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

# Cloud environment
## Endpoints from https://github.com/microsoft/CSS-Exchange/blob/main/Shared/AzureFunctions/Get-CloudServiceEndpoint.ps1
## Environment names must match https://learn.microsoft.com/en-us/dotnet/api/microsoft.identity.client.azurecloudinstance?view=msal-dotnet-latest
switch ($CloudEnvironment) {
	{ $_ -iin @('Public', 'Global', 'AzurePublic', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC') } {
		$CloudEnvironmentEnvironmentName = 'AzurePublic'
		$CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.com'
		$CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office.com'
		$CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.outlook.com'
		$CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.com'
		break
	}

	{ $_ -iin @('AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4') } {
		$CloudEnvironmentEnvironmentName = 'AzureUSGovernment'
		$CloudEnvironmentGraphApiEndpoint = 'https://graph.microsoft.us'
		$CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook.office365.us'
		$CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.office365.us'
		$CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.us'
		break
	}

	{ $_ -iin @('AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5') } {
		$CloudEnvironmentEnvironmentName = 'AzureUSGovernment'
		$CloudEnvironmentGraphApiEndpoint = 'https://dod-graph.microsoft.us'
		$CloudEnvironmentExchangeOnlineEndpoint = 'https://outlook-dod.office365.us'
		$CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s-dod.office365.us'
		$CloudEnvironmentAzureADEndpoint = 'https://login.microsoftonline.us'
		break
	}

	{ $_ -iin @('China', 'AzureChina', 'ChinaCloud', 'AzureChinaCloud') } {
		$CloudEnvironmentEnvironmentName = 'AzureChina'
		$CloudEnvironmentGraphApiEndpoint = 'https://microsoftgraph.chinacloudapi.cn'
		$CloudEnvironmentExchangeOnlineEndpoint = 'https://partner.outlook.cn'
		$CloudEnvironmentAutodiscoverSecureName = 'https://autodiscover-s.partner.outlook.cn'
		$CloudEnvironmentAzureADEndpoint = 'https://login.partner.microsoftonline.cn'
		break
	}
}

$SetOutlookSignaturesScriptParameters.CloudEnvironment = $CloudEnvironment # not $CloudEnvironmentEnvironmentName


foreach ($VariableName in ('SimulateResultPath', 'SetOutlookSignaturesScriptPath')) {
	Set-Variable -Name $VariableName -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Get-Variable -Name $VariableName).Value).trimend('\')
}

if (-not (Test-Path $SimulateResultPath)) {
	New-Item -ItemType Directory $SimulateResultPath | Out-Null
} else {
	RemoveItemAlternativeRecurse -Path $SimulateResultPath -SkipFolder
}

@(
	(Join-Path -Path $SimulateResultPath -ChildPath '_log_started.txt'),
	(Join-Path -Path $SimulateResultPath -ChildPath '_log_success.txt'),
	(Join-Path -Path $SimulateResultPath -ChildPath '_log_error.txt')
) | ForEach-Object {
	New-Item -ItemType File $_ | Out-Null
}


Start-Transcript -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log.txt') -Force


# Connect to Graph
if (-not $ConnectOnpremInsteadOfCloud) {
	Write-Host "Connect to Graph @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	Write-Host '  Microsoft Graph'
	$SimulateAndDeployGraphCredentialFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "$((New-Guid).guid).xml"

	$SetOutlookSignaturesScriptParameters['SimulateAndDeployGraphCredentialFile'] = "'$($SimulateAndDeployGraphCredentialFile)'"

	Import-Module $(Join-Path -Path (Split-Path $SetOutlookSignaturesScriptPath -Parent) -ChildPath '\bin\MSAL.PS') -Force

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
if (($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) -and ($SetOutlookSignaturesScriptParameters.UseHtmTemplates -inotin (1, '1', 'true', '$true', 'yes'))) {
	Write-Host "Export Word security setting and disable it @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
	$WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Word.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
	if ($WordRegistryVersion.major -gt 16) {
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
}

# Run simulation mode for each user
Write-Host "Run simulation mode for each user and its mailbox(es) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

Write-Host '  Remove old jobs'

Get-Job | Remove-Job -Force

$JobsToStartTotal = ($SimulateList | Measure-Object).count
$JobsToStartOpen = ($SimulateList | Measure-Object).count
$JobsStarted = 0
$JobsCompleted = 0

Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


$UpdateTime = (Get-Date).Add($UpdateInterval)

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

		"    User $($SimulateList[$Jobsstarted].SimulateUser) started @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
			Write-Host $($_)
			Add-Content -Value $($_) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_started.txt') -Force -Encoding UTF8
		}

		$JobsToStartOpen--
		$JobsStarted++

		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
	}

	foreach ($x in (Get-Job | Where-Object { $_.State -ieq 'Running' -and (((Get-Date) - $_.PSBeginTime) -gt $JobTimeout) })) {
		"    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) canceled due to timeout @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
			Write-Host $($_) -ForegroundColor Red
			Add-Content -Value $($_) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_error.txt') -Force -Encoding UTF8
		}

		$x | Remove-Job -Force

		$JobsCompleted++

		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
	}

	foreach ($x in (Get-Job -State Completed)) {
		$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser)) -ChildPath '_log.txt'

		if ((Get-Content -Path $LogFilePath -Encoding UTF8 -Raw).trim().Contains('xxxSimulateAndDeployExitCode0xxx')) {
			"    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) ended with no errors @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
				Write-Host $($_) -ForegroundColor Green
				Add-Content -Value $($_) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_success.txt') -Force -Encoding UTF8
			}
		} else {
			"    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) ended with errors @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
				Write-Host $($_) -ForegroundColor Red
				Add-Content -Value $($_) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_error.txt') -Force -Encoding UTF8
			}
		}

		$x | Remove-Job -Force

		$JobsCompleted++

		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
	}

	if ((Get-Date) -ge $UpdateTime) {
		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
		$UpdateTime = (Get-Date).Add($UpdateInterval)
	}

	Start-Sleep -Seconds 1
} until (($JobsToStartTotal -eq $JobsStarted) -and ($JobsCompleted -eq $JobsToStartTotal))


# Restore Word security setting for embedded images
if (($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) -and ($SetOutlookSignaturesScriptParameters.UseHtmTemplates -inotin (1, '1', 'true', '$true', 'yes'))) {
	Write-Host "Restore original Word security setting @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	if ($null -eq $WordDisableWarningOnIncludeFieldsUpdate) {
		Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
	} else {
		Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
	}
}

Write-Host "Cleanup @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

if (-not $ConnectOnpremInsteadOfCloud) {
	Remove-Module MSAL.PS
	Remove-Item -Force $SimulateAndDeployGraphCredentialFile -ErrorAction SilentlyContinue
}


Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


Stop-Transcript
