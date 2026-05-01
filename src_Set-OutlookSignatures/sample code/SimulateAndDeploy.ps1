<#
This sample code shows how to achieve two things:
  - Running simulation mode for multiple users
   How to use simulation mode together with the Benefactor Circle add-on to push signatures and out-of-office replies into mailboxes, without involving end users or their devices

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.


Features
  - Automate simulation mode for all given mailboxes
    - SimulateAndDeploy considers additional mailboxes when the user added them in Outlook for the web, when they are passed via the 'SimulateMailboxes' parameter, or when being added dynamically via the 'VirtualMailboxConfigFile' parameter
  - A configurable number of Set-OutlookSignatures instances run in parallel for better performance
  - Set default signature in Outlook for the web, no matter if classic signature or roaming signatures (requires the Benefactor Circle add-on)
  - Set internal and external out-of-office (OOF) message (requires the Benefactor Circle add-on)
  - Supports on-prem, hybrid and cloud-only environments


Requirements
  Follow the requirements exactly and in full. SimulateAndDeploy will not work correctly when even one requirement is not met.
  Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.

  - For mailboxes on-prem
      - The software needs to be run with an account that
          - has a mailbox
          - and is granted "full access" to all simulated mailboxes
      - If you do not want to simulate cloud mailboxes, set $ConnectOnpremInsteadOfCloud to $true
  - For mailboxes in Exchange Online
      - The software needs to be run with an account that
          - has a mailbox
          - and is granted "full access" to all simulated mailboxes
      - MFA is not supported
          - MFA would require interactivity, breaking the possibility for complete automation
          - Better configure a Conditional Access Policy that only allows logon from a controlled network and does not require MFA
      - Service Principals are not supported by the API
      - Create a new app registration in Entra ID
          - Option A: Create the app automatically by using the script '.\sample code\Create-EntraApp.ps1'
		      - The sample code creates the app with all required settings automatically, only providing admin consent is a manual task
		  - Option B: Create the Entra ID app manually, with the following properties:
		      - Application (!) permissions with admin consent
                  - Microsoft Graph
				      - Files.Read.All
					    https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadall
					    Read template and configuration files hosted on SharePoint Online. Alternative: Files.SelectedOperations.Selected.
					  - GroupMember.Read.All
						https://learn.microsoft.com/en-us/graph/permissions-reference#groupmemberreadall
						Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
					  - Mail.ReadWrite
					    https://learn.microsoft.com/en-us/graph/permissions-reference#mailreadwrite
						Create signature collection in drafts, provide signatures for Outlook add-in.
					  - MailboxSettings.ReadWrite
						https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsreadwrite
						Detect mailbox environment, get and set out-of-office data.
					  - User.Read.All
						https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall
						Data for replacement variables, SMTP to UPN, group membership.
					  - MailboxConfigItem.ReadWrite
					    https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxconfigitemreadwrite
						Read data from Outlook Web, set Outlook web signatures.
			  - Delegated (!) permissions with admin consent
				These permissions equal those mentioned in '.\config\default graph config.ps1'
			      - Microsoft Graph
					  - email
				        https://learn.microsoft.com/en-us/graph/permissions-reference#email
					    Authenticate the signed-in user.
					  - MailboxConfigItem.ReadWrite
						https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxconfigitemreadwrite
						Read data from Outlook Web, set Outlook web signatures.
					  - Files.Read.All
					    https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadall
					    Read template and configuration files hosted on SharePoint Online. Alternative: Files.SelectedOperations.Selected.
					  - GroupMember.Read.All
						https://learn.microsoft.com/en-us/graph/permissions-reference#groupmemberreadall
						Required to find groups by name and to get their security identifier (SID) and the number of transitive members.
					  - Mail.ReadWrite
						https://learn.microsoft.com/en-us/graph/permissions-reference#mailreadwrite
						Create signature collection in drafts, provide signatures for Outlook add-in.
					  - MailboxSettings.ReadWrite
						https://learn.microsoft.com/en-us/graph/permissions-reference#mailboxsettingsreadwrite
						Detect mailbox environment, get and set out-of-office data.
					  - offline_access
						https://learn.microsoft.com/en-us/graph/permissions-reference#offline_access
						Required to get a refresh token from Graph.
					  - openid
						https://learn.microsoft.com/en-us/graph/permissions-reference#openid
						Authenticate the signed-in user.
					  - profile
						https://learn.microsoft.com/en-us/graph/permissions-reference#profile
						Authenticate the signed-in user, get basic properties.
					  - User.Read.All
						https://learn.microsoft.com/en-us/graph/permissions-reference#userreadall
						Data for replacement variables, SMTP to UPN, group membership.
			  - Define a client secret (and set a reminder to update it, because it will expire)
				The code can easily be adapted for certificate authentication at application level (which is not possible for user authentication)
			  - Set supported account types to "Accounts in this organizational directory only" (for security reasons)
		  - You can limit the access of the app to specific mailboxes. This is recommended because of the "MailboxSettings.ReadWrite" and "full_access_as-app" permission required at application level.
		      - Use the New-ApplicationAccessPolicy cmdlet to limit access or to deny access to specific mailboxes or to mailboxes organized in groups.
			  - See https://learn.microsoft.com/en-us/powershell/module/exchange/new-applicationaccesspolicy for details
  - Microsoft Word when using DOCX tempates
  - File paths can get very long and be longer than the default OS limit. Make sure you allow long file paths.
      - https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation
  - Do not forget to adapt the "Variables" section of this script according to your needs and your configuration


Limitations and remarks
  - Despitze parallelization, the execution time can be too long for a higher number of users. The reason usually is the Word background process.
      - If you use DOCX templates and/or require signatures in RTF format, Word is needed for document conversion and you can only shorten runtime by adding hardware (scale up or scale out)
	  - If you do need HTML signatures only, you can use the following workaround to avoid starting Word:
	      - Use HTM templates instead of DOCX templates (parameter '-UseHtmTemplates true')
			There are features in DOCX templates that cannot be replicated HTM templates, such as applying Word specific image and text filters
		  - Do not create signatures in RTF format (parameter '-CreateRtfSignatures false')
  - Roaming signatures can currently not be deployed for shared mailboxes, as the API does not support this scenario.
      - Roaming signatures for shared mailboxes pose a general problem, as only signatures with replacement variables from the $CurrentMailbox[…]$ namespace would make sense anyhow
  - SimulateAndDeploy cannot solve problems around the Classic Outlook for Windows roaming signature sync engine, only Microsoft can do this (but unfortunately does not since years).
      - Until Microsoft solves this in Classic Outlook for Windows, expect problems with character encoding (umlauts, diacritics, emojis, etc.) and more.
	  - These Outlook-internal problems will come and go depending on the patch level of Outlook.
	  - These Outlook-internal problems can also be observed when Set-OutlookSignatures is not involved at all.
	  - The only workaround currently known is to disable the Classic Outlook for Windows sync engine and let Set-OutlookSignatures do it by running it on the client regularly.
  - Signatures are directly usable in Outlook for the web and New Outlook (when based on Outlook for the web). Other Outlook editions may work but are not supported.
      - Consider using the Outlook add-in to access signatures created by SimulateAndDeploy on other editions of Outlook in a supported way.
	    See https://set-outlooksignatures.com/outlookaddin for details.
	  - Also see FAQ 'Roaming signatures in Classic Outlook for Windows look different' at https://set-outlooksignatures.com/faq.
  - Consider using the 'VirtualMailboxConfigFile' parameter of Set-OutlookSignatures, ideally together with the output of the Export-RecipientPermissions script.
      - This allows you to automatically create up-to-date lists of mailboxes based on the permissions granted in Exchange, as well as the according INI file lines.
	  - Visit https://github.com/Export-RecipientPermissions for details about Export-RecipientPermissions.
  - Some Word builds throw an error message when run in a non-interactive mode (such as using a scheduled task configured with "Run whether user is logged on or not").
      - The only known workarounds are to run SimulateAndDeploy in interactive mode, or to use HTM templates instead of DOCX templates.
	    ExplicIT Consulting can help you create code converting DOCX templates to HTM templates automatically.


It is recommended to not modify or copy this sample script, but to call it with parameters.
  - The "param" section at the beginning of the script defines all parameters that can be used to call this script.
#>


# Suppress specific PSScriptAnalyzerRules for specific variables
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'CloudEnvironment')]


#Requires -Version 5.1


[CmdletBinding()]

# Variables
param (
	$ConnectOnpremInsteadOfCloud = $false,

	# $GraphUserCredential, $GraphClientID and $GraphClientSecret are only there for backward compatibility.
	# It is recommended to use $GraphData instead. In cross-tenant and multitenant scenarios this is a must.
	[pscredential]$GraphUserCredential = (
		@(, @('SimulateAndDeployUser@example.com', 'P@ssw0rd!')) | ForEach-Object { New-Object System.Management.Automation.PSCredential ($_[0], $(ConvertTo-SecureString $_[1] -AsPlainText -Force)) }
	), # Use Get-Credential for interactive mode or (Get-Content -LiteralPath '.\Config\password.secret') to retrieve info from a separate file (MFA is not supported in any case)
	$GraphClientID = '<The application (client) ID of the Entra ID app for SimulateAndDeploy>', # not the same ID as defined in 'default graph config.ps1' or a custom Graph config file
	$GraphClientSecret = '<The client secret of the Entra ID app for SimulateAndDeploy>', # to load the secret from a file, use (Get-Content -LiteralPath '.\Config\app.secret')
	[ValidateNotNullOrEmpty()]
	[string]$CloudEnvironment = 'Public',

	# Define custom cloud environments, such as for not yet publicly documented sovereign clouds
	$CustomCloudEnvironments = @(
		# @{
		# 	Aliases                 = @('AzureExample', 'Example', 'ExampleCloud', 'AzureExampleCloud') # Mandatory. Each value must be unique across all environments.
		# 	AzureADEndpoint         = 'https://login.sovcloud-identity.example' # Mandatory.
		# 	GraphApiEndpoint        = 'https://graph.svc.sovcloud.example' # Mandatory.
		# 	ExchangeOnlineEndpoint  = 'https://outlook.sovcloud.example' # Mandatory.
		# 	AutodiscoverSecureName  = 'https://autodiscover-s.outlook.sovcloud.example' # Mandatory.
		# 	SharePointOnlineDomains = @('sharepoint.example') # Mandatory for accessing SharePoint via Graph.
		# }

		# @{
		#   ...
		# }
	),

	# As soon as $GraphData is not just an empty array, it takes priority over $GraphUserCredential, $GraphClientID, and $GraphClientSecret
	# Format:
	# $GraphData = @(
	#     , @('Tenant A ID', 'Tenant A SimulateAndDeployUser', 'Tenant A SimulateAndDeployUserPassword', 'Tenant A GraphClientID', 'Tenant A GraphClientSecret')
	#     , @('Tenant B ID', 'Tenant B SimulateAndDeployUser', 'Tenant B SimulateAndDeployUserPassword', 'Tenant B GraphClientID', 'Tenant B GraphClientSecret')
	# )
	$GraphData = @(),

	$SetOutlookSignaturesScriptPath = '..\Set-OutlookSignatures.ps1',
	$SetOutlookSignaturesScriptParameters = @{
		# Do not use: SimulateUser, SimulateMailboxes, AdditionalSignaturePath, SimulateAndDeployGraphCredentialFile
		#
		# The "Deploy" part of "SimulateAndDeploy" requires a Benefactor Circle license
		# Without the license, signatures cannot be read from and written to mailboxes
		BenefactorCircleLicenseFile   = '\\server\share\folder\license.dll'
		BenefactorCircleID            = '<BenefactorCircleID>'
		# The "Deploy" part of "SimulateAndDeploy" requires a Benefactor Circle license
		# Without the license, signatures cannot be read from and written to mailboxes
		#
		SimulateAndDeploy             = $false # $false simulates but does not deploy, $true simulates and deploys
		UseHtmTemplates               = $false
		SignatureTemplatePath         = '.\sample templates\Signatures DOCX'
		SignatureIniFile              = '.\sample templates\Signatures DOCX\_Signatures.ini'
		OOFTemplatePath               = '.\sample templates\Out-of-Office DOCX'
		OOFIniFile                    = '.\sample templates\Out-of-Office DOCX\_OOF.ini'
		ReplacementVariableConfigFile = '.\config\default replacement variables.ps1'
		GraphConfigFile               = '.\config\default graph config.ps1'
		GraphOnly                     = $false
		# Use current verbose mode for later execution of Set-OutlookSignatures
		Verbose                       = $($VerbosePreference -ne [System.Management.Automation.ActionPreference]::SilentlyContinue)
	},

	$SimulateResultPath = 'c:\test\SimulateAndDeploy',
	$JobsConcurrent = 2,
	$JobTimeout = [timespan]::FromMinutes(10),

	$UpdateInterval = [timespan]::FromMinutes(1),

	# List of users and mailboxes to simulate
	#   SimulateUser: Logon name in UPN or pre-Windows 2000 format
	#   SimulateMailboxes: Separate multiple mailboxes by spaces or commas. Leave empty to get mailboxes from Outlook for the web (recommended).
	#   Examples:
	#     ExampleDomain\ExampleUser;
	#     a@example.com;
	#     b@example.com;b@example.com
	#     c@example.com;c@example.com,b@example.com
	#   Consider using the 'VirtualMailboxConfigFile' parameter of Set-OutlookSignatures, ideally together with the output of the Export-RecipientPermissions script.
	#     This allows you to automatically create up-to-date lists of mailboxes based on the permissions granted in Exchange, as well as the according INI file lines.
	#     Visit https://github.com/Export-RecipientPermissions for details about Export-RecipientPermissions.
	$SimulateList = (@'
SimulateUser;SimulateMailboxes
alex.alien@example.com;
bobby.busy@example.com;bobby.busy@example.com
fenix.fish@example.com;fenix.fish@example.com,nat.nuts@example.com
'@ | ConvertFrom-Csv -Delimiter ';')
)


#
# Do not change anything from here on
#


# Functions
function ParseJwtToken {
	# Idea for this code: https://www.michev.info/blog/post/2140/decode-jwt-access-and-id-tokens-via-powershell

	[cmdletbinding()]
	param([Parameter(Mandatory = $true)][string]$token)

	try { global:WatchCatchableExitSignal } catch {}

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


function CreateUpdateSimulateAndDeployGraphCredentialFile {
	$local:GraphTokenDictionary = @{}
	$returnValuesCollection = @()

	foreach ($GraphDataObject in $GraphData) {
		$local:GraphTenantId = GraphDomainToTenantID $GraphDataObject[0]
		$local:GraphUserCredentialUser = $GraphDataObject[1]
		$local:GraphUserCredentialPassword = $GraphDataObject[2]
		$local:GraphClientId = $GraphDataObject[3]
		$local:GraphClientSecret = $GraphDataObject[4]

		$local:GraphUserCredential = (New-Object System.Management.Automation.PSCredential ($local:GraphUserCredentialUser, $(ConvertTo-SecureString $local:GraphUserCredentialPassword -AsPlainText -Force)))

		try {
			# User authentication
			$auth = get-msaltoken -Authority "$($script:CloudEnvironmentAzureADEndpoint)/$(@($local:GraphTenantId, 'organizations') | Where-Object { $_ } | Select-Object -First 1)" -UserCredential $local:GraphUserCredential -ClientId $local:GraphClientId -TenantId $local:GraphTenantId -RedirectUri 'http://localhost' -Scopes "$($script:CloudEnvironmentGraphApiEndpoint)/.default"
			$authExo = get-msaltoken -Authority "$($script:CloudEnvironmentAzureADEndpoint)/$(@($local:GraphTenantId, 'organizations') | Where-Object { $_ } | Select-Object -First 1)" -UserCredential $local:GraphUserCredential -ClientId $local:GraphClientId -TenantId $local:GraphTenantId -RedirectUri 'http://localhost' -Scopes "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default"

			# App authentication
			$Appauth = get-msaltoken -Authority "$($script:CloudEnvironmentAzureADEndpoint)/$(@($local:GraphTenantId, 'organizations') | Where-Object { $_ } | Select-Object -First 1)" -ClientId $local:GraphClientId -ClientSecret ($local:GraphClientSecret | ConvertTo-SecureString -AsPlainText -Force) -TenantId $local:GraphTenantId -RedirectUri 'http://localhost' -Scopes "$($script:CloudEnvironmentGraphApiEndpoint)/.default"
			$AppauthExo = get-msaltoken -Authority "$($script:CloudEnvironmentAzureADEndpoint)/$(@($local:GraphTenantId, 'organizations') | Where-Object { $_ } | Select-Object -First 1)" -ClientId $local:GraphClientId -ClientSecret ($local:GraphClientSecret | ConvertTo-SecureString -AsPlainText -Force) -TenantId $local:GraphTenantId -RedirectUri 'http://localhost' -Scopes "$($script:CloudEnvironmentExchangeOnlineEndpoint)/.default"

			$returnValues = @{
				'error'             = $false
				'AccessToken'       = $auth.AccessToken
				'AuthHeader'        = $auth.createauthorizationheader()
				'AccessTokenExo'    = $authExo.AccessToken
				'AuthHeaderExo'     = $authExo.createauthorizationheader()
				'AppAccessToken'    = $Appauth.AccessToken
				'AppAuthHeader'     = $Appauth.createauthorizationheader()
				'AppAccessTokenExo' = $AppauthExo.AccessToken
				'AppAuthHeaderExo'  = $AppauthExo.createauthorizationheader()
			}
		} catch {
			$returnValues = @{
				'error'             = $error[0] | Out-String
				'AccessToken'       = $null
				'AuthHeader'        = $null
				'AccessTokenExo'    = $null
				'AuthHeaderExo'     = $null
				'AppAccessToken'    = $null
				'AppAuthHeader'     = $null
				'AppAccessTokenExo' = $null
				'AppAuthHeaderExo'  = $null
			}
		}

		$local:GraphTokenDictionary[$local:GraphTenantId] = $returnValues

		$returnValuesCollection += , $returnValues
	}

	if ($returnValuesCollection.count -eq 1) {
		$null = $returnValues | Export-Clixml -Path $SimulateAndDeployGraphCredentialFile
	} else {
		$null = $(
			@{
				GraphDomainToTenantIDCache  = $script:GraphDomainToTenantIDCache
				GraphDomainToCloudNameCache = $script:GraphDomainToCloudNameCache
				GraphTokenDictionary        = $local:GraphTokenDictionary
			} | Export-Clixml -Path $SimulateAndDeployGraphCredentialFile
		)
	}

	return $returnValuesCollection
}


function GraphDomainToTenantID {
	param (
		[string]$domain = 'explicitconsulting.at',
		[uri]$SpecificGraphApiEndpointOnly = $null
	)

	if (-not $script:GraphDomainToTenantIDCache) {
		$script:GraphDomainToTenantIDCache = @{}
	}

	if (-not $script:GraphDomainToCloudNameCache) {
		$script:GraphDomainToCloudNameCache = @{}
	}

	$domain = $domain.Trim().ToLower()


	# If $domain is a mail address, extract the domain part
	try {
		$tempDomain = [mailaddress]$domain

		if ($tempDomain.Host) {
			$domain = $tempDomain.Host
		}
	} catch {}

	# If $domain is a URL, extract the DNS safe host
	try {
		$tempDomain = [uri]$domain
		if ($tempDomain.DnsSafeHost) {
			$domain = $tempDomain.DnsSafeHost
		}
	} catch {
		# Not a URI, do nothing
	}

	foreach ($SharePointDomain in $script:CloudEnvironmentSharePointOnlineDomains) {
		if ($domain.EndsWith("-my.$($SharePointDomain)")) {
			$domain = $domain -ireplace "-my.$($SharePointDomain)", '.onmicrosoft.com'

			break
		}
	}

	if ([string]::IsNullOrWhitespace($domain)) {
		return
	}

	if ($script:GraphDomainToTenantIDCache.ContainsKey($domain)) {
		return $script:GraphDomainToTenantIDCache[$domain]
	}

	try {
		try { global:WatchCatchableExitSignal } catch {}

		$local:QueryRetryMaxRetries = 3
		$local:QueryRetryRetryCount = 0
		$local:QueryRetryDone = $false

		do {
			try {
				$local:result = Invoke-RestMethod -UseBasicParsing -Uri "https://odc.officeapps.live.com/odc/v2.1/federationprovider?domain=$($domain)" -ErrorAction Stop
				$local:QueryRetryDone = $true
			} catch {
				$local:QueryRetryWaitTime = 0

				$local:QueryRetryResponse = $_.Exception.Response

				if ($null -ne $local:QueryRetryResponse) {
					$local:QueryRetryRetryAfterHeader = $local:QueryRetryResponse.Headers['Retry-After']

					$local:QueryRetryStatusCode = [int]$local:QueryRetryResponse.StatusCode

					if ($null -ne $local:QueryRetryRetryAfterHeader) {
						if ($local:QueryRetryRetryAfterHeader -match '^\d+$') {
							$local:QueryRetryWaitTime = [Math]::Max([int]$local:QueryRetryRetryAfterHeader + 1, 3 * [Math]::Pow(($local:QueryRetryRetryCount + 1), 2))
						} else {
							try {
								$retryDate = [DateTime]::Parse($local:QueryRetryRetryAfterHeader).ToUniversalTime()
								$local:QueryRetryWaitTime = [Math]::Max([int]($retryDate - (Get-Date).ToUniversalTime()).TotalSeconds + 1, 3 * [Math]::Pow(($local:QueryRetryRetryCount + 1), 2))
							} catch {
								$local:QueryRetryWaitTime = 3 * [Math]::Pow(($local:QueryRetryRetryCount + 1), 2)
							}
						}
					} elseif ($local:QueryRetryStatusCode -in @(408, 429, 500, 502, 503, 504)) {
						$local:QueryRetryWaitTime = 3 * [Math]::Pow(($local:QueryRetryRetryCount + 1), 2)
					}
				} else {
					$local:QueryRetryWaitTime = 3 * [Math]::Pow(($local:QueryRetryRetryCount + 1), 2)
				}

				if ($local:QueryRetryWaitTime -gt 0 -and $local:QueryRetryRetryCount -lt $local:QueryRetryMaxRetries) {
					$local:QueryRetryRetryCount++

					Write-Verbose "Retry attempt $local:QueryRetryRetryCount. Retryable error ($($_.Exception.Response.StatusCode.value__)). Waiting $($local:QueryRetryWaitTime)s."

					Start-Sleep -Seconds $local:QueryRetryWaitTime
				} else {
					throw $_
				}
			}
		} while (-not $local:QueryRetryDone)

		try { global:WatchCatchableExitSignal } catch {}

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
			$script:GraphDomainToCloudNameCache[$domain] = $script:GraphDomainToCloudNameCache[$local:result.tenantId] = (
				@(
					@(@($script:CloudEnvironmentsData | Where-Object { $($local:result.authority_host -ieq ([uri]$_.AzureADEndpoint).DnsSafeHost -and [uri]$local:result.graph -ieq [uri]$_.GraphApiEndpoint) }) | Select-Object -Last 1) +
					@($script:CloudEnvironmentsData | Where-Object { $_.Aliases -icontains 'Public' })
				) | Select-Object -First 1
			).Name

			return $local:result.tenantId
		} else {
			return
		}
	} catch {
		if ($domain -and $script:GraphDomainToTenantIDCache.ContainsKey($domain)) {
			$script:GraphDomainToTenantIDCache.Remove($domain)
		}

		return
	}
}


function GraphSwitchContext {
	param (
		$TenantID = $null
	)

	try {
		if ($null -eq $TenantID -and $script:GraphUser) {
			$TenantID = GraphDomainToTenantID -domain ($script:GraphUser -split '@')[1]
		} else {
			$TenantID = GraphDomainToTenantID -domain $TenantID
		}

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

		if ($TenantID -and $script:GraphDomainToCloudNameCache.ContainsKey($TenantID) -and $script:GraphDomainToCloudNameCache[$TenantID]) {
			$CloudEnvironment = $script:GraphDomainToCloudNameCache[$TenantID]
		}

		if (-not $script:CloudEnvironmentsData) {
			# Endpoints from
			#  https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/blob/main/src/client/Microsoft.Identity.Client/Instance/Discovery/KnownMetadataProvider.cs
			#  https://github.com/microsoft/CSS-Exchange/blob/main/Shared/AzureFunctions/Get-CloudServiceEndpoint.ps1

			$script:CloudEnvironmentsRawData = @(
				@{
					Aliases                 = @('AzurePublic', 'Public', 'Global', 'AzureGlobal', 'AzureCloud', 'AzureUSGovernmentGCC', 'USGovernmentGCC')
					AzureADEndpoint         = 'https://login.microsoftonline.com'
					GraphApiEndpoint        = 'https://graph.microsoft.com'
					ExchangeOnlineEndpoint  = 'https://outlook.office.com'
					AutodiscoverSecureName  = 'https://autodiscover-s.outlook.com'
					SharePointOnlineDomains = @('sharepoint.com')
				}

				@{
					Aliases                 = @('AzureUSGovernment', 'AzureUSGovernmentGCCHigh', 'AzureUSGovernmentL4', 'USGovernmentGCCHigh', 'USGovernmentL4')
					AzureADEndpoint         = 'https://login.microsoftonline.us'
					GraphApiEndpoint        = 'https://graph.microsoft.us'
					ExchangeOnlineEndpoint  = 'https://outlook.office365.us'
					AutodiscoverSecureName  = 'https://autodiscover-s.office365.us'
					SharePointOnlineDomains = @('sharepoint.us')
				}

				@{
					Aliases                 = @('AzureUSGovernmentDOD', 'AzureUSGovernmentL5', 'USGovernmentDOD', 'USGovernmentL5')
					AzureADEndpoint         = 'https://login.microsoftonline.us'
					GraphApiEndpoint        = 'https://dod-graph.microsoft.us'
					ExchangeOnlineEndpoint  = 'https://outlook-dod.office365.us'
					AutodiscoverSecureName  = 'https://autodiscover-s-dod.office365.us'
					SharePointOnlineDomains = @('dps.mil', 'sharepoint-mil.us')
				}

				@{
					Aliases                 = @('AzureChina', 'China', 'ChinaCloud', 'AzureChinaCloud')
					AzureADEndpoint         = 'https://login.partner.microsoftonline.cn'
					GraphApiEndpoint        = 'https://microsoftgraph.chinacloudapi.cn'
					ExchangeOnlineEndpoint  = 'https://partner.outlook.cn'
					AutodiscoverSecureName  = 'https://autodiscover-s.partner.outlook.cn'
					SharePointOnlineDomains = @('sharepoint.cn')
				}

				@{
					Aliases                 = @('AzureBleu', 'Bleu', 'BleuCloud', 'AzureBleuCloud')
					AzureADEndpoint         = 'https://login.sovcloud-identity.fr'
					GraphApiEndpoint        = 'https://graph.svc.sovcloud.fr'
					ExchangeOnlineEndpoint  = 'https://outlook.sovcloud.fr'
					AutodiscoverSecureName  = 'https://autodiscover-s.outlook.sovcloud.fr'
					SharePointOnlineDomains = @('sovcloud-sharepoint.fr')
				}

				@{
					Aliases                 = @('AzureDelos', 'Delos', 'DelosCloud', 'AzureDelosCloud')
					AzureADEndpoint         = 'https://login.sovcloud-identity.de'
					GraphApiEndpoint        = 'https://graph.svc.sovcloud.de'
					ExchangeOnlineEndpoint  = 'https://outlook.sovcloud.de'
					AutodiscoverSecureName  = 'https://autodiscover-s.outlook.sovcloud.de'
					SharePointOnlineDomains = @('sovcloud-sharepoint.de')
				}

				@{
					Aliases                 = @('AzureGovSG', 'GovSG', 'GovSGCloud', 'AzureGovSGCloud')
					AzureADEndpoint         = 'https://login.sovcloud-identity.sg'
					GraphApiEndpoint        = 'https://graph.svc.sovcloud.sg'
					ExchangeOnlineEndpoint  = ''
					AutodiscoverSecureName  = ''
					SharePointOnlineDomains = @()
				}
			)


			if ($CustomCloudEnvironments) {
				$script:CloudEnvironmentsRawData += $CustomCloudEnvironments
			}


			$script:CloudEnvironmentsRawDataDuplicateAliases = $script:CloudEnvironmentsRawData.Aliases | Group-Object | Where-Object { $_.Count -gt 1 }

			if ($script:CloudEnvironmentsRawDataDuplicateAliases) {
				Write-Host "Duplicate cloud environment aliases found: $($script:CloudEnvironmentsRawDataDuplicateAliases.Name -join ', ')" -ForegroundColor Red
				$script:ExitCode = 42
				$script:ExitCodeDescription = 'Cloud environments not configured correctly.'
				exit
			}

			foreach ($script:CloudEnvironmentsRawDataEntry in $script:CloudEnvironmentsRawData) {
				if ($null -eq $script:CloudEnvironmentsRawDataEntry.Aliases -or
					$script:CloudEnvironmentsRawDataEntry.Aliases.Count -eq 0 -or
					$null -in $script:CloudEnvironmentsRawDataEntry.Aliases) {
					Write-Host "Validation Failed: 'Aliases' must not be null or contain null values." -ForegroundColor Red
					$script:ExitCode = 42
					$script:ExitCodeDescription = 'Cloud environments not configured correctly.'
					exit
				}

				@($script:CloudEnvironmentsRawDataEntry.Keys) | Where-Object { $_ -inotin @('Aliases', 'SharePointOnlineDomains') } | ForEach-Object {
					$script:CloudEnvironmentsRawDataEntry.$_ = $script:CloudEnvironmentsRawDataEntry.$_.TrimEnd('/')

					if ([String]::IsNullOrWhiteSpace($script:CloudEnvironmentsRawDataEntry.$_)) {
						continue
					}

					try {
						$script:CloudEnvironmentsRawDataEntryUri = [System.Uri]$script:CloudEnvironmentsRawDataEntry.$_

						if (
							$script:CloudEnvironmentsRawDataEntryUri.Scheme -ine 'https' -or
							$script:CloudEnvironmentsRawDataEntryUri.IsFile
						) {
							throw 'Validation Failed'
						}
					} catch {
						Write-Host "Invalid URL format in '$($script:CloudEnvironmentsRawDataEntry.Aliases[0])','$($script:CloudEnvironmentsRawDataEntry.$_)': Is not https, or is file." -ForegroundColor Red
						$script:ExitCode = 42
						$script:ExitCodeDescription = 'Cloud environments not configured correctly.'
						exit
					}
				}

				@($script:CloudEnvironmentsRawDataEntry.Keys) | Where-Object { $_ -ieq 'AzureAdEndpoint' } | ForEach-Object {
					if (([uri]$script:CloudEnvironmentsRawDataEntry.$_.TrimEnd('/')).Segments.Count -eq 1) {
						$script:CloudEnvironmentsRawDataEntry.$_ = $script:CloudEnvironmentsRawDataEntry.$_.TrimEnd('/')
					}
				}

				foreach ($local:SharePointOnlineDomainEntry in $script:CloudEnvironmentsRawDataEntry.SharePointOnlineDomains) {
					try {
						$local:DnsSafeHost = ([uri]$local:SharePointOnlineDomainEntry).DnsSafeHost

						if ($local:DnsSafeHost) {
							$local:SharePointOnlineDomainEntry = $local:DnsSafeHost
						}
					} catch {}
				}

				$script:CloudEnvironmentsRawDataEntry.SharePointOnlineDomains = @($script:CloudEnvironmentsRawDataEntry.SharePointOnlineDomains | Where-Object { -not [String]::IsNullOrWhiteSpace($_) })
			}

			$script:CloudEnvironmentsData = foreach ($Entry in $script:CloudEnvironmentsRawData) {
				$Obj = [PSCustomObject]$Entry
				$Obj | Add-Member -MemberType ScriptProperty -Name 'Name' -Value { $this.Aliases[0] } -Force
				$Obj
			}
		}

		# Get SharePoint Online domains
		if (-not $script:CloudEnvironmentSharePointOnlineDomains) {
			$OldProgressPreference = $ProgressPreference
			$ProgressPreference = 'SilentlyContinue'

			$script:CloudEnvironmentSharePointOnlineDomains = @($script:CloudEnvironmentsData.SharePointOnlineDomains)

			try {
				@((Invoke-RestMethod -Uri "https://endpoints.office.com/version?ClientRequestId=$((New-Guid).Guid)" -Method Get -ErrorAction Stop).instance) | ForEach-Object {
					# https://learn.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide says:
					#  If the query does not contain the TenantName, tenant name specific domain and host names are returned with a '*' wildcard.
					#  As SharePoint Online does not allow custom domains, we only need the '*' entries.
					#  We can then match against them using ilike, for example.
					#  Hint: We only filter for '*' and not also for 'sharepoint' because of '*.dps.mil', '*.cloud.microsoft', '*.sovcloud.fr', and others.

					$script:CloudEnvironmentSharePointOnlineDomains += @(@(@((Invoke-RestMethod -Method Get -Uri "https://endpoints.office.com/endpoints/$($_)?ServiceAreas=SharePoint&ClientRequestId=$((New-Guid).Guid)" -ErrorAction Stop) | Where-Object { ($_.serviceArea -ieq 'SharePoint') -and ($_.ips) }).urls) | Where-Object { $_ -and $_.Contains('*.') })
				}

				$script:CloudEnvironmentSharePointOnlineDomains = @($script:CloudEnvironmentSharePointOnlineDomains | ForEach-Object { ($_ -split '\*\.')[-1] } | Where-Object { $_ } | Select-Object -Unique)
			} catch {
			}

			$ProgressPreference = $OldProgressPreference
		}

		if ($CloudEnvironment -inotin $script:CloudEnvironmentsData.Aliases) {
			Write-Host "Cloud environment '$($CloudEnvironment)' is not defined." -ForegroundColor Red
			$script:ExitCode = 42
			$script:ExitCodeDescription = 'Cloud environments not configured correctly.'
			exit
		}

		$script:CloudEnvironmentsData | Where-Object { $_.Aliases -icontains $CloudEnvironment } | ForEach-Object {
			$CloudEnvironment = $_.Aliases[0]
			$script:CloudEnvironmentAzureADEndpoint = $_.AzureADEndpoint
			$script:CloudEnvironmentGraphApiEndpoint = $_.GraphApiEndpoint
			$script:CloudEnvironmentExchangeOnlineEndpoint = $_.ExchangeOnlineEndpoint
			$script:CloudEnvironmentAutodiscoverSecureName = $_.AutodiscoverSecureName
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
}


function RemoveItemAlternativeRecurse {
	# Function to avoid problems with OneDrive throwing "Access to the cloud file is denied"

	param(
		[alias('LiteralPath')][string] $Path,
		[switch] $SkipFolder # when $Path is a folder, do not delete $path, only it's content
	)

	try { global:WatchCatchableExitSignal } catch {}

	$local:ToDelete = @()

	if (Test-Path -LiteralPath $path) {
		foreach ($SinglePath in @(Get-Item -LiteralPath $Path)) {
			try { global:WatchCatchableExitSignal } catch {}

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
		try { global:WatchCatchableExitSignal } catch {}

		try {
			if ((Test-Path -LiteralPath $SingleItemToDelete.FullName) -eq $true) {
				Remove-Item -LiteralPath $SingleItemToDelete.FullName -Force -Recurse
			}
		} catch {
			Write-Verbose "Could not delete $($SingleItemToDelete.FullName), error: $($_.Exception.Message)"
			Write-Verbose $_
		}
	}

	try { global:WatchCatchableExitSignal } catch {}
}

try {
	# Start script
	Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	# Remove unnecessary ETS type data associated with arrays in Windows PowerShell
	Remove-TypeData System.Array -ErrorAction SilentlyContinue

	if ($psISE) {
		Write-Host '  PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
		Write-Host '  Required features are not available in ISE. Exit.' -ForegroundColor Red
		exit 1
	}

	if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
		Write-Host "This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode." -ForegroundColor Red
		Write-Host 'Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
		exit 1
	}

	Write-Host '  Prepare objects'

	$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

	if ($PSScriptRoot) {
		Set-Location -LiteralPath $PSScriptRoot
	} else {
		Write-Host '  Could not determine the script path, which is essential for this script to work.' -ForegroundColor Red
		Write-Host '  Make sure to run this script as a file from a PowerShell console, and not just as a text selection in a code editor.' -ForegroundColor Red
		Write-Host '  Exit.' -ForegroundColor Red
		exit 1
	}

	if (-not $GraphData) {
		$GraphData = @(
			, @($(@($GraphUserCredential.username -split '@')[1]), $GraphUserCredential.username, (New-Object PSCredential 0, $GraphUserCredential.Password).GetNetworkCredential().Password; $GraphClientID, $GraphClientSecret)
		)
	}

	$SetOutlookSignaturesScriptParameters.GraphClientID = $GraphData

	GraphSwitchContext -TenantID (GraphDomainToTenantID -domain $GraphData[0])

	$SetOutlookSignaturesScriptParameters.CloudEnvironment = $CloudEnvironment


	Write-Host '  Prepare folders'

	$script:tempDir = (New-Item -Path ([System.IO.Path]::GetTempPath()) -Name (New-Guid).Guid -ItemType Directory).FullName

	foreach ($VariableName in ('SimulateResultPath', 'SetOutlookSignaturesScriptPath')) {
		Set-Variable -Name $VariableName -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Get-Variable -Name $VariableName).Value).trimend('\')
	}

	if (-not (Test-Path -LiteralPath $SimulateResultPath)) {
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

	try {
		$TranscriptFullName = (Join-Path -Path $SimulateResultPath -ChildPath '_log.txt')
		$TranscriptFullName = (Start-Transcript -LiteralPath $TranscriptFullName -Force).Path

		Write-Host "  Log file: '$TranscriptFullName'"
		Write-Host "    Ignore log lines starting with 'PS>TerminatingError' or '>> TerminatingError' unless instructed otherwise."
	} catch {
		$TranscriptFullName = $null
	}


	# Connect to Graph
	if (-not $ConnectOnpremInsteadOfCloud) {
		Write-Host
		Write-Host "Connect to Graph @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

		Write-Host '  Microsoft Graph'
		$SimulateAndDeployGraphCredentialFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "$((New-Guid).Guid).xml"

		$SetOutlookSignaturesScriptParameters['SimulateAndDeployGraphCredentialFile'] = $SimulateAndDeployGraphCredentialFile

		$script:MsalModulePath = (Join-Path -Path $script:tempDir -ChildPath 'MSAL.PS')

		Copy-Item -LiteralPath $([System.Io.Path]::GetFullPath($((Join-Path -Path (Split-Path $SetOutlookSignaturesScriptPath) -ChildPath 'deps\MSAL.PS')))) -Destination $script:MsalModulePath -Recurse
		if (-not ((Test-Path -LiteralPath 'variable:IsLinux') -and $IsLinux)) { Get-ChildItem -LiteralPath $script:MsalModulePath -Recurse | Unblock-File }
		Import-Module $script:MsalModulePath -Force

		$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

		if (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0) {
			Start-Sleep -Seconds 10

			$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

			if (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0) {
				Write-Host '    Exiting because of repeated Graph connection error' -ForegroundColor Red
				Write-Host "    $($GraphConnectResult.error)" -ForegroundColor Red
				exit 1
			}
		}
	}


	# Load and check SimulateList
	Write-Host
	Write-Host "Load and check list of mailboxes to simulate @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
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
			$SimulateEntry.SimulateMailboxes = $null
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
	if (((-not (Test-Path -LiteralPath 'variable:IsWindows')) -or $IsWindows) -and ($SetOutlookSignaturesScriptParameters.UseHtmTemplates -inotin (1, '1', 'true', '$true', 'yes'))) {
		Write-Host
		Write-Host "Memorize Word security setting and disable it @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
		$script:WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace [Regex]::Escape('Word.Application.'), '') + '.0.0.0.0')) -ireplace '^\.', '' -split '\.')[0..3] -join '.'))
		if ($script:WordRegistryVersion.major -gt 16) {
			Write-Host "  Word version $($script:WordRegistryVersion) is newer than 16 and not yet known. Please inform your administrator. Exit." -ForegroundColor Red
			exit 1
		} elseif ($script:WordRegistryVersion.major -eq 16) {
			$script:WordRegistryVersion = '16.0'
		} elseif ($script:WordRegistryVersion.major -eq 15) {
			$script:WordRegistryVersion = '15.0'
		} elseif ($script:WordRegistryVersion.major -eq 14) {
			$script:WordRegistryVersion = '14.0'
		} elseif ($script:WordRegistryVersion.major -lt 14) {
			Write-Host "    Word version $($script:WordRegistryVersion) is older than Word 2010 and not supported. Please inform your administrator. Exit." -ForegroundColor Red
			exit 1
		}

		if ($null -eq (Get-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name 'DisableWarningOnIncludeFieldsUpdate' -ErrorAction SilentlyContinue).DisableWarningOnIncludeFieldsUpdate) {
			$null = "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 0 -Force
		}

		if ($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) {
			if (Test-Path -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security") {
				$script:WordDisableWarningOnIncludeFieldsUpdate = (Get-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security").'DisableWarningOnIncludeFieldsUpdate'
			}
		}

		if (($null -eq $script:WordDisableWarningOnIncludeFieldsUpdate) -or ($script:WordDisableWarningOnIncludeFieldsUpdate -ne 1)) {
			$null = "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" | ForEach-Object { if (Test-Path -LiteralPath $_) { Get-Item -LiteralPath $_ } else { New-Item $_ -Force } } | New-ItemProperty -Name 'DisableWarningOnIncludeFieldsUpdate' -Type DWORD -Value 1 -Force
		}
	}

	# Run simulation mode for each user
	Write-Host
	Write-Host "Run simulation mode for each user and its mailbox(es) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	Write-Host '  Remove old jobs'

	Get-Job | Remove-Job -Force

	$JobsToStartTotal = ($SimulateList | Measure-Object).count
	$JobsToStartOpen = ($SimulateList | Measure-Object).count
	$JobsStarted = 0
	$JobsCompleted = 0

	Write-Host "  $JobstoStartTotal jobs total: $JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress, $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"


	$UpdateTime = (Get-Date).Add($UpdateInterval)

	do {
		while ((($JobsToStartOpen -gt 0) -and ((Get-Job).count -lt $JobsConcurrent))) {
			$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$Jobsstarted].SimulateUser)) -ChildPath '_log.txt'

			if ((Test-Path -LiteralPath (Split-Path -LiteralPath $LogFilePath)) -eq $false) {
				New-Item -ItemType Directory -Path (Split-Path -LiteralPath $LogFilePath) | Out-Null
			}

			# Update Graph credential file before starting a job
			#   this makes sure that the token is still valid when the software runs longer than token lifetime
			if (
				$($ConnectOnpremInsteadOfCloud -ne $true) -and
				$(
					$(-not $GraphConnectResult) -or
					$($GraphConnectResult -and (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0)) -or
					$($GraphConnectResult -and
						$(
							@(
								$(
									foreach ($SingleGraphConnectResult in $GraphConnectResult) {
										foreach ($tempTokenType in @('AccessToken', 'AccessTokenExo', 'AppAccessToken', 'AppAccessTokenExo')) {
											$tempParsedToken = ParseJwtToken -token $($SingleGraphConnectResult.$tempTokenType)
											$tempCurrentTimeUtcUnixTimestamp = Get-Date -UFormat %s -Millisecond 0 -Date (Get-Date).ToUniversalTime()

											# True if
											#   Token is expired
											#   The remaining token lifetime is less than or equals the job timeout
											#   At least half of the token lifetime has already passed
											$(
												$($tempCurrentTimeUtcUnixTimestamp -ge $tempParsedToken.payload.exp) -or
												$(($tempParsedToken.payload.exp - $tempCurrentTimeUtcUnixTimestamp) -le $JobTimeout.TotalSeconds) -or
												$($tempCurrentTimeUtcUnixTimestamp -ge ($tempParsedToken.payload.nbf + (($tempParsedToken.payload.exp - $tempParsedToken.payload.nbf) / 2)))
											)
										}
									}
								)
							) -icontains $true
						)
					)
				)
			) {
				Write-Host '    Renewing Graph token'

				$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

				if (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0) {
					Start-Sleep -Seconds 70

					$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

					if (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0) {
						Start-Sleep -Seconds 70

						$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

						if (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0) {
							Start-Sleep -Seconds 70

							$GraphConnectResult = CreateUpdateSimulateAndDeployGraphCredentialFile

							if (($GraphConnectResult | Where-Object { $_.error -ne $false }).Count -gt 0) {
								Write-Host '    Exiting because of repeated Graph connection error' -ForegroundColor Red
								Write-Host "    $($GraphConnectResult.error)" -ForegroundColor Red
								exit 1
							}
						}
					}
				}
			}

			Start-Job {
				param (
					$SetOutlookSignaturesScriptPath,
					$SimulateUser,
					$SimulateMailboxes,
					$SimulateResultPath,
					$LogFilePath,
					$SetOutlookSignaturesScriptParameters
				)

				Start-Transcript -LiteralPath $LogFilePath -Force

				try {
					Write-Host 'CREATE SIGNATURE FILES BY USING SIMULATON MODE OF SET-OUTLOOKSIGNATURES'

					$SetOutlookSignaturesScriptParameters['SimulateUser'] = $SimulateUser
					$SetOutlookSignaturesScriptParameters['SimulateMailboxes'] = $SimulateMailboxes
					$SetOutlookSignaturesScriptParameters['AdditionalSignaturePath'] = $(Join-Path -Path $SimulateResultPath -ChildPath $SimulateUser)

					& $SetOutlookSignaturesScriptPath @SetOutlookSignaturesScriptParameters

					if ($?) {
						Write-Host 'xxxSimulateAndDeployExitCode0xxx'
					} else {
						Write-Host 'xxxSimulateAndDeployExitCode999xxx'
					}
				} catch {
					Write-Host ($error[0] | Format-List * | Out-String)
					Write-Host 'xxxSimulateAndDeployExitCode999xxx'
				}

				Stop-Transcript
			} -Name ("$($Jobsstarted)_Job") -ArgumentList $SetOutlookSignaturesScriptPath,
			$($SimulateList[$Jobsstarted].SimulateUser),
			$($SimulateList[$Jobsstarted].SimulateMailboxes),
			$SimulateResultPath,
			$LogFilePath,
			$SetOutlookSignaturesScriptParameters | Out-Null

			"    User $($SimulateList[$Jobsstarted].SimulateUser) started @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
				Write-Host $($_)
				Add-Content -Value $($_.TrimStart()) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_started.txt') -Force -Encoding UTF8
			}

			$JobsToStartOpen--
			$JobsStarted++

			Write-Host "  $JobstoStartTotal jobs total: $JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress, $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
		}

		foreach ($x in (Get-Job | Where-Object { (-not $_.PSEndTime) -and (((Get-Date) - $_.PSBeginTime) -gt $JobTimeout) })) {
			"    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) canceled due to timeout @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
				Write-Host $($_) -ForegroundColor Red
				Add-Content -Value $($_.TrimStart()) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_error.txt') -Force -Encoding UTF8
			}

			$x | Remove-Job -Force

			$JobsCompleted++

			Write-Host "  $JobstoStartTotal jobs total: $JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress, $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
		}

		foreach ($x in (Get-Job | Where-Object { $_.PSEndTime })) {
			$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser)) -ChildPath '_log.txt'

			if ((Get-Content -LiteralPath $LogFilePath -Encoding UTF8 -Raw).trim().Contains('xxxSimulateAndDeployExitCode0xxx')) {
				"    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) ended with no errors @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
					Write-Host $($_) -ForegroundColor Green
					Add-Content -Value $($_.TrimStart()) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_success.txt') -Force -Encoding UTF8
				}
			} else {
				"    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) ended with errors @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@" | ForEach-Object {
					Write-Host $($_) -ForegroundColor Red
					Add-Content -Value $($_.TrimStart()) -LiteralPath (Join-Path -Path $SimulateResultPath -ChildPath '_log_error.txt') -Force -Encoding UTF8
				}
			}

			$x | Remove-Job -Force

			$JobsCompleted++

			Write-Host "  $JobstoStartTotal jobs total: $JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress, $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
		}

		if ((Get-Date) -ge $UpdateTime) {
			Write-Host "  $JobstoStartTotal jobs total: $JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress, $JobsToStartOpen in queue @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
			$UpdateTime = (Get-Date).Add($UpdateInterval)
		}

		Start-Sleep -Seconds 1
	} until (($JobsToStartTotal -eq $JobsStarted) -and ($JobsCompleted -eq $JobsToStartTotal))
} catch {
	Write-Host
	Write-Host ($error[0] | Format-List * | Out-String) -ForegroundColor Red
	Write-Host
	Write-Host 'Unexpected error. Exit.' -ForegroundColor red
} finally {
	Write-Host
	Write-Host "Clean-up @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	Get-Job | Remove-Job -Force

	# Restore Word security setting for embedded images
	if (((-not (Test-Path -LiteralPath 'variable:IsWindows')) -or $IsWindows) -and ($SetOutlookSignaturesScriptParameters.UseHtmTemplates -inotin (1, '1', 'true', '$true', 'yes'))) {
		Set-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\$($script:WordRegistryVersion)\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $script:WordDisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
	}

	if ($script:MsalModulePath) {
		Remove-Module -Name MSAL.PS -Force -ErrorAction SilentlyContinue
		Remove-Item -LiteralPath $script:MsalModulePath -Recurse -Force -ErrorAction SilentlyContinue
	}

	if ($SimulateAndDeployGraphCredentialFile) {
		Remove-Item -LiteralPath $SimulateAndDeployGraphCredentialFile -Force -ErrorAction SilentlyContinue
	}

	if ($script:tempDir) {
		Remove-Item -LiteralPath $script:tempDir -Recurse -Force -ErrorAction SilentlyContinue
	}

	if ($TranscriptFullName) {
		Write-Host
		Write-Host 'Log file'
		Write-Host "  '$TranscriptFullName'"
	}

	Write-Host
	Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

	if ($TranscriptFullName) {
		try { Stop-Transcript | Out-Null } catch {}
	}
}