<#
This script shows how the simulation mode of Set-OutlookSignatures can be used to deploy Outlook text signatures without client involvement.

You have to adopt it to fit your environment.
The sample code is written in a generic way, which allows for easy adoption.

Features
- Automate simulation mode for all given mailboxes
- A configurable number of Set-OutlookSignatures instances run in parallel for better performance
- Copy resulting signatures to a file path (can be the redirected signatures folder of the corresponding user)
- Set default signature in Outlook Web
- Set internal and external Out of Office (OOF) message
- Supports on-prem, hybrid and cloud-only environments

Requirements
- Script needs to be run with an account with approriate read and write permissions to all mailboxes
- Microsoft Word
- Exchange Online PowerShell V2 module (when connecting to the cloud)

Limitations
- Outlook signatures can be deployed to a network share, but the default signatures for new e-mails and replies/forwards can't be configured, as there is no access to the users registry settings. This will be addressed for cloud mailboxes when Microsoft makes their signature roaming API available.
- Requires desktop interaction if Windows Integrated Authentication can not be used to connect to the cloud

Future enhancements
- Enhanced authentication for full non-interactive mode
- Support for upcoming Microsoft signature roaming API (very likely cloud-only)

This scripts performs the following steps
1. Connect to on-prem oder EXO, get list of all mailboxes and export them to a csv file. Format: <UPN>;<PrimarySmtpAddress>
2. Disconnect from on-prem or EXO
3. Import csv
4. (Future feature) If Microsoft roaming signature API is available: Connect to mailbox and download existing signatures
5. Run simulation mode for each CSV line, each using a separate AdditionalSignaturePath (use UPN as folder name). This path could point to a redirected folder network path of the signatures folder.
6. Take the results of simulation mode and apply them via EWS/Graph
a. OOF internal and OOF external (folder "UPN\PrimarySmtpAddress")
b. DefaultNew and DefaultReplyForward (folder "UPN\PrimarySmtpAddress")
c. (Future feature) If Microsoft roaming signature API is available: Delete existing signatures and upload updated ones.
#>


# Variables
$ConnectOnpremInsteadOfCloud = $false
$OnPremServerFqdn = 'server.exchange.example.com'
$SimulateResultPath = $([IO.Path]::Combine([environment]::GetFolderPath('MyDocuments'), 'Set-OutlookSignatures simulation'))
$SimulateListFile = $([IO.Path]::Combine($SimulateResultPath, 'SimulateList.csv'))
$JobsConcurrent = 2
$SetOutlookSignaturesScriptPath = '..\Set-OutlookSignatures.ps1'
$SetOutlookSignaturesScriptParameters = "-SignatureTemplatePath `"C:\temp\Signatures DOCX`" -SignatureIniPath `"C:\temp\Signatures DOCX\_.ini`" -SetCurrentUserOOFMessage `$false" # Do not use: SimulateUser, SimulateMailbox, AdditionalSignaturePath


Set-Location $PSScriptRoot | Out-Null

$PSDefaultParameterValues['out-file:width'] = 2000


('SimulateResultPath', 'SimulateListFile', 'SetOutlookSignaturesScriptPath') | ForEach-Object {
	Set-Variable -Name $_ -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Get-Variable -Name $_).Value).trimend('\')
}

if (-not (Test-Path $SimulateResultPath)) {
	New-Item -ItemType Directory $SimulateResultPath | Out-Null
}


Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"


Write-Host "Get Exchange credentials @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$CredentialPath = Join-Path -Path $env:temp -ChildPath "$((New-Guid).guid).xml"
$Credential = Get-Credential
$Credential | Export-Clixml -Path $CredentialPath
$Credential = Import-Clixml -Path $CredentialPath


Write-Host "Conncect to Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
if ($ConnectOnpremInsteadOfCloud) {
	Write-Host '  On premises'
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$($OnPremServerFqdn)/PowerShell/" -Authentication Kerberos -Credential $Credential
	Import-PSSession $Session -DisableNameChecking
	Set-AdServerSettings -ViewEntireForest $True
} else {
	Write-Host '  Microsoft Graph'
	$GraphCredentialFile = Join-Path -Path $env:temp -ChildPath "$((New-Guid).guid).xml"
	if ($ConnectOnpremInsteadOfCloud -eq $false) {
		$SetOutlookSignaturesScriptParameters = $SetOutlookSignaturesScriptParameters + " -GraphCredentialFile `"$GraphCredentialFile`""
	}
	Import-Module $(Join-Path -Path (Split-Path $SetOutlookSignaturesScriptPath -Parent) -ChildPath '\bin\msal.ps')
	
	# ClientId comes from Set-OutlookSignatures 'default graph config.ps1'
	$auth = get-msaltoken -ClientId 'beea8249-8c98-4c76-92f6-ce3c468a61e6' -tenantid ($credential.username -split '@')[1] -RedirectUri 'http://localhost' -UserCredential $credential
	@{'accessToken' = $auth.accessToken; 'authHeader' = $($auth.createauthorizationheader()) } | Export-Clixml -Path $GraphCredentialFile
	Remove-Module msal.ps

	Write-Host '  Exchange Online'
	Import-Module ExchangeOnlineManagement
	Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
}


Write-Host "Get list of maiboxes and create CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
if ($ConnectOnpremInsteadOfCloud) {
	Write-Host '  On premises'
	$SimulateList = Get-Mailbox | Select-Object -Property @{name = 'SimulateUser'; expression = { $_.userprincipalname } }, @{name = 'SimulateMailbox'; expression = { $_.primarysmtpaddress } }, @{name = 'Environment'; expression = { if ($_.RecipientTypeDetails -like 'Remote*') { 'Cloud' } else { 'On-Prem' } } }
} else {
	Write-Host '  Exchange Online'
	$SimulateList = Get-EXOMailbox | Select-Object -Property @{name = 'SimulateUser'; expression = { $_.userprincipalname } }, @{name = 'SimulateMailbox'; expression = { $_.primarysmtpaddress } }, @{name = 'Environment'; expression = { if ($_.RecipientTypeDetails -like 'Remote*') { 'On-Prem' } else { 'Cloud' } } }
}

$SimulateList | Export-Csv -Path $SimulateListFile -NoTypeInformation -Delimiter ';' -Force


Write-Host "Load CSV file @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$SimulateList = Import-Csv -Path $SimulateListFile -Delimiter ';'
Write-Host "  $(($SimulateList | Measure-Object).count) entries found"


Write-Host "Export Word security setting and disable it @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
$WordRegistryVersion = [System.Version]::Parse(((((((Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Word.Application\CurVer' -ErrorAction SilentlyContinue).'(default)' -ireplace 'Word.Application.', '') + '.0.0.0.0')) -replace '^\.', '' -split '\.')[0..3] -join '.'))
if ($WordRegistryVersion.major -eq 0) {
	$WordRegistryVersion = $null
} elseif ($WordRegistryVersion.major -gt 16) {
	Write-Host "Word version $WordRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exiting." -ForegroundColor Red
	exit 1
} elseif ($WordRegistryVersion.major -eq 16) {
	$WordRegistryVersion = '16.0'
} elseif ($WordRegistryVersion.major -eq 15) {
	$WordRegistryVersion = '15.0'
} elseif ($WordRegistryVersion.major -eq 14) {
	$WordRegistryVersion = '14.0'
} elseif ($WordRegistryVersion.major -lt 14) {
	Write-Host "Word version $WordRegistryVersion is older than Word 2010 and not supported. Please inform your administrator. Exiting." -ForegroundColor Red
	exit 1
}

$WordDisableWarningOnIncludeFieldsUpdate = Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
if (($null -eq $WordDisableWarningOnIncludeFieldsUpdate) -or ($WordDisableWarningOnIncludeFieldsUpdate.DisableWarningOnIncludeFieldsUpdate -ne 1)) {
	New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -PropertyType DWord -Value 1 -ErrorAction Ignore | Out-Null
	Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value 1 -ErrorAction Ignore | Out-Null
}


Write-Host "Run simulation mode for each user and his personal mailbox @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
Get-Job | Remove-Job -Force

$JobsToStartTotal = ($SimulateList | Measure-Object).count
$JobsToStartOpen = ($SimulateList | Measure-Object).count
$JobsStarted = 0
$JobsCompleted = 0


Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"
do {
	while ((($JobsToStartOpen -gt 0) -and ((Get-Job -State running).count -lt $JobsConcurrent))) {
		$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$Jobsstarted].SimulateUser)) -ChildPath '_log.txt'
		if ((Test-Path (Split-Path $LogFilePath -Parent)) -eq $false) {
			New-Item -ItemType Directory -Path (Split-Path $LogFilePath -Parent) | Out-Null
		}

		Start-Job {
			Param (
				$PowershellPath,
				$SetOutlookSignaturesScriptPath,
				[string]$SimulateUser,
				[string]$SimulateMailbox,
				$SimulateResultPath,
				$CredentialPath,
				$LogFilePath,
				$SetOutlookSignaturesScriptParameters,
				$GraphCredentialFile
			)
	
			$PSDefaultParameterValues['out-file:width'] = 2000
	
			Remove-Item -Path $LogFilePath -Force
	
			. {
				try {
					Write-Host 'CREATE SIGNATURE FILES BY USING SIMULATON MODE OF SET-OUTLOOKSIGNATURES'
					#Invoke-Expression $("& `"$PowershellPath`" -executionpolicy bypass -file `"$SetOutlookSignaturesScriptPath`" -SimulateUser `"$SimulateUser`" -SimulateMailbox `"$SimulateMailbox`" -AdditionalSignaturePath `"$(Join-Path -Path $SimulateResultPath -ChildPath $SimulateUser)`" $SetOutlookSignaturesScriptParameters")
					Invoke-Expression $(". `"$SetOutlookSignaturesScriptPath`" -SimulateUser `"$SimulateUser`" -SimulateMailbox `"$SimulateMailbox`" -AdditionalSignaturePath `"$(Join-Path -Path $SimulateResultPath -ChildPath $SimulateUser)`" $SetOutlookSignaturesScriptParameters")
					if ($LASTEXITCODE -eq 0) {
						Write-Host 'xxxExitCode0xxx'
					} else {
						Write-Host "xxxExitCode$($LASTEXITCODE)xxx"
					}
				} catch {
					$error[0]
					Write-Host 'xxxExitCode999xxx'
				}
			} 2>&1 3>&1 4>&1 5>&1 6>&1 | Out-File -FilePath $LogFilePath -Append -Force -Encoding utf8
		} -Name ("$($Jobsstarted)_Job") -ArgumentList (Get-Process -Id $pid).Path,
		$SetOutlookSignaturesScriptPath,
		$($SimulateList[$Jobsstarted].SimulateUser),
		$($SimulateList[$Jobsstarted].SimulateMailbox),
		$SimulateResultPath,
		$CredentialPath,
		$LogFilePath,
		$SetOutlookSignaturesScriptParameters,
		$GraphCredentialFile | Out-Null

		Write-Host "    User $($SimulateList[$Jobsstarted].SimulateUser) (mailbox $($SimulateList[$Jobsstarted].SimulateMailbox)) started @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"

		$JobsToStartOpen--
		$JobsStarted++
	}

	foreach ($x in (Get-Job -State Completed)) {
		$LogFilePath = Join-Path -Path (Join-Path -Path $SimulateResultPath -ChildPath $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser)) -ChildPath '_log.txt'
		if (-not (Get-Content -Path $LogFilePath -Encoding UTF8 -Raw).trim().EndsWith('xxxExitCode0xxx')) {
			Write-Host "    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) (mailbox $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)) ended @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
			Write-Host "      User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) (mailbox $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)): Error creating signatures, please check log." -ForegroundColor Red
		} else {
			. {
				try {
					Write-Host
					Write-Host
					Write-Host 'SET SIGNATURES AND OOF MESSAGES'

					Set-Location (Split-Path $LogFilePath -Parent)

					Write-Host 'All signature names'
					Get-ChildItem '*.htm' | ForEach-Object {
						Write-Host "  $($_.basename)"
					}

					Write-Host 'Default signature name for new e-mails'
					if (Test-Path ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox))\default new.htm") {
						$hash = (Get-FileHash ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\default new.htm").hash
						$SignatureFilePathDefaultNew = (Get-FileHash '*.htm' | Where-Object { $_.hash -eq $hash })[0].path
						Write-Host "  $((Get-ChildItem $SignatureFilePathDefaultNew).basename)"
					} else {
						$SignatureFilePathDefaultNew = $null
					}

					Write-Host 'Default signature name for replies and forwards'
					if (Test-Path ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\default reply-forward.htm") {
						$hash = (Get-FileHash ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\default reply-forward.htm").hash
						$SignatureFilePathDefaultReplyforward = (Get-FileHash '*.htm' | Where-Object { $_.hash -eq $hash })[0].path
						Write-Host "  $((Get-ChildItem $SignatureFilePathDefaultReplyforward).basename)"
					} else {
						$SignatureFilePathDefaultReplyforward = $null
					}

					Write-Host 'Determine signature to use in Outlook Web'
					if ($null -ne $SignatureFilePathDefaultNew) {
						$WebSigHtml = Get-Content -Path $SignatureFilePathDefaultNew -Encoding UTF8 -Raw
						$WebSigTxt = Get-Content -Path	$([System.IO.Path]::ChangeExtension($SignatureFilePathDefaultNew, '.txt')) -Encoding UTF8 -Raw
					} else {
						if ($null -ne $SignatureFilePathDefaultReplyforward) {
							$WebSigHtml = Get-Content -Path $SignatureFilePathDefaultReplyforward -Encoding UTF8 -Raw
							$WebSigTxt = Get-Content -Path	$([System.IO.Path]::ChangeExtension($SignatureFilePathDefaultReplyforward, '.txt')) -Encoding UTF8 -Raw
						} else {
							$WebSigHtml = $null
							$WebSigTxt = $null
						}
					}

					if ($WebSigHtml) {
						Write-Host 'Set signature in Outlook Web'

						# Set-MailboxMessageConfiguration requires a specific, non-standard definition for inline images
						# With Exchange Web Services this is not necessary
						($WebSigHtml | Select-String -Pattern '\s*src\="(data:image\/.*?"\s*>)' -AllMatches).Matches | ForEach-Object {
							if ($_.groups.count -ge 2) {
								if ($null -ne $_.groups[1]) {
									$WebSigHtml = $WebSigHtml.Replace($_.groups[1].value, ('"><span id="dataURI" style="display:none">' + (($_.groups[1].value) -Replace '"\s*>', '') + '</span>'))
								}
							}
						}

						Set-MailboxMessageConfiguration `
							-Identity $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox) `
							-SignatureHTML $WebSigHtml `
							-SignatureText $WebSigTxt `
							-SignatureTextOnMobile $WebSigTxt `
							-AutoAddSignature $( if ($SignatureFilePathDefaultNew) { $true } else { $false } ) `
							-AutoAddSignatureOnMobile $( if ($SignatureFilePathDefaultNew) { $true } else { $false } ) `
							-AutoAddSignatureOnReply $( if ($SignatureFilePathDefaultReplyforward -eq $SignatureFilePathDefaultNew) { $true } else { $false } ) `
							-UseDefaultSignatureOnMobile $true
					}

					Write-Host 'Determine internal Out of Office (OOF) auto reply message'
					if (Test-Path ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\oof internal.htm") {
						$OOFInternalHtml = Get-Content ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\oof internal.htm" -Encoding UTF8 -Raw
					} else {
						$OOFInternalHtml = $null
					}

					Write-Host 'Determine external Out of Office (OOF) auto reply message'
					if (Test-Path ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\oof external.htm") {
						$OOFExternalHtml = Get-Content ".\$($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)\oof external.htm" -Encoding UTF8 -Raw
					} else {
						$OOFExternalHtml = $null
					}


					if ($OOFInternalHtml -or $OOFExternalHtml) {
						Write-Host 'Set OOF messages'

						# Set-MailboxAutoReplyConfiguration can't handle inline images, they need to be removed to avoid display errors
						# With Exchange Web Services this is not necessary
						$OOFInternalHtml = $OOFInternalHtml -replace '(?ms)<\s*?img.*?src\="data:image\/.*?".*?>', ''
						$OOFExternalHtml = $OOFExternalHtml -replace '(?ms)<\s*?img.*?src\="data:image\/.*?".*?>', ''

						if ((Get-MailboxAutoReplyConfiguration -Identity $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)).AutoReplyState -ieq 'disabled') {
							Set-MailboxAutoReplyConfiguration -Identity $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox) -InternalMessage $OOFInternalHtml -ExternalMessage $OOFExternalHtml
						}
					}

					Write-Host 'xxxExitCode0xxx'
				} catch {
					$error[0]
					Write-Host 'xxxExitCode999xxx'
				}
			} 2>&1 3>&1 4>&1 5>&1 6>&1 | Out-File -FilePath $LogFilePath -Append -Force -Encoding utf8

			Write-Host "    User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) (mailbox $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)) ended @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
			if (-not (Get-Content -Path $LogFilePath -Encoding UTF8 -Raw).trim().EndsWith('xxxExitCode0xxx')) {
				Write-Host "      User $($SimulateList[$($x.name.trimend('_Job'))].SimulateUser) (mailbox $($SimulateList[$($x.name.trimend('_Job'))].SimulateMailbox)): Error setting signatures, please check log." -ForegroundColor Red
			}
			Set-Location $PSScriptRoot
		}
		$x | Remove-Job -Force
		
		$JobsCompleted++
		
		Write-Host "  $JobstoStartTotal jobs total: $JobsStarted started ($JobsCompleted completed, $($JobsStarted - $JobsCompleted) in progress), $JobsToStartOpen in queue"
	}

} until (($JobsToStartTotal -eq $JobsStarted) -and ($JobsCompleted -eq $JobsToStartTotal))



Write-Host "Restore original Word security setting @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
if ($null -eq $WordDisableWarningOnIncludeFieldsUpdate) {
	Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore
} else {
	Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\$WordRegistryVersion\Word\Security" -Name DisableWarningOnIncludeFieldsUpdate -Value $WordDisableWarningOnIncludeFieldsUpdate.DisableWarningOnIncludeFieldsUpdate -ErrorAction Ignore | Out-Null
}


Write-Host "Disconncect from Exchange @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
if ($ConnectOnpremInsteadOfCloud) {
	Write-Host '  On premises'
	Remove-PSSession $Session -Confirm:$false
} else {
	Write-Host '  Exchange Online'
	Disconnect-ExchangeOnline -Confirm:$false
}


Write-Host "Cleanup @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
Remove-Item -Force $CredentialPath
Remove-Item -Force $GraphCredentialFile


Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"