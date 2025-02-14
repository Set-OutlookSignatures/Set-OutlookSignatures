<#
This sample code shows how to make it easier to use simulation mode for template creators.

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers commercial support for this and other open source code.
#>

[CmdletBinding()] param ()

Write-Host 'Set-OutlookSignatures simulation mode helper'


# Admin part
## Define parameters to use for Set-OutlookSignatures
## It is sufficient to only list parameters where the default value should not be used
$params = [ordered]@{
	SignatureTemplatePath         = '.\sample templates\Signatures DOCX'
	SignatureIniFile              = '.\sample templates\Signatures DOCX\_Signatures.ini'
	OOFTemplatePath               = '.\sample templates\Out-of-office DOCX'
	OOFIniFile                    = '.\sample templates\Out-of-office DOCX\_OOF.ini'
	ReplacementVariableConfigFile = '.\config\default replacement variables.ps1'
	GraphConfigFile               = '.\config\default graph config.ps1'
	GraphOnly                     = $false
}


#
# Do not change anything from here on
#


if ($psISE) {
	Write-Host '  PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
	Write-Host '  Required features are not available in ISE. Exit.' -ForegroundColor Red
	exit 1
}

$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

Set-Location $PSScriptRoot

## User part

Write-Host
Write-Host '  Please enter the login name of the user to simulate'
Write-Host '    Allowed formats:'
Write-Host '      user.x@example.com (UPN, a.k.a. User Principal Name)'
Write-Host '      EXAMPLE\User (pre-Windows 2000 logon name)'

do {
	$tempSimulateUser = Read-Host '    Your input'
} until (
	$(
		$tempSimulateUser = $tempSimulateUser.trim()
		if ($tempSimulateUser -match '^\S+@\S+$|^\S+\\\S+$') {
			Write-Host "      Simulate user: $($tempSimulateUser)"
			$params['SimulateUser'] = $tempSimulateUser
			$true
		} else {
			Write-Host '      Wrong format. Please try again.' -ForegroundColor yellow
		}
	)
)


Write-Host
Write-Host '  Please enter the email addresses of the mailboxes to simulate'
Write-Host '    Separate multiple mailboxes by spaces, commas or semicolons'
Write-Host '    Leave empty to get mailboxes from Outlook Web'
Write-Host '      Example: user.x@domain.com, user.a@domain.com, sharedmailbox.y@domain.com'

do {
	$tempSimulateMailboxes = Read-Host '    Your input'
} until (
	$(
		try {
			[mailaddress[]] $tempSimulateMailboxes = @(@(($tempSimulateMailboxes -replace '\s+', ',' -replace ';+', ',' -replace ',+', ',') -split ',') | Where-Object { $_ })
			Write-Host "      Simulate mailboxes: $($tempSimulateMailboxes -join ', ')"
			$params['SimulateMailboxes'] = $tempSimulateMailboxes
			$true
		} catch {
			Write-Host '      Wrong format. Please try again.' -ForegroundColor yellow
			$false
		}
	)
)




Write-Host
Write-Host '  Please enter the time use for simulation mode'
Write-Host '    Keep blank to use current date and time'
Write-Host '    Input must be in the international format yyyyMMddHHmm'
Write-Host '      yyyy = year in 4 digits'
Write-Host '      MM = month in 2 digits'
Write-Host '      dd = day in 2 digits'
Write-Host '      HH = hour in 2 digits, using 24-hour-format'
Write-Host '      mm = minute in 2 digits'
Write-Host '    Examples:'
Write-Host "      202303152249 is March 15th 2023 at 22:49 o'clock (22:49 is 10:49 p.m.)"

do {
	$tempTime = Read-Host '    Your input'
} until (
	$(
		if ($tempTime) {
			try {
				[DateTime]::ParseExact($tempTime, 'yyyyMMddHHmm', $null)
				Write-Host "      In local time: $([DateTime]::ParseExact($tempTime, 'yyyyMMddHHmm', $null))"
				$params['SimulateTime'] = $tempTime
				$true
			} catch {
				Write-Host '      Time is not valid. Please try again.' -ForegroundColor yellow
				$false
			}
		} else {
			$true
		}
	)
)


Write-Host
Write-Host '  Please enter the file path to use for simulation mode'
Write-Host '    The folder must already exist'
Write-Host '      Examples:'
Write-Host '        c:\users\userx\documents\Set-OutlookSignatures simulation folder'
Write-Host '        \\server\share\folder'
Write-Host '        https://server.example.com/site/library/folder'

do {
	$tempPath = Read-Host '    Your input'
} until (
	$(
		try {
			$tempPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($tempPath)
			if (Test-Path $tempPath) {
				Write-Host "      Path: $($tempPath)"
				$params['AdditionalSignaturePath'] = $tempPath
				$true
			} else {
				throw
			}
		} catch {
			Write-Host '      Folder does not exist. Please try again.' -ForegroundColor yellow
			$false
		}
	)
)


$paramsString = @(
	$params.GetEnumerator() | Where-Object { $_.Value } | ForEach-Object {
		if ($_.Value -is [array]) {
			"-$($_.Name) '$($_.Value -join "', '")'"
		} elseif ($_.value.tostring().startswith("'") -or $_.value.tostring().startswith('"')) {
			"-$($_.Name) $($_.Value)"
		} else {
			"-$($_.Name) '$($_.Value)'"
		}
	}
) -join ' '


Write-Host
Write-Host 'Resulting commands'
Write-Host '  For starting Set-OutlookSignatures directly from within this script'
Write-Host '    & ..\Set-OutlookSignatures.ps1 @params'
Write-Host '  For the command line'
Write-Host "    powershell.exe -command `"& ..\Set-OutlookSignatures.ps1 $($paramsString)`""


Write-Host
Write-Host 'Thank you for using Set-OutlookSignatures!'
