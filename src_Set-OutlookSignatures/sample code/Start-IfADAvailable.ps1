<#
This sample code shows how to start Set-OutlookSignatures only when Active Directory can be reached.

It covers the following cases:
  - At least one DC from the user's domain is reachable
  - At least one Global Catalog server from the user's domain is reachable via a GC query
  - The querying user exists and is not locked
  - All domains in the user's forest are reachable via LDAP and GC queries

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.
#>


$testIntervalSeconds = 5 # Interval between retries
$testTimeoutSeconds = 120 # For how long to retry (in seconds) before giving up

#Requires -Version 5.1

Write-Host 'Start AD connectivity test'

Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$testCurrentUserDN = ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).DistinguishedName

if (
  $($null -eq $testCurrentUserDN) -or
  $(($testCurrentUserDN -split ',DC=').Count -lt 3)
) {
  Write-Host '  User is not a member of a domain, do not go on with further tests.'
} else {
  $testStartTime = Get-Date
  $testSuccess = $false

  do {
    if (Test-Connection $(($testCurrentUserDN -split ',DC=')[1..999] -join '.') -Count 1 -Quiet) {
      Write-Host '  User on-prem AD can be reached, perform test query against AD.'

      $testCurrentUserADProps = $null

      try {
        $testSearch = New-Object DirectoryServices.DirectorySearcher
        $testSearch.PageSize = 1000
        $testSearch.SearchRoot = "GC://$(($testCurrentUserDN -split ',DC=')[1..999] -join '.')"
        $testSearch.Filter = "((distinguishedname=$($testCurrentUserDN)))"

        $testCurrentUserADProps = $testSearch.FindOne().Properties
      } catch {
        $testCurrentUserADProps = $null
      }

      if ($null -ne $testCurrentUserADProps) {
        Write-Host '  AD query was successful, user is not locked, DC is reachable via GC query: Start Set-OutlookSignatures.'

        # Get all domains of the current user forest, as they must be reachable, too
        Write-Host '  Testing child domains'

        $testCurrentUserForest = (([ADSI]"LDAP://$(($testCurrentUserDN -split ',DC=')[1..999] -join '.')/RootDSE").rootDomainNamingContext -ireplace [Regex]::Escape('DC='), '' -ireplace [Regex]::Escape(','), '.').tolower()

        $testSearch.SearchRoot = "GC://$($testCurrentUserForest)"
        $testSearch.Filter = '(ObjectClass=trustedDomain)'
        $testTrustedDomains = @($testSearch.FindAll())

        $testTrustedDomains = @(
          @() +
          $testCurrentUserForest +
          @(
            @($testTrustedDomains) | Where-Object { (($_.properties.trustattributes -eq 32) -and ($_.properties.name -ine $testCurrentUserForest)) }
          ).properties.name
        ) | Select-Object -Unique

        $testTrustedDomainFailCount = 0

        foreach ($testTrustedDomain in $testTrustedDomains) {
          if ($testTrustedDomainFailCount -gt 0) {
            break
          }

          Write-Host "    $($testTrustedDomain)"

          foreach ($CheckProtocolText in @('LDAP', 'GC')) {
            if ($testTrustedDomainFailCount -gt 0) {
              break
            }

            $testSearch.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($CheckProtocolText)://$testTrustedDomain")
            $testSearch.filter = '(objectclass=user)'

            try {
              $null = ([ADSI]"$(($testSearch.FindOne()).path)")

              Write-Host "      $($CheckProtocolText): Passed"
            } catch {
              $testTrustedDomainFailCount++

              Write-Host "      $($CheckProtocolText): Failed"
            }
          }
        }

        if ($testTrustedDomainFailCount -eq 0) {
          $testSuccess = $true
        }

        #
        # Start Set-OutlookSignatures here
        #
      } else {
        Write-Host '  AD query failed, user might be locked or DCs can not be reached via GC query: Do not start Set-OutlookSignatures.'
      }

    } else {
      Write-Host '  User on-prem AD can not be reached, do not go on with further tests.'
    }

    if ($testSuccess -ne $true) {
      $testElapsedSeconds = [math]::Ceiling((New-TimeSpan -Start $testStartTime).TotalSeconds)

      if ($testElapsedSeconds -ge $testTimeoutSeconds) {
        Write-Host "  Timeout reached ($($testTimeoutSeconds) seconds). Tests stopped."
        break
      } else {
        Write-Host "  Retrying in $($testIntervalSeconds) seconds. $($testTimeoutSeconds - $testElapsedSeconds) seconds left until timeout."
        Start-Sleep -Seconds $testIntervalSeconds
      }
    }
  } while ($testSuccess -ne $true)
}