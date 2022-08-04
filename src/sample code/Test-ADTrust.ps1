# This script assumes that the trust to check is either a cross-forest trust,
# or that the trusted domain is the only domain in it's forest


param (
    [string]$CrossForestTrustRootDomain = 'example.com',
    [string[]]$DcPorts = (88, 389, 636),
    [string[]]$GcPorts = (3268, 3269),
    [string[]]$DcProtocols = ('LDAP'),
    [string[]]$GcProtocols = ('GC'),
    [int]$JobsParallel = 10,
    [int]$JobTimeoutSeconds = 1200
)


Clear-Host


try {
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host '  Ports ususally required for LDAP and Global Catalog communication:'
    Write-Host '    88 TCP/UDP (Kerberos authentication)'
    Write-Host '    389 TCP/UPD (LDAP)'
    Write-Host '    636 TCP (LDAP SSL)'
    Write-Host '    3268 TCP (Global Catalog)'
    Write-Host '    3269 TCP (Global Catalog SSL)'
    Write-Host '    49152-65535 TCP (high ports)'
    Write-Host '  DNS name resolution must work flawlessly, too.'


    Write-Host
    Write-Host "Check parameters and script environment @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    Write-Host "  PowerShell: '$((($($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition), $($PSVersionTable.Platform), $($PSVersionTable.OS)) | Where-Object {$_}) -join "', '")'"

    Write-Host "  PowerShell bitness: $(if ([Environment]::Is64BitProcess -eq $false) {'Non-'})64-bit process on a $(if ([Environment]::Is64OperatingSystem -eq $false) {'Non-'})64-bit operating system"

    Write-Host '  Parameters'
    foreach ($parameter in (Get-Command -Name $PSCommandPath).Parameters.keys) {
        Write-Host "    $($parameter): " -NoNewline

        if ((Get-Variable -Name $parameter -EA SilentlyContinue -ValueOnly) -is [hashtable]) {
            Write-Host "'$(@((Get-Variable -Name $parameter -ValueOnly).GetEnumerator() | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ', ')'"
        } else {
            Write-Host "'$((Get-Variable -Name $parameter -EA SilentlyContinue -ValueOnly) -join ', ')'"
        }
    }

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


    Write-Host
    Write-Host "Check forest root domain via LDAP @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $ADForestRootDomain = ([ADSI]"LDAP://$($CrossForestTrustRootDomain)/RootDSE").rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')

    if (-not $ADForestRootDomain) {
        Write-Host "  Could not connect to '$($CrossForestTrustRootDomain)' via LDAP to query RootDSE. Exiting."
        exit 1
    }

    if ($ADForestRootDomain -ine $CrossForestTrustRootDomain) {
        Write-Host "  '$($CrossForestTrustRootDomain)' is not the forest root domain, using '$($ADForestRootDomain)' from now on."
        $CrossForestTrustRootDomain = $ADForestRootDomain
    } else {
        Write-Host "  '$($CrossForestTrustRootDomain)' is the forest root domain, continue using this name."
    }


    Write-Host
    Write-Host "Get FQDN of all Global Catalog servers via DNS query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"

    $AllGCs = @((Resolve-DnsName -Name "_gc._tcp.$($CrossForestTrustRootDomain)" -Type srv).nametarget)

    Write-Host "  $($AllGCs.count) found"

    if ($AllGCs.count -lt 1) {
        Write-Host '  No Global Catalog servers found. Check input and DNS resolution. Exiting.'
        exit 1
    }


    Write-Host
    Write-Host "Get FQDN of all Domain Controller servers via DNS query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    $AllDCs = @()

    foreach ($DomainName in @(@(foreach ($GC in $AllGCs) { ($GC -split '\.', 2)[1] }) | Select-Object -Unique)) {
        $AllDCs += (Resolve-DnsName -Name "_ldap._tcp.$($DomainName)" -Type srv).nametarget
    }

    Write-Host "  $($AllDCs.count) found"

    $JobsParallel = [math]::min($JobsParallel, $AllDCs.count)


    Write-Host
    Write-Host "Testing server connectivity ($($JobsParallel) in parallel) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Write-Host '  This can take very long due to long, non-configurable timeouts.'

    $script:jobs = New-Object System.Collections.ArrayList

    [void][runspacefactory]::CreateRunspacePool()
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $JobsParallel)
    $RunspacePool.Open()

    foreach ($DC in (@($AllDCs + $AllGCs) | Sort-Object -Unique)) {
        $PowerShell = [powershell]::Create()
        $PowerShell.RunspacePool = $RunspacePool

        [void]$PowerShell.AddScript( {
                Param (
                    $DC,
                    $IsGC,
                    $DcPorts,
                    $DcProtocols,
                    $GcPorts,
                    $GcProtocols
                )


                $DebugPreference = 'Continue'

                Write-Debug "Start(Ticks) = $((Get-Date).Ticks)"


                Write-Output "  $($DC)"



                if ($IsGC) {
                    Write-Output '    Role: Domain Controller and Global Catalog'
                    $Ports = @($DcPorts) + @($GcPorts)
                    $Protocols = @($DcProtocols) + @($GcProtocols)
                } else {
                    Write-Output '    Role: Domain Controller only, no Global Catalog'
                    $Ports = @($DcPorts)
                    $Protocols = @($DcProtocols)
                }


                $IPs = @(([System.Net.Dns]::GetHostAddresses($DC)).IPAddressToString)
                Write-Output "    IP(s): $($IPs -join ', ')"


                foreach ($Port in $Ports) {
                    Write-Output "    Port $($Port) via DNS name: $((Test-NetConnection -ComputerName $DC -Port $Port -WarningAction silentlycontinue).TcpTestSucceeded)"

                    foreach ($IP in $IPs) {
                        Write-Output "    Port $($Port) via IP $($IP): $((Test-NetConnection -ComputerName $IP -Port $Port -WarningAction silentlycontinue).TcpTestSucceeded)"
                    }
                }


                foreach ($Protocol in $Protocols) {
                    $Search = New-Object DirectoryServices.DirectorySearcher
                    $Search.PageSize = 1000
                    $Search.filter = '(objectclass=user)'

                    try {
                        $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($Protocol)://$DC")
                        $UserAccount = [ADSI]"$(($Search.FindOne()).path)"
                        Write-Output "    $($Protocol) query via DNS: True"
                    } catch {
                        Write-Output "    $($Protocol) query via DNS: False"
                        #Write-Output "      $($Error[0])"
                    }

                    foreach ($IP in $IPs) {
                        try {
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($Protocol)://$IP")
                            $UserAccount = [ADSI]"$(($Search.FindOne()).path)"
                            Write-Output "    $($Protocol) query via IP $($IP): True"
                        } catch {
                            Write-Output "    $($Protocol) query via IP $($IP): False"
                            #Write-Output "      $($Error[0])"
                        }
                    }
                }
            }).AddParameters(
            @{
                DC          = $DC
                IsGC        = ($DC -iin $AllGCs)
                DcPorts     = $DcPorts
                DcProtocols = $DcProtocols
                GcPorts     = $GcPorts
                GcProtocols = $GcProtocols
            }
        )

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
                if ((($job.handle.IsCompleted -eq $true) -and ($job.Done -eq $false)) -or (($job.Done -eq $false) -and ((New-TimeSpan -Start $job.StartTime -End (Get-Date)).TotalSeconds -gt $JobTimeoutSeconds))) {
                    $job.object
                    $job.Done = $true
                }
            }
        }
    }
} catch {
    Write-Host
    $error[0]
    Write-Host "Unknown error, exiting. @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
    Exit 1

} finally {
    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:sszzz')@"
}
