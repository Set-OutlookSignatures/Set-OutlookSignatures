<#
This sample code is used to check AD trusts and AD connectivity from a client computer.

Connection ist tested for every combination of
- DNS name of domain and domain controllers
- IP address of domain and domain controllers
- Protocols LDAP and GC, with and without encryption

This script assumes that the trust to check is either a cross-forest trust, or that the trusted domain is the only domain in it's forest

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.
#>

[CmdletBinding()]

param (
    [string[]]$CrossForestTrustRootDomains = @('example.com'),
    [string[]]$DcPorts = @(88, 389, 636, 9389),
    [string[]]$GcPorts = @(3268, 3269),
    [string[]]$DcProtocols = @('LDAP'),
    [string[]]$GcProtocols = @('GC'),
    [int]$JobsParallel = 10,
    [int]$JobTimeoutSeconds = 1200
)


Clear-Host


try {
    Write-Host "Start script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Write-Host '  Ports ususally required for LDAP and Global Catalog communication:'
    Write-Host '    88 TCP/UDP (Kerberos authentication)'
    Write-Host '    389 TCP/UPD (LDAP)'
    Write-Host '    636 TCP (LDAPS)'
    Write-Host '    3268 TCP (Global Catalog)'
    Write-Host '    3269 TCP (Global Catalog TLS)'
    Write-Host '    9389 TCP (Active Directory Web Services)'
    Write-Host '    49152-65535 TCP (high ports)'
    Write-Host '  DNS name resolution must work flawlessly, too.'


    Write-Host
    Write-Host "Check parameters and script environment @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    
    if ($psISE) {
        Write-Host '  PowerShell ISE detected. Use PowerShell in console or terminal instead.' -ForegroundColor Red
        Write-Host '  Required features are not available in ISE. Exit.' -ForegroundColor Red
        exit 1
    }
        
    $OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding

    Set-Location $PSScriptRoot

    Write-Host "  PowerShell: '$((($($PSVersionTable.PSVersion), $($PSVersionTable.PSEdition), $($PSVersionTable.Platform), $($PSVersionTable.OS)) | Where-Object {$_}) -join "', '")'"

    Write-Host "  PowerShell bitness: $(if ([Environment]::Is64BitProcess -eq $false) {'Non-'})64-bit process on a $(if ([Environment]::Is64OperatingSystem -eq $false) {'Non-'})64-bit operating system"

    Write-Host '  Parameters'
    foreach ($parameter in (Get-Command -Name $PSCommandPath).Parameters.keys) {
        if ((Get-Variable -Name $parameter -EA SilentlyContinue -ValueOnly) -is [hashtable]) {
            Write-Host "    $($parameter): '$(@((Get-Variable -Name $parameter -ValueOnly).GetEnumerator() | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ', ')'"
        } else {
            Write-Host "    $($parameter): '$((Get-Variable -Name $parameter -EA SilentlyContinue -ValueOnly) -join ', ')'"
        }
    }

    Write-Host "  Script path: '$PSCommandPath'"

    if ($IsWindows -or (-not (Test-Path 'variable:IsWindows'))) {
    } else {
        Write-Host "  Your OS: $($PSVersionTable.OS)" -ForegroundColor Red
        Write-Host '  This script is supported on Windows only. Exit.' -ForegroundColor Red
    }

    if (($ExecutionContext.SessionState.LanguageMode) -ine 'FullLanguage') {
        Write-Host "  This PowerShell session runs in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode." -ForegroundColor Red
        Write-Host '  Required features are only available in FullLanguage mode. Exit.' -ForegroundColor Red
        exit 1
    }

    foreach ($CrossForestTrustRootDomain in $CrossForestTrustRootDomains) {
        Write-Host
        Write-Host "$CrossForestTrustRootDomain"

        Write-Host "  Check forest root domain via LDAP @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $ADForestRootDomain = ([ADSI]"LDAP://$($CrossForestTrustRootDomain)/RootDSE").rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')

        if (-not $ADForestRootDomain) {
            Write-Host "    Could not connect to '$($CrossForestTrustRootDomain)' via LDAP to query RootDSE. Skipping."
            continue
        }

        if ($ADForestRootDomain -ine $CrossForestTrustRootDomain) {
            Write-Host "    '$($CrossForestTrustRootDomain)' is not the forest root domain, using '$($ADForestRootDomain)' from now on."
            $CrossForestTrustRootDomain = $ADForestRootDomain
        } else {
            Write-Host "    '$($CrossForestTrustRootDomain)' is the forest root domain, continue using this name."
        }


        Write-Host "  Get FQDN of all Global Catalog servers via DNS query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"

        $AllGCs = @((Resolve-DnsName -Name "_gc._tcp.$($CrossForestTrustRootDomain)" -Type srv).nametarget)

        Write-Host "    $($AllGCs.count) found"

        if ($AllGCs.count -lt 1) {
            Write-Host '    No Global Catalog servers found. Check input and DNS resolution. Skipping.'
            continue
        }


        Write-Host "  Get FQDN of all Domain Controller servers via DNS query @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        $AllDCs = @()

        foreach ($DomainName in @(@(foreach ($GC in $AllGCs) { ($GC -split '\.', 2)[1] }) | Select-Object -Unique)) {
            $AllDCs += (Resolve-DnsName -Name "_ldap._tcp.$($DomainName)" -Type srv).nametarget
        }

        Write-Host "    $($AllDCs.count) found"

        $JobsParallel = [math]::min($JobsParallel, $AllDCs.count)


        Write-Host "  Testing server connectivity ($($JobsParallel) in parallel) @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
        Write-Host '    This can take very long due to long, non-configurable timeouts.'

        $script:jobs = New-Object System.Collections.ArrayList

        [void][runspacefactory]::CreateRunspacePool()
        $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $JobsParallel)
        $RunspacePool.Open()

        Write-Output '"Time";"Client";"Target forest/domain";"Target server";"Check";"DNS name or IP address";"Port";"Result";"Time in ms";"Error"'

        foreach ($DC in (@($AllDCs + $AllGCs) | Sort-Object -Culture 127 -Unique)) {
            $PowerShell = [powershell]::Create()
            $PowerShell.RunspacePool = $RunspacePool

            [void]$PowerShell.AddScript( {
                    Param (
                        $CrossForestTrustRootDomain,
                        $DC,
                        $IsGC,
                        $DcPorts,
                        $DcProtocols,
                        $GcPorts,
                        $GcProtocols
                    )

                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

                    $DebugPreference = 'Continue'

                    Write-Debug "  Start(Ticks) = $((Get-Date).Ticks)"

                    $Client = [System.Net.Dns]::GetHostByName($env:computerName).HostName

                    #Write-Output "    $($DC)"



                    if ($IsGC) {
                        #Write-Output '      Role: Domain Controller and Global Catalog'
                        $Ports = @($DcPorts) + @($GcPorts)
                        $Protocols = @($DcProtocols) + @($GcProtocols)
                    } else {
                        #Write-Output '      Role: Domain Controller only, no Global Catalog'
                        $Ports = @($DcPorts)
                        $Protocols = @($DcProtocols)
                    }


                    $IPs = @(([System.Net.Dns]::GetHostAddresses($DC)).IPAddressToString)
                    #Write-Output "      IP(s): $($IPs -join ', ')"

                    foreach ($Port in $Ports) {
                        Write-Output $('"' +
                            (
                                @(
                                    @(
                                        $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                        $($Client),
                                        $($CrossForestTrustRootDomain),
                                        $($DC),
                                        'Port via DNS',
                                        $($DC),
                                        $($Port),
                                        $($stopwatch.Restart(); (Test-NetConnection -ComputerName $DC -Port $Port -WarningAction silentlycontinue).TcpTestSucceeded),
                                        $($stopwatch.ElapsedMilliseconds),
                                        ''
                                    ) | ForEach-Object { $_ -ireplace '"', '""' }
                                ) -join '";"'
                            ) + '"'
                        )

                        foreach ($IP in $IPs) {
                            Write-Output $('"' +
                                (
                                    @(
                                        @(
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            $($Client),
                                            $($CrossForestTrustRootDomain),
                                            $($DC),
                                            'Port via IP',
                                            $($IP),
                                            $($Port),
                                            $($stopwatch.Restart(); (Test-NetConnection -ComputerName $IP -Port $Port -WarningAction silentlycontinue).TcpTestSucceeded),
                                            $($stopwatch.ElapsedMilliseconds),
                                            ''
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }
                                    ) -join '";"'
                                ) + '"'
                            )
                        }
                    }


                    foreach ($Protocol in $Protocols) {
                        $Search = New-Object DirectoryServices.DirectorySearcher
                        $Search.PageSize = 1000
                        $Search.filter = '(objectclass=user)'

                        try {
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($Protocol)://$DC")
                            $stopwatch.Restart()
                            $UserAccount = [ADSI]"$(($Search.FindOne()).path)"
                            Write-Output $('"' +
                                (
                                    @(
                                        @(
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            $($Client),
                                            $($CrossForestTrustRootDomain),
                                            $($DC),
                                            'LDAP/GC Query via DNS',
                                            $($DC),
                                            $($Protocol),
                                            $true,
                                            $($stopwatch.ElapsedMilliseconds),
                                            ''
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }
                                    ) -join '";"'
                                ) + '"'
                            )
                        } catch {
                            Write-Output $('"' +
                                (
                                    @(
                                        @(
                                            $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                            $($Client),
                                            $($CrossForestTrustRootDomain),
                                            $($DC),
                                            'LDAP/GC Query via DNS',
                                            $($DC),
                                            $($Protocol),
                                            $false,
                                            $($stopwatch.ElapsedMilliseconds),
                                            $($Error[0])
                                        ) | ForEach-Object { $_ -ireplace '"', '""' }
                                    ) -join '";"'
                                ) + '"'
                            )
                        }

                        foreach ($IP in $IPs) {
                            try {
                                $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("$($Protocol)://$IP")
                                $stopwatch.Restart();
                                $UserAccount = [ADSI]"$(($Search.FindOne()).path)"
                                Write-Output $('"' +
                                    (
                                        @(
                                            @(
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                $($Client),
                                                $($CrossForestTrustRootDomain),
                                                $($DC),
                                                'LDAP/GC Query via IP',
                                                $($IP),
                                                $($Protocol),
                                                $true,
                                                $($stopwatch.ElapsedMilliseconds),
                                                ''
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }
                                        ) -join '";"'
                                    ) + '"'
                                )
                            } catch {
                                Write-Output $('"' +
                                    (
                                        @(
                                            @(
                                                $(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK'),
                                                $($Client),
                                                $($CrossForestTrustRootDomain),
                                                $($DC),
                                                'LDAP/GC Query via IP',
                                                $($IP),
                                                $($Protocol),
                                                $false,
                                                $($stopwatch.ElapsedMilliseconds),
                                                $($Error[0])
                                            ) | ForEach-Object { $_ -ireplace '"', '""' }
                                        ) -join '";"'
                                    ) + '"'
                                )
                            }
                        }
                    }


                    $stopwatch.Stop()
                }).AddParameters(
                @{
                    CrossForestTrustRootDomain = $CrossForestTrustRootDomain
                    DC                         = $DC
                    IsGC                       = ($DC -iin $AllGCs)
                    DcPorts                    = $DcPorts
                    DcProtocols                = $DcProtocols
                    GcPorts                    = $GcPorts
                    GcProtocols                = $GcProtocols
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
    }
} catch {
    Write-Host $error[0]
    Write-Host
    Write-Host "Unknown error, exiting. @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
    Exit 1

} finally {
    Write-Host
    Write-Host "End script @$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')@"
}
