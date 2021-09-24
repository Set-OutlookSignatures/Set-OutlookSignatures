param (
    # Provide module configuration
    [Parameter(Mandatory = $false)]
    [psobject] $ModuleConfiguration
)

## Set Strict Mode for Module. https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode
Set-StrictMode -Version 3.0

#Write-Warning 'The MSAL.PS PowerShell module wraps MSAL.NET functionality into PowerShell-friendly cmdlets and is not supported by Microsoft. Microsoft support does not extend beyond the underlying MSAL.NET library. For any inquiries regarding the PowerShell module itself, you may contact the author on GitHub or PowerShell Gallery.'

$script:ModuleConfigDefault = Import-Config -Path (Join-Path $PSScriptRoot 'config.json')
$script:ModuleConfig = $script:ModuleConfigDefault.psobject.Copy()
Import-Config | Set-Config
Set-Config -ResolveFromEnvironmentVariables
if ($PSBoundParameters.ContainsKey('ModuleConfiguration')) { Set-Config $ModuleConfiguration }
#Export-Config

$script:ModuleFeatureSupport = [ordered]@{
    WebView1Support        = $PSVersionTable.PSEdition -eq 'Desktop'
    WebView2Support        = [System.Environment]::OSVersion.Platform -eq 'Win32NT' -and [System.Environment]::Is64BitProcess -and ($PSVersionTable.PSVersion -lt [version]'6.0' -or $PSVersionTable.PSVersion -ge [version]'7.0' -and (Get-Item 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}' -ErrorAction SilentlyContinue))
    #EmbeddedWebViewSupport = $WebView1Support -or $WebView2Support
    DeviceCodeSupport      = $true
    TokenCacheSupport      = [System.Environment]::OSVersion.Platform -eq 'Win32NT' -and $PSVersionTable.PSVersion -lt [version]'6.0'
    AuthBrokerSupport      = [System.Environment]::OSVersion.Platform -eq 'Win32NT' -and $PSVersionTable.PSVersion -lt [version]'7.0'
}

## Get Device Registration Status
[hashtable] $Dsreg = @{}
#if ([System.Environment]::OSVersion.Platform -eq 'Win32NT' -and [System.Environment]::OSVersion.Version -ge '10.0') {
    try {
        Dsregcmd /status | foreach { if ($_ -match '\s*(.+) : (.+)') { $Dsreg.Add($Matches[1], $Matches[2]) } }
    }
    catch {}
#}

## PowerShell Desktop 5.1 does not dot-source ScriptsToProcess when a specific version is specified on import. This is a bug.
# if ($PSEdition -eq 'Desktop') {
#     $ModuleManifest = Import-PowershellDataFile (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name.Replace('.psm1','.psd1'))
#     if ($ModuleManifest.ContainsKey('ScriptsToProcess')) {
#         foreach ($Path in $ModuleManifest.ScriptsToProcess) {
#             . (Join-Path $PSScriptRoot $Path)
#         }
#     }
# }

## Azure Automation module import fails when ScriptsToProcess is specified in manifest. Referencing import script directly.
. (Join-Path $PSScriptRoot $MyInvocation.MyCommand.Name.Replace('.psm1', '.ps1'))

## Global Variables
[System.Collections.Generic.List[Microsoft.Identity.Client.IPublicClientApplication]] $PublicClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IPublicClientApplication]'
[System.Collections.Generic.List[Microsoft.Identity.Client.IConfidentialClientApplication]] $ConfidentialClientApplications = New-Object 'System.Collections.Generic.List[Microsoft.Identity.Client.IConfidentialClientApplication]'
$script:ModuleState = @{
    UseWebView2 = $script:ModuleFeatureSupport.WebView2Support -and ($Dsreg['AzureAdPrt'] -eq 'NO' -or !$script:ModuleFeatureSupport.WebView1Support)
}

# SIG # Begin signature block
# MIIZrAYJKoZIhvcNAQcCoIIZnTCCGZkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAv32wBX0OE33Lg
# BSBsVpAbiFbGerhG+Er5SfMyPc0Gd6CCFJUwggT+MIID5qADAgECAhANQkrgvjqI
# /2BAIc4UAPDdMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3RhbXBpbmcgQ0EwHhcN
# MjEwMTAxMDAwMDAwWhcNMzEwMTA2MDAwMDAwWjBIMQswCQYDVQQGEwJVUzEXMBUG
# A1UEChMORGlnaUNlcnQsIEluYy4xIDAeBgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFt
# cCAyMDIxMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwuZhhGfFivUN
# CKRFymNrUdc6EUK9CnV1TZS0DFC1JhD+HchvkWsMlucaXEjvROW/m2HNFZFiWrj/
# ZwucY/02aoH6KfjdK3CF3gIY83htvH35x20JPb5qdofpir34hF0edsnkxnZ2OlPR
# 0dNaNo/Go+EvGzq3YdZz7E5tM4p8XUUtS7FQ5kE6N1aG3JMjjfdQJehk5t3Tjy9X
# tYcg6w6OLNUj2vRNeEbjA4MxKUpcDDGKSoyIxfcwWvkUrxVfbENJCf0mI1P2jWPo
# GqtbsR0wwptpgrTb/FZUvB+hh6u+elsKIC9LCcmVp42y+tZji06lchzun3oBc/gZ
# 1v4NSYS9AQIDAQABo4IBuDCCAbQwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQC
# MAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQQYDVR0gBDowODA2BglghkgBhv1s
# BwEwKTAnBggrBgEFBQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMB8G
# A1UdIwQYMBaAFPS24SAd/imu0uRhpbKiJbLIFzVuMB0GA1UdDgQWBBQ2RIaOpLqw
# Zr68KC0dRDbd42p6vDBxBgNVHR8EajBoMDKgMKAuhixodHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vc2hhMi1hc3N1cmVkLXRzLmNybDAyoDCgLoYsaHR0cDovL2NybDQu
# ZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC10cy5jcmwwgYUGCCsGAQUFBwEBBHkw
# dzAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME8GCCsGAQUF
# BzAChkNodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRTSEEyQXNz
# dXJlZElEVGltZXN0YW1waW5nQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4IBAQBIHNy1
# 6ZojvOca5yAOjmdG/UJyUXQKI0ejq5LSJcRwWb4UoOUngaVNFBUZB3nw0QTDhtk7
# vf5EAmZN7WmkD/a4cM9i6PVRSnh5Nnont/PnUp+Tp+1DnnvntN1BIon7h6JGA078
# 9P63ZHdjXyNSaYOC+hpT7ZDMjaEXcw3082U5cEvznNZ6e9oMvD0y0BvL9WH8dQgA
# dryBDvjA4VzPxBFy5xtkSdgimnUVQvUtMjiB2vRgorq0Uvtc4GEkJU+y38kpqHND
# Udq9Y9YfW5v3LhtPEx33Sg1xfpe39D+E68Hjo0mh+s6nv1bPull2YYlffqe0jmd4
# +TaY4cso2luHpoovMIIFJjCCBA6gAwIBAgIQCm8Gpkn9Nk686mPMJKDEczANBgkq
# hkiG9w0BAQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBT
# SEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDMzMTAwMDAwMFoX
# DTIzMDQwNTEyMDAwMFowYzELMAkGA1UEBhMCVVMxDTALBgNVBAgTBE9oaW8xEzAR
# BgNVBAcTCkNpbmNpbm5hdGkxFzAVBgNVBAoTDkphc29uIFRob21wc29uMRcwFQYD
# VQQDEw5KYXNvbiBUaG9tcHNvbjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
# ggEBAMVnygZO0wvpQ3NjGpEREqp0r/FN5C0X0Gn6HxrhPXAyGZaBlEjV0eO6bz8N
# BVFwyHsQ0BFxT7CrGvCCwvekm7bqIZaIJe9kFYAvOVBDK+S042dGaT8cUSxU6QIk
# gXL2IZKZu8R8H0+26rehGpadj+onbqzFshaS8C18/1oFv27W/3FeOwAkXbE8Mbpu
# c9ntR/6PUV4biw3AYUITVps0PmfTB1f06DmrbWa3orHVDO1yEL/E1hoe0jpXPAHz
# vtNlLMtZg5LeRrGdkfasq8V94XicNWU8XFy6D5cFlIg0RPcSzMJRJb78nfpQInrp
# DAagviDCUVR5ZwLsvDk096h8kCUCAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrE
# uXsqCqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBS+5845JPvDWenjXahLo4XUCcTn
# MjAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAw
# bjA1oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1j
# cy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFz
# c3VyZWQtY3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYB
# BQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGE
# BggrBgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQC
# MAAwDQYJKoZIhvcNAQELBQADggEBAER9rMHu+w+qJrQmh6at6GrAPYuHi2zuU04n
# dRRzTSmHUKvzS1DvEYxLp6cO//3gHEqBV1S0YV58Rn5idMii7fmANSfO1Og4x77/
# CmmnpwB8aoSCpbRxqcIBE+pUm7r7JBT4xNEKT3FkgcpVymE4VuIscBgnekEmmaVf
# Doh1Xm4cQ+hvtyZ8+3+bNQ/Oe008RSk5zmiWiS++eGeB1D5v6yLs2bHAHldKKCp8
# Mg322VqRB2C9bFlQSxS97FB/s4J4jGxjSSl6MmcYLzkw+Copc5/9c1QEzBe+9rZM
# aAPwb6e977tkFtFOCfiekESAjku2NPqjj83EtLOOllrv3r81oWcwggUwMIIEGKAD
# AgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0x
# MzEwMjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0Ew
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZ
# z9D7RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnko
# On7p0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXU
# LaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q
# 8grkV7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5S
# lsHyDxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPR
# AgMBAAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIB
# hjATBgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# dDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBG
# MDgGCmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNl
# cnQuY29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt
# 9mV1DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcN
# AQELBQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L
# /e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xG
# Tlz/kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGiv
# m6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZ
# Hen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdi
# bqFT+hKUGIUukpHqaGxEMrJmoecYpJpkUe8wggUxMIIEGaADAgECAhAKoSXW1jIb
# fkHkBdo2l8IVMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xNjAxMDcxMjAwMDBa
# Fw0zMTAxMDcxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lD
# ZXJ0IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3RhbXBpbmcgQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQC90DLuS82Pf92puoKZxTlUKFe2I0rEDgdFM1EQ
# fdD5fU1ofue2oPSNs4jkl79jIZCYvxO8V9PD4X4I1moUADj3Lh477sym9jJZ/l9l
# P+Cb6+NGRwYaVX4LJ37AovWg4N4iPw7/fpX786O6Ij4YrBHk8JkDbTuFfAnT7l3I
# mgtU46gJcWvgzyIQD3XPcXJOCq3fQDpct1HhoXkUxk0kIzBdvOw8YGqsLwfM/fDq
# R9mIUF79Zm5WYScpiYRR5oLnRlD9lCosp+R1PrqYD4R/nzEU1q3V8mTLex4F0IQZ
# chfxFwbvPc3WTe8GQv2iUypPhR3EHTyvz9qsEPXdrKzpVv+TAgMBAAGjggHOMIIB
# yjAdBgNVHQ4EFgQU9LbhIB3+Ka7S5GGlsqIlssgXNW4wHwYDVR0jBBgwFoAUReui
# r/SSy4IxLVGLp6chnfNtyA8wEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwUAYDVR0g
# BEkwRzA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4IBAQBx
# lRLpUYdWac3v3dp8qmN6s3jPBjdAhO9LhL/KzwMC/cWnww4gQiyvd/MrHwwhWiq3
# BTQdaq6Z+CeiZr8JqmDfdqQ6kw/4stHYfBli6F6CJR7Euhx7LCHi1lssFDVDBGiy
# 23UC4HLHmNY8ZOUfSBAYX4k4YU1iRiSHY4yRUiyvKYnleB/WCxSlgNcSR3CzddWT
# hZN+tpJn+1Nhiaj1a5bA9FhpDXzIAbG5KHW3mWOFIoxhynmUfln8jA/jb7UBJrZs
# pe6HUSHkWGCbugwtK22ixH67xCUrRwIIfEmuE7bhfEJCKMYYVs9BNLZmXbZ0e/VW
# MyIvIjayS6JKldj1po5SMYIEbTCCBGkCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEx
# MC8GA1UEAxMoRGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBD
# QQIQCm8Gpkn9Nk686mPMJKDEczANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA3F4pWYo7d
# viC14RrnJ9wkQUoW0Ckke5rlhZTvEzyVEDANBgkqhkiG9w0BAQEFAASCAQAof/Gq
# M6n6DUXqRmkIUUrHn6/W9rqod04+PI2mAmSIQ37n9zEDwOqHesPS8pcGeNv1r/Ds
# eciqM0tGamGtUf3Myd0+2HTvmFpIHPlIVVAS4Ouy0Z2rdD/pN2WAboc+zb7jdQ1y
# S3zSh8yPsFKPffEZKT+Qzi6FdF1e8GJH07x4BV1wz37vlK4IWJN2v7+TjXSuEgf2
# C39X0jG/epodTmpo3kKUQi27X20xFmYvAiRQvGnvs8KsDklwLqIcRv1xJmqReT+r
# jhmsBuwedZZW9QzECz7xFjf/s3nmeiESHOyI+o8YvCsd3R1eZ4I3b58K+s27FXJo
# UL9nbA+ax+qqvxrBoYICMDCCAiwGCSqGSIb3DQEJBjGCAh0wggIZAgEBMIGGMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBUaW1lc3RhbXBpbmcgQ0ECEA1CSuC+Ooj/YEAhzhQA8N0wDQYJYIZIAWUDBAIB
# BQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0y
# MTA5MjExMzI3NTNaMC8GCSqGSIb3DQEJBDEiBCD+266zuVQVGgs+SEzZFieag+jX
# GguflNCuNNNz505V4TANBgkqhkiG9w0BAQEFAASCAQBkO/YEgNqhi2tsAAQA5N6Y
# MYKJLBpf05gDHzlh6OjmGjRZZcguPvqDDL1krfZGiDuIs/csG8K9zz9A5xC+1n3l
# iS0wm6mcnMdDIIc7LF6Oc1ka4BRnYEJjEmrfDwqW1aPXbUmRcT7y7hrezAnHh8wH
# eVEpVBQBqHXOleGhaGYT/0fQMPRm2YZ9xHhwn0cUi2CWMkODARyBO5BKrmTlKc8M
# v8dnMmB/aw3QnbdLk8IL4xYpK34A0WS9h5jEr0JNx57qYHqZIjfaNa7vsHMnisBS
# RpUVJy/SljBXYhfKlNqDTrvohCAD6oeGkOdK4s5IjptIzC89VhuT4vMz06V4YCpf
# SIG # End signature block
