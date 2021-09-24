<#
.SYNOPSIS
    Add client application to local session cache.
.DESCRIPTION
    This cmdlet will add a client application object to the local session cache.
.EXAMPLE
    PS C:\>Add-MsalClientApplication $ClientApplication
    Add client application to the local session cache.
.EXAMPLE
    PS C:\>Add-MsalClientApplication $ClientApplication -PassThru
    Add client application to the local session cache and return application object.
#>
function Add-MsalClientApplication {
    [CmdletBinding(DefaultParameterSetName = 'PublicClient')]
    [OutputType([Microsoft.Identity.Client.PublicClientApplication], [Microsoft.Identity.Client.ConfidentialClientApplication])]
    param
    (
        # Public client application
        [Parameter(Mandatory = $true, ParameterSetName = 'PublicClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IPublicClientApplication] $PublicClientApplication,
        # Confidential client application
        [Parameter(Mandatory = $true, ParameterSetName = 'ConfidentialClient', Position = 0, ValueFromPipeline = $true)]
        [Microsoft.Identity.Client.IConfidentialClientApplication] $ConfidentialClientApplication,
        # Returns client application
        [Parameter(Mandatory = $false)]
        [switch] $PassThru
    )

    switch ($PSCmdlet.ParameterSetName) {
        "PublicClient" {
            $ClientApplication = $PublicClientApplication
            if (!$PublicClientApplications.Contains($ClientApplication)) {
                $PublicClientApplications.Add($ClientApplication)
            }
            else {
                $Exception = New-Object ArgumentException -ArgumentList 'The client application provided already exists in the session cache.'
                Write-Error -Exception $Exception -Category ([System.Management.Automation.ErrorCategory]::ResourceExists) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'AddMsalClientApplicationFailureAlreadyExists' -TargetObject $ClientApplication #-ErrorAction Stop
            }
            break
        }
        "ConfidentialClient" {
            $ClientApplication = $ConfidentialClientApplication
            if (!$ConfidentialClientApplications.Contains($ClientApplication)) {
                $ConfidentialClientApplications.Add($ClientApplication)
            }
            else {
                $Exception = New-Object ArgumentException -ArgumentList 'The client application provided already exists in the session cache.'
                Write-Error -Exception $Exception -Category ([System.Management.Automation.ErrorCategory]::ResourceExists) -CategoryActivity $MyInvocation.MyCommand -ErrorId 'AddMsalClientApplicationFailureAlreadyExists' -TargetObject $ClientApplication #-ErrorAction Stop
            }
            break
        }
    }

    if ($PassThru) {
        Write-Output $ClientApplication
    }
}

# SIG # Begin signature block
# MIIZrAYJKoZIhvcNAQcCoIIZnTCCGZkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD2Zzm1mh2ED3uu
# OIu2Noev6F5H9gAHJbYgxVqyVx1vlKCCFJUwggT+MIID5qADAgECAhANQkrgvjqI
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA8cqF7VfRK
# Rxf+6rdwPoI5ZjV290dZpDSUnD8KPsMhPDANBgkqhkiG9w0BAQEFAASCAQA0xIee
# /9PMbvMFO/KQQnrmxEjb6EnIltFmNbHLqLLRj/Ht3UTFH/nZm4JqWj8YBK4CbyZr
# B6wTTmE6ox41azLQ8gFx+24DVjWGYG3mrElRIbgGSzs5mM+5cZswGmaBG8huylUk
# VvPBze+N2BMonS4gW7s8d9jHUR373ZqFzkdWlYHzz/GxmN4RxPjy1rbuc0KEfrxR
# brULY8CgEbheVG5/a/9J8ABEelytKdBXnOqFeJzUmYnmqe28/4ql+HgZARKUx+C6
# SvLsgY3M+LtgS35qThKcr4MoV80haVyl09L8IKDkV6fzjpSHXs5CaKBaHbW2f5Ir
# c8j7LFh6zcxouTZJoYICMDCCAiwGCSqGSIb3DQEJBjGCAh0wggIZAgEBMIGGMHIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJ
# RCBUaW1lc3RhbXBpbmcgQ0ECEA1CSuC+Ooj/YEAhzhQA8N0wDQYJYIZIAWUDBAIB
# BQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0y
# MTA5MjExMzI3NTJaMC8GCSqGSIb3DQEJBDEiBCD1sG70vnhW8244MSK1X0iBNEGS
# MMOPm0FlT8sB861BsTANBgkqhkiG9w0BAQEFAASCAQDCA6hNPJbGKJeOSDj3P8hS
# ui+JAOF2Bpx60opfxI/oy236vXPMMa4xSLRc/Preodf96a6NYu88n1MSNHDOclc/
# Ng4p8/7700nIWsdKAsfRILQaFEUh0BCcRb/plQ0tr0C1URJxwqJUl0ZN0M1TFY8L
# mfna8nUHwqru3eP/VSVFy2c3fq8R/jUdliH4j7MNiHBMz1HnUjuPCRvMVLU/FEdO
# /PFYw43crLWA6LL33j+khzX0vgGdBs3XB7Vh1D/0lgeyW6nG9mRtao2W9+Amav8M
# pvVrmj2vQCGjoF9K0geLIYVpWabijyyLX2XICAzBa+rSq8B8iMESpsnnIbOFoI7k
# SIG # End signature block
