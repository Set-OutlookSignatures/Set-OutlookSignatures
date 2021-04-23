Param(
    # Path to centrally managed signature templates
    [ValidateScript( {
            Test-Path $_ -PathType Container
        }
    )]
    [string]
    $SignatureTemplatePath = '.\Signature templates',

    [string[]]$DomainsToCheckForGroups = ('*')
)


function Set-Signatures {
    Write-Host "    '$($Signature.Name)'"

    $SignatureFileAlreadyDone = ($global:SignatureFilesDone -contains $($Signature.Name))
    if ($SignatureFileAlreadyDone) {
        Write-Host '      File already processed before'
    } else {
        $global:SignatureFilesDone += $($Signature.Name)
    }

    if ($SignatureFileAlreadyDone -eq $false) {
        $tempPath = (New-Item -ItemType Directory -Path (Join-Path -Path $env:temp -ChildPath ([System.Guid]::NewGuid()))).FullName

        Write-Host '      Copy file and open it in Word'

        $path = $(Join-Path -Path $tempPath -ChildPath $Signature.value).tostring()
        try {
            Copy-Item -LiteralPath $Signature.Name -Destination $path -Force
        } catch {
            Write-Host '        Error copying file. Skipping signature.'
            continue
        }

        $Signature.value = $([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))
        $global:SignatureFilesDone += $Signature.Value#([System.IO.Path]::ChangeExtension($($Signature.value), '.htm'))

        $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdOpenFormat], 'wdOpenFormatAuto')
        $COMWord.Documents.Open($path, $false) | Out-Null

        Write-Host '      Replace variables'
        $ReplaceHash = @{}

        # Currently logged on user
        $replaceHash.Add('$CURRENTUSERGIVENNAME$', [string]$ADPropsCurrentUser.givenname)
        $replaceHash.Add('$CURRENTUSERSURNAME$', [string]$ADPropsCurrentUser.sn)
        $replaceHash.Add('$CURRENTUSERNAMEWITHTITLES$', (((((([string]$ADPropsCurrentUser.svstitelvorne, [string]$ADPropsCurrentUser.givenname, [string]$ADPropsCurrentUser.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUser.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', '))
        $replaceHash.Add('$CURRENTUSERDEPARTMENT$', [string]$ADPropsCurrentUser.department)
        $replaceHash.Add('$CURRENTUSERTITLE$', [string]$ADPropsCurrentUser.title)
        $replaceHash.Add('$CURRENTUSERSTREETADDRESS$', [string]$ADPropsCurrentUser.streetaddress)
        $replaceHash.Add('$CURRENTUSERPOSTALCODE$', [string]$ADPropsCurrentUser.postalcode)
        $replaceHash.Add('$CURRENTUSERLOCATION$', [string]$ADPropsCurrentUser.l)
        $replaceHash.Add('$CURRENTUSERCOUNTRY$', [string]$ADPropsCurrentUser.co)
        $replaceHash.Add('$CURRENTUSERTELEPHONE$', [string]$ADPropsCurrentUser.telephonenumber)
        $replaceHash.Add('$CURRENTUSERFAX$', [string]$ADPropsCurrentUser.facsimiletelephonenumber)
        $replaceHash.Add('$CURRENTUSERMOBILE$', [string]$ADPropsCurrentUser.mobile)
        $replaceHash.Add('$CURRENTUSERMAIL$', [string]$ADPropsCurrentUser.mail)

        # Manager of currently logged on user
        $replaceHash.Add('$CURRENTUSERMANAGERGIVENNAME$', [string]$ADPropsCurrentUserManager.givenname)
        $replaceHash.Add('$CURRENTUSERMANAGERSURNAME$', [string]$ADPropsCurrentUserManager.sn)
        $replaceHash.Add('$CURRENTUSERMANAGERNAMEWITHTITLES$', (((((([string]$ADPropsCurrentUserManager.svstitelvorne, [string]$ADPropsCurrentUserManager.givenname, [string]$ADPropsCurrentUserManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentUserManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', '))
        $replaceHash.Add('$CURRENTUSERMANAGERDEPARTMENT$', [string]$ADPropsCurrentUserManager.department)
        $replaceHash.Add('$CURRENTUSERMANAGERTITLE$', [string]$ADPropsCurrentUserManager.title)
        $replaceHash.Add('$CURRENTUSERMANAGERSTREETADDRESS$', [string]$ADPropsCurrentUserManager.streetaddress)
        $replaceHash.Add('$CURRENTUSERMANAGERPOSTALCODE$', [string]$ADPropsCurrentUserManager.postalcode)
        $replaceHash.Add('$CURRENTUSERMANAGERLOCATION$', [string]$ADPropsCurrentUserManager.l)
        $replaceHash.Add('$CURRENTUSERMANAGERCOUNTRY$', [string]$ADPropsCurrentUserManager.co)
        $replaceHash.Add('$CURRENTUSERMANAGERTELEPHONE$', [string]$ADPropsCurrentUserManager.telephonenumber)
        $replaceHash.Add('$CURRENTUSERMANAGERFAX$', [string]$ADPropsCurrentUserManager.facsimiletelephonenumber)
        $replaceHash.Add('$CURRENTUSERMANAGERMOBILE$', [string]$ADPropsCurrentUserManager.mobile)
        $replaceHash.Add('$CURRENTUSERMANAGERMAIL$', [string]$ADPropsCurrentUserManager.mail)

        # Current mailbox
        $replaceHash.Add('$CURRENTMAILBOXGIVENNAME$', [string]$ADPropsCurrentMailbox.givenname)
        $replaceHash.Add('$CURRENTMAILBOXSURNAME$', [string]$ADPropsCurrentMailbox.sn)
        $replaceHash.Add('$CURRENTMAILBOXNAMEWITHTITLES$', (((((([string]$ADPropsCurrentMailbox.svstitelvorne, [string]$ADPropsCurrentMailbox.givenname, [string]$ADPropsCurrentMailbox.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailbox.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', '))
        $replaceHash.Add('$CURRENTMAILBOXDEPARTMENT$', [string]$ADPropsCurrentMailbox.department)
        $replaceHash.Add('$CURRENTMAILBOXTITLE$', [string]$ADPropsCurrentMailbox.title)
        $replaceHash.Add('$CURRENTMAILBOXSTREETADDRESS$', [string]$ADPropsCurrentMailbox.streetaddress)
        $replaceHash.Add('$CURRENTMAILBOXPOSTALCODE$', [string]$ADPropsCurrentMailbox.postalcode)
        $replaceHash.Add('$CURRENTMAILBOXLOCATION$', [string]$ADPropsCurrentMailbox.l)
        $replaceHash.Add('$CURRENTMAILBOXCOUNTRY$', [string]$ADPropsCurrentMailbox.co)
        $replaceHash.Add('$CURRENTMAILBOXTELEPHONE$', [string]$ADPropsCurrentMailbox.telephonenumber)
        $replaceHash.Add('$CURRENTMAILBOXFAX$', [string]$ADPropsCurrentMailbox.facsimiletelephonenumber)
        $replaceHash.Add('$CURRENTMAILBOXMOBILE$', [string]$ADPropsCurrentMailbox.mobile)
        $replaceHash.Add('$CURRENTMAILBOXMAIL$', [string]$ADPropsCurrentMailbox.mail)

        # Manager of current mailbox
        $replaceHash.Add('$CURRENTMAILBOXMANAGERGIVENNAME$', [string]$ADPropsCurrentMailboxManager.givenname)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERSURNAME$', [string]$ADPropsCurrentMailboxManager.sn)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERNAMEWITHTITLES$', (((((([string]$ADPropsCurrentMailboxManager.svstitelvorne, [string]$ADPropsCurrentMailboxManager.givenname, [string]$ADPropsCurrentMailboxManager.sn) | Where-Object { $_ -ne '' }) -join ' '), [string]$ADPropsCurrentMailboxManager.svstitelhinten) | Where-Object { $_ -ne '' }) -join ', '))
        $replaceHash.Add('$CURRENTMAILBOXMANAGERDEPARTMENT$', [string]$ADPropsCurrentMailboxManager.department)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERTITLE$', [string]$ADPropsCurrentMailboxManager.title)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERSTREETADDRESS$', [string]$ADPropsCurrentMailboxManager.streetaddress)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERPOSTALCODE$', [string]$ADPropsCurrentMailboxManager.postalcode)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERLOCATION$', [string]$ADPropsCurrentMailboxManager.l)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERCOUNTRY$', [string]$ADPropsCurrentMailboxManager.co)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERTELEPHONE$', [string]$ADPropsCurrentMailboxManager.telephonenumber)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERFAX$', [string]$ADPropsCurrentMailboxManager.facsimiletelephonenumber)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERMOBILE$', [string]$ADPropsCurrentMailboxManager.mobile)
        $replaceHash.Add('$CURRENTMAILBOXMANAGERMAIL$', [string]$ADPropsCurrentMailboxManager.mail)

        $wdFindContinue = 1
        $MatchCase = $true
        $MatchWholeWord = $true
        $MatchWildcards = $False
        $MatchSoundsLike = $False
        $MatchAllWordForms = $False
        $Forward = $True
        $Wrap = $wdFindContinue
        $Format = $False
        $wdFindContinue = 1
        $ReplaceAll = 2

        # Replace in current view (show or hide field codes)
        foreach ($replaceKey in $replaceHash.Keys) {
            $FindText = $replaceKey
            $ReplaceWith = $replaceHash.$replaceKey
            $COMWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, `
                    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                    $Wrap, $Format, $ReplaceWith, $ReplaceAll) | Out-Null
        }

        # Invert current view (show or hide field codes)
        # This is neccessary to be able to replace variables in hyperlinks and quicktips of hyperlinks
        $COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = (-not $COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes)
        foreach ($replaceKey in $replaceHash.Keys) {
            $FindText = $replaceKey
            $ReplaceWith = $replaceHash.$replaceKey
            $COMWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord, `
                    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, `
                    $Wrap, $Format, $ReplaceWith, $ReplaceAll) | Out-Null
        }

        # Restore original view
        $COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes = (-not $COMWord.ActiveDocument.ActiveWindow.View.ShowFieldCodes)

        Write-Host '      Save as filtered .HTM file'
        $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatFilteredHTML')
        $path = $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.htm'))
        $COMWord.ActiveDocument.Weboptions.encoding = 65001
        $COMWord.ActiveDocument.SaveAs($path, $saveFormat)

        Write-Host '      Save as .RTF file'
        $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatRTF')
        $path = $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.rtf'))
        $COMWord.ActiveDocument.SaveAs($path, $saveFormat)

        Write-Host '      Save as .TXT file'
        $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], 'wdFormatUnicodeText')
        $path = $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.txt'))
        $COMWord.ActiveDocument.SaveAs($path, $saveFormat)

        $COMWord.ActiveDocument.Close($false)

        Write-Host '      Embed local files in .HTM file and add marker'
        $path = $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.htm'))

        $tempFileContent = Get-Content -LiteralPath $path -Raw -Encoding UTF8

        if ($tempFileContent -notlike "*$HTMLMarkerTag*") {
            if ($tempFileContent -like '*<head>*') {
                $tempFileContent = $tempFileContent -ireplace ('<HEAD>', ('<head>' + $HTMLMarkerTag))
            } else {
                $tempFileContent = $tempFileContent -ireplace ('<HTML>', ('<HTML><head>' + $HTMLMarkerTag + '</head>'))
            }
        }

        $src = @()
        ([regex]'(?i)src="(.*?)"').Matches($tempFileContent) | ForEach-Object {
            $src += $_.Groups[0].Value
            $src += (Join-Path -Path (Split-Path -Path $path -Parent) -ChildPath ([uri]::UnEscapeDataString($_.Groups[1].Value)))
        }

        for ($x = 0; $x -lt $src.count; $x = $x + 2) {
            if ($src[$x].StartsWith('src="data:')) {
            } elseif (Test-Path -LiteralPath $src[$x + 1] -PathType leaf) {
                $fmt = $null
                switch ((Get-ChildItem -LiteralPath $src[$x + 1]).Extension) {
                    '.apng' {
                        $fmt = 'data:image/apng;base64,'
                    }
                    '.avif' {
                        $fmt = 'data:image/avif;base64,'
                    }
                    '.gif' {
                        $fmt = 'data:image/gif;base64,'
                    }
                    '.jpg' {
                        $fmt = 'data:image/jpeg;base64,'
                    }
                    '.jpeg' {
                        $fmt = 'data:image/jpeg;base64,'
                    }
                    '.jfif' {
                        $fmt = 'data:image/jpeg;base64,'
                    }
                    '.pjpeg' {
                        $fmt = 'data:image/jpeg;base64,'
                    }
                    '.pjp' {
                        $fmt = 'data:image/jpeg;base64,'
                    }
                    '.png' {
                        $fmt = 'data:image/png;base64,'
                    }
                    '.svg' {
                        $fmt = 'data:image/svg+xml;base64,'
                    }
                    '.webp' {
                        $fmt = 'data:image/webp;base64,'
                    }
                    '.css' {
                        $fmt = 'data:text/css;base64,'
                    }
                    '.less' {
                        $fmt = 'data:text/css;base64,'
                    }
                    '.js' {
                        $fmt = 'data:text/javascript;base64,'
                    }
                    '.otf' {
                        $fmt = 'data:font/otf;base64,'
                    }
                    '.sfnt' {
                        $fmt = 'data:font/sfnt;base64,'
                    }
                    '.ttf' {
                        $fmt = 'data:font/ttf;base64,'
                    }
                    '.woff' {
                        $fmt = 'data:font/woff;base64,'
                    }
                    '.woff2' {
                        $fmt = 'data:font/woff2;base64,'
                    }
                }
                if ($fmt) {
                    $tempFileContent = $tempFileContent.replace( `
                            $src[$x], `
                        ('src="' + $fmt + [Convert]::ToBase64String([IO.File]::ReadAllBytes($src[$x + 1])) + '"') `
                    )

                } else {
                }
            } else {
            }
        }

        $tempFileContent | Out-File -LiteralPath $path -Encoding UTF8 -Force

        $SignaturePaths | ForEach-Object {
            Write-Host "      Copy signature files to '$_'"
            Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.htm')) -Destination $_ -Force
            Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.rtf')) -Destination $_ -Force
            Copy-Item -LiteralPath $([System.IO.Path]::ChangeExtension((Join-Path -Path $tempPath -ChildPath $Signature.value), '.txt')) -Destination $_ -Force
        }
        Remove-Item -LiteralPath $tempPath -Force -Recurse
    }

    # Set default signature for new mails
    if ($SignatureFilesDefaultNew.contains('' + $Signature.name + '')) {
        for ($j = 0; $j -lt $MailAddresses.count; $j++) {
            if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                Write-Host '      Set signature as default for new messages'
                Set-ItemProperty -Path $RegistryPaths[$j] -Name 'New Signature' -Type String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force
            }
        }
    }

    # Set default signature for replies and forwarded mails
    if ($SignatureFilesDefaultReplyFwd.contains($Signature.name)) {
        for ($j = 0; $j -lt $MailAddresses.count; $j++) {
            if ($MailAddresses[$j] -ieq $MailAddresses[$AccountNumberRunning]) {
                Write-Host '      Set signature as default for reply/forward messages'
                Set-ItemProperty -Path $RegistryPaths[$j] -Name 'Reply-Forward Signature' -Type String -Value (($Signature.value -split '\.' | Select-Object -SkipLast 1) -join '.') -Force
            }
        }
    }
}


Clear-Host

Write-Host 'Script started'

Write-Host '  Check parameters and script environment'
Set-Location $PSScriptRoot | Out-Null
$Search = New-Object DirectoryServices.DirectorySearcher
$Search.PageSize = 1000


if ((Test-Path $SignatureTemplatePath -PathType Container) -eq $false) {
    Write-Host "  Problem connecting to or reading from folder '$SignatureTemplatePath'. Check path."
    exit 1
}

if (($ExecutionContext.SessionState.LanguageMode) -eq 'FullLanguage') {
} else {
    Write-Host "This PowerShell session is in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode."
    Write-Host 'Base64 conversion not possible. Exiting.'
    exit 1
}


Write-Host '  Check Outlook version and profile'
try {
    $COMOutlook = New-Object -ComObject outlook.application
    $OutlookRegistryVersion = [System.Version]::Parse($COMOutlook.Version)
    $OutlookDefaultProfile = $COMOutlook.DefaultProfileName
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($COMOutlook) | Out-Null
    Remove-Variable COMOutlook
} catch {
    Write-Host 'Outlook not installed or not working correctly. Exiting.'
    exit 1
}

if ($OutlookRegistryVersion.major -gt 16) {
    Write-Host "Outlook version $OutlookRegistryVersion is newer than 16 and not yet known. Please inform your administrator. Exiting."
} elseif ($OutlookRegistryVersion.major -eq 16) {
    $OutlookRegistryVersion = '16.0'
} elseif ($OutlookRegistryVersion.major -eq 15) {
    $OutlookRegistryVersion = '15.0'
} elseif ($OutlookRegistryVersion.major -eq 14) {
    $OutlookRegistryVersion = '14.0'
} else {
    Write-Host "Outlook version $OutlookRegistryVersion is below minimum required version 14 (Outlook 2010). Exiting."
    exit 1
}

$HTMLMarkerTag = '<meta name=data-SignatureFileInfo content="Set-OutlookSignatures.ps1">'


Write-Host 'Enumerate domains to check for group memberships'
$x = $DomainsToCheckForGroups
$DomainsToCheckForGroups = @()

# Users own domain/forest is always included
$y = ([ADSI]'LDAP://RootDSE').rootDomainNamingContext -replace ('DC=', '') -replace (',', '.')
if ($y -ne '') {
    Write-Host "  Current user forest: $y"
    $DomainsToCheckForGroups += $y
} else {
    Write-Host '  Problem connecting to Active Directory, or user is a local user. Exiting.'
    exit 1
}

# Other domains - either the list provided, or all outgoing and bidirectional trusts
if (($x.count -eq 1) -and ($x[0] -eq '*')) {
    $Search.SearchRoot = "GC://$(([ADSI]'LDAP://RootDSE').rootDomainNamingContext)"
    $Search.Filter = '(ObjectClass=trustedDomain)'

    $Search.FindAll() | ForEach-Object {
        $TrustNameSID = (New-Object system.security.principal.securityidentifier($($_.properties.securityidentifier), 0)).tostring()
        $TrustOrigin = ($_.properties.distinguishedname -split ',DC=')[1..999] -join '.'
        $TrustName = $_.properties.name
        $TrustDirectionNumber = $_.properties.trustdirection
        $TrustTypeNumber = $_.properties.trusttype
        $TrustAttributesNumber = $_.properties.trustattributes

        #http://msdn.microsoft.com/en-us/library/cc220955.aspx
        Switch ($TrustTypeNumber) {
            1 { $TrustType = 'Downlevel (Windows NT domain external)' }
            2 { $TrustType = 'Uplevel (Active Directory domain - parent-child, root domain, shortcut, external, or forest)' }
            3 { $TrustType = 'MIT (non-Windows) Kerberos version 5 realm' }
            4 { $TrustType = "DCE (Theoretical trust type - DCE refers to Open Group's Distributed Computing Environment specification)" }
            Default { $TrustType = $TrustTypeNumber }
        }

        #http://msdn.microsoft.com/en-us/library/cc223779.aspx
        Switch ($TrustAttributesNumber) {
            1 { $TrustAttributes = 'Non-Transitive' }
            2 { $TrustAttributes = 'Uplevel clients only (Windows 2000 or newer)' }
            4 { $TrustAttributes = 'Quarantined Domain (External)' }
            8 { $TrustAttributes = 'Forest Trust' }
            16 { $TrustAttributes = 'Cross-Organizational Trust (Selective Authentication)' }
            32 { $TrustAttributes = 'Intra-Forest Trust (trust within the forest)' }
            64 { $TrustAttributes = 'Inter-Forest Trust (trust with another forest)' }
            Default { $TrustAttributes = $TrustAttributesNumber }
        }

        #http://msdn.microsoft.com/en-us/library/cc223768.aspx
        Switch ($TrustDirectionNumber) {
            0 { $TrustDirection = 'Disabled (The trust relationship exists but has been disabled)' }
            1 { $TrustDirection = "Incoming (TrustING domain: $Trustname can be authenticated in $TrustOrigin)" }
            2 { $TrustDirection = "Outgoing (TrustED domain: $TrustOrigin can be authenticated in $TrustName)" }
            3 { $TrustDirection = 'Bidirectional (two-way trust)' }
            Default { $TrustDirection = $TrustDirectionNumber }
        }

        # which domains does the current user have access to?
        if (($TrustAttributesNumber -ne 32) -and (($TrustDirectionNumber -eq 2) -or ($TrustDirectionNumber -eq 3)) ) {
            Write-Host "  Trusted domain: $TrustName"
            $DomainsToCheckForGroups += $TrustName
        }
    }
} else {
    $x | ForEach-Object {
        $y = ($_ -replace ('DC=', '') -replace (',', '.'))
        if ($y -eq $_) { 
            Write-Host "  User provided domain/forest: $y" 
        } else {
            Write-Host "  User provided domain/forest: $_ -> $y" 
        }
        if ($y -match '[^a-zA-Z0-9.-]') {
            Write-Host '    Skipping domain. Allowed characters are a-z, A-Z, ., -.' 
        } else {
            $DomainsToCheckForGroups += $y
        }
    }
}


Write-Host 'Get AD properties of currently logged on user and his manager'
try {
    $ADPropsCurrentUser = ([adsisearcher]"(samaccountname=$env:username)").FindOne().Properties
} catch {
    $ADPropsCurrentUser = $null
    Write-Host '  Problem connecting to Active Directory, or user is a local user. Exiting.'
    exit 1
}

try {
    $ADPropsCurrentUserManager = ([adsisearcher]('(distinguishedname=' + $ADPropsCurrentUser.manager + ')')).FindOne().Properties
} catch {
    $ADPropsCurrentUserManager = $null
}

Write-Host 'Get Outlook signature file path(s)'
$SignaturePaths = @()
Get-ItemProperty 'hkcu:\software\microsoft\office\*\common\general' | Where-Object { $_.'Signatures' -ne '' } | ForEach-Object {
    Push-Location (Join-Path -Path $env:AppData -ChildPath 'Microsoft')
    $x = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($_.Signatures)
    if (Test-Path $x -IsValid) {
        if (-not (Test-Path $x -type container)) {
            New-Item -Path $x -ItemType directory -Force
        }
        $SignaturePaths += $x
        Write-Host "  $x"
    }
    Pop-Location
}

Write-Host 'Get mail addresses from Outlook profiles and corresponding registry paths'
$MailAddresses = @()
$RegistryPaths = @()
$LegacyExchangeDNs = @()

if ($OutlookDefaultProfile.length -eq '') {
    Get-ItemProperty "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*" | Where-Object { (($_.'Account Name' -like '*@*.*') -and ($_.'Identity Eid' -ne '')) } | ForEach-Object {
        $MailAddresses += $_.'Account Name'
        $RegistryPaths += $_.PSPath
        $LegacyExchangeDN = ('/O=' + (((($_.'Identity Eid' | ForEach-Object { [char]$_ }) -join '' -replace [char]0) -split '/O=')[-1]).ToString().trim())
        if ($LegacyExchangeDN.length -le 3) {
            $LegacyExchangeDN = ''
        }
        $LegacyExchangeDNs += $LegacyExchangeDN
        Write-Host "  $($_.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $_.PSDrive)"
        Write-Host "    $($_.'Account Name')"
        if ($LegacyExchangeDN -eq '') {
            Write-Host '      No legacyExchangeDN found, assuming mailbox is no Exchange mailbox'
        } else {
            Write-Host '      Found legacyExchangeDN, assuming mailbox is an Exchange mailbox'
            write-host "        $LegacyExchangeDN"
        }
    }
} else {
    # current users mailbox in default profile
    Get-ItemProperty "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\*" | Where-Object { $_.'Account Name' -ieq $ADPropsCurrentUser.mail } | ForEach-Object {
        $MailAddresses += $_.'Account Name'
        $RegistryPaths += $_.PSPath
        $LegacyExchangeDN = ('/O=' + (((($_.'Identity Eid' | ForEach-Object { [char]$_ }) -join '' -replace [char]0) -split '/O=')[-1]).ToString().trim())
        if ($LegacyExchangeDN.length -le 3) {
            $LegacyExchangeDN = ''
        }
        $LegacyExchangeDNs += $LegacyExchangeDN
        Write-Host "  $($_.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $_.PSDrive)"
        Write-Host "    $($_.'Account Name')"
        if ($LegacyExchangeDN -eq '') {
            Write-Host '      No legacyExchangeDN found, assuming mailbox is no Exchange mailbox'
        } else {
            Write-Host '      Found legacyExchangeDN, assuming mailbox is an Exchange mailbox'
            write-host "        $LegacyExchangeDN"
        }
    }

    # other mailboxes in default profile
    Get-ItemProperty "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\$OutlookDefaultProfile\9375CFF0413111d3B88A00104B2A6676\*" | Where-Object { ($_.'Account Name' -like '*@*.*') -and ($_.'Account Name' -ine $ADPropsCurrentUser.mail) } | ForEach-Object {
        $MailAddresses += $_.'Account Name'
        $RegistryPaths += $_.PSPath
        $LegacyExchangeDN = ('/O=' + (((($_.'Identity Eid' | ForEach-Object { [char]$_ }) -join '' -replace [char]0) -split '/O=')[-1]).ToString().trim())
        if ($LegacyExchangeDN.length -le 3) {
            $LegacyExchangeDN = ''
        }
        $LegacyExchangeDNs += $LegacyExchangeDN
        Write-Host "  $($_.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $_.PSDrive)"
        Write-Host "    $($_.'Account Name')"
        if ($LegacyExchangeDN -eq '') {
            Write-Host '      No legacyExchangeDN found, assuming mailbox is no Exchange mailbox'
        } else {
            Write-Host '      Found legacyExchangeDN, assuming mailbox is an Exchange mailbox'
            write-host "        $LegacyExchangeDN"
        }
    }

    # all other mailboxes in all other profiles
    Get-ItemProperty "hkcu:\Software\Microsoft\Office\$OutlookRegistryVersion\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*" | Where-Object { $_.'Account Name' -like '*@*.*' } | ForEach-Object {
        if ($RegistryPaths -notcontains $_.PSPath) {
            $MailAddresses += $_.'Account Name'
            $RegistryPaths += $_.PSPath
            $LegacyExchangeDN = ('/O=' + (((($_.'Identity Eid' | ForEach-Object { [char]$_ }) -join '' -replace [char]0) -split '/O=')[-1]).ToString().trim())
            if ($LegacyExchangeDN.length -le 3) {
                $LegacyExchangeDN = ''
            }
            $LegacyExchangeDNs += $LegacyExchangeDN
            Write-Host "  $($_.PSPath -ireplace [regex]::escape('Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER'), $_.PSDrive)"
            Write-Host "    $($_.'Account Name')"
            if ($LegacyExchangeDN -eq '') {
                Write-Host '      No legacyExchangeDN found, assuming mailbox is no Exchange mailbox'
            } else {
                Write-Host '      Found legacyExchangeDN, assuming mailbox is an Exchange mailbox'
                write-host "        $LegacyExchangeDN"
            }
        }
    }
}


Write-Host 'Get all signature files and categorize them'
$SignatureFilesCommon = @{}
$SignatureFilesGroup = @{}
$SignatureFilesGroupFilePart = @{}
$SignatureFilesMailbox = @{}
$SignatureFilesMailboxFilePart = @{}
$SignatureFilesDefaultNew = @{}
$SignatureFilesDefaultReplyFwd = @{}
$global:SignatureFilesDone = @()

foreach ($SignatureFile in (Get-ChildItem -Path $SignatureTemplatePath -File -Filter '*.docx')) {
    Write-Host ("  '$($SignatureFile.Name)'")
    $x = $SignatureFile.name -split '\.(?![\w\s\d]*\[*(\]|@))'
    if ($x.count -ge 3) {
        $SignatureFilePart = $x[-2]
        $SignatureFileTargetName = ($x[($x.count * -1)..-3] -join '.') + '.' + $x[-1]
    } else {
        $SignatureFilePart = ''
        $SignatureFileTargetName = $SignatureFile.Name
    }

    [regex]::Matches((($SignatureFilePart -replace '(?i)\[DefaultNew\]', '') -replace '(?i)\[DefaultReplyFwd\]', ''), '\[(.*?)\]').captures.value | ForEach-Object {
        if ($_ -eq $null) {
            Write-Host '    Common signature'
            $SignatureFilesCommon.add($SignatureFile.FullName, $SignatureFileTargetName)
        } elseif ($_ -match '(.*?)@(.*?)\.(.*?)') {
            Write-Host '    Mailbox specific signature'
            $SignatureFilesMailbox.add($SignatureFile.FullName, $SignatureFileTargetName)
            $SignatureFilesMailboxFilePart.add($SignatureFile.FullName, $SignatureFilePart)
        } else {
            Write-Host '    Group specific signature'
            $SignatureFilesGroup.add($SignatureFile.FullName, $SignatureFileTargetName)
            $SignatureFilesGroupFilePart.add($SignatureFile.FullName, $SignatureFilePart)
        }
    }

    if ($SignatureFilePart -match '(?i)\[DefaultNew\]') {
        $SignatureFilesDefaultNew.add($SignatureFile.FullName, $SignatureFileTargetName)
        Write-Host '    Default signature for new mails'
    }

    if ($SignatureFilePart -match '(?i)\[DefaultReplyFwd\]') {
        $SignatureFilesDefaultReplyFwd.add($SignatureFile.FullName, $SignatureFileTargetName)
        Write-Host '    Default signature for replies and forwards'
    }
}


# Start Word, as we need it to edit signatures
try {
    $COMWord = New-Object -ComObject word.application
} catch {
    Write-Host 'Word not installed or not working correctly. Exiting.'
    exit 1
}


# Process each mail address only once, but each corresponding registry path
for ($AccountNumberRunning = 0; $AccountNumberRunning -lt $MailAddresses.count; $AccountNumberRunning++) {
    if ($AccountNumberRunning -le $MailAddresses.IndexOf($MailAddresses[$AccountNumberRunning])) {
        Write-Host "Mailbox $($MailAddresses[$AccountNumberRunning])"
        write-host "  $($LegacyExchangeDNs[$AccountNumberRunning])"

        $UserDomain = ''

        Write-Host '  Get AD properties and group membership of mailbox'
        $Groups = @()
        $TokenGroups = @()

        if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
            # Loop through domains until the first one knows the legacyExchangeDN
            for ($DomainNumber = 0; (($DomainNumber -lt $DomainsToCheckForGroups.count) -and ($UserDomain -eq '')); $DomainNumber++) {
                if (($DomainsToCheckForGroups[$DomainNumber] -ne '')) {
                    Write-Host "    $($DomainsToCheckForGroups[$DomainNumber]) (mailbox user object)"
                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($DomainsToCheckForGroups[$DomainNumber])")
                    $Search.filter = '(objectclass=user)'
                    try {
                        $UserAccount = ([ADSI]"$(($Search.FindOne()).path)")
                    } catch {
                        Write-Host "      Error connecting to $($DomainsToCheckForGroups[$DomainNumber]). Removing domain from list."
                        Write-Host '      If this error is permanent, check AD trust and firewall config. Consider using parameter DomainsToCheckForGroups.'
                        $DomainsToCheckForGroups[$DomainNumber] = ''
                        continue
                    }

                    $Search.filter = "(&(objectclass=user)(legacyExchangeDN=$($LegacyExchangeDNs[$AccountNumberRunning])))"
                    $u = $Search.FindOne()
                    if (($u.path -ne '') -and ($null -ne $u.path)) {
                        $UserAccount = [ADSI]"LDAP://$($u.properties.distinguishedname)"
                        $ADPropsCurrentMailbox = $UserAccount.Properties
                        try {
                            $Search.filter = "(distinguishedname=$($ADPropsCurrentMailbox.Manager))"
                            $ADPropsCurrentMailboxManager = ([ADSI]"$(($Search.FindOne()).path)").Properties
                        } catch {
                        }
                        $UserDomain = $DomainsToCheckForGroups[$DomainNumber]
                        $SIDsToCheckInTrusts = @()
                        $SIDsToCheckInTrusts += $UserAccount.objectSid
                        $UserAccount.GetInfoEx(@('tokengroups'), 0)

                        foreach ($sidBytes in $UserAccount.Properties.tokenGroups) {
                            $translated = $null
                            $sid = New-Object System.Security.Principal.SecurityIdentifier($sidbytes, 0)
                            try {
                                $translated = $sid.Translate('System.Security.Principal.NTAccount').ToString()
                            } catch {
                                try {
                                    $adObject = ([ADSI]('LDAP://<SID=' + $sid.ToString() + '>'))
                                    $translated = $adObject.Properties['samAccountName'][0].ToString()
                                } catch {
                                }
                            }
                            $objTrans = New-Object -ComObject 'NameTranslate'
                            $objNT = $objTrans.GetType()
                            $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
                            try {
                                $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $sid.tostring())) # 12 = SIDORSIDHISTORY
                                $Groups += ($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)) -replace '\\', ' ' # 3 = NT4NAME, 4 = DISPLAYNAME
                                Write-Host "      [$($Groups[-1])]" -NoNewline
                                try {
                                    $x = (($Groups[-1] -split ' ')[0] + ' ' + ($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 4))) # 3 = NT4NAME, 4 = DISPLAYNAME
                                    if ($x -ine $Groups[-1]) {
                                        $Groups += $x
                                        Write-Host ", [$x]"
                                    } else {
                                        Write-Host
                                    }
                                } catch {
                                    Write-Host
                                }
                            } catch {
                                continue
                            }
                        }
                        $UserAccount.GetInfoEx(@('tokengroupsglobalanduniversal'), 0)
                        $SIDsToCheckInTrusts += $UserAccount.properties.tokengroupsglobalanduniversal
                    }
                }
            }

            # Loop through all domains to check if the mailbox account has a group membership there
            # Across a trust, a user can only be added to a domain local group.
            # Domain local groups can not be used outside their own domain, so we don't need to query recursively
            if ($SIDsToCheckInTrusts.count -gt 0) {
                $LdapFilterSIDs = '(|'
                $SIDsToCheckInTrusts | ForEach-Object {
                    try {
                        $SidHex = @()
                        $ot = New-Object System.Security.Principal.SecurityIdentifier($_, 0)
                        $c = New-Object 'byte[]' $ot.BinaryLength
                        $ot.GetBinaryForm($c, 0)
                        $c | ForEach-Object {
                            $SidHex += $('\{0:x2}' -f $_)
                        }
                        $LdapFilterSIDs += ('(objectsid=' + $($SidHex -join '') + ')')
                    } catch {
                    }
                }
                $LdapFilterSIDs += ')'
            } else {
                $LdapFilterSIDs = ''
            }

            for ($DomainNumber = 0; $DomainNumber -lt $DomainsToCheckForGroups.count; $DomainNumber++) {
                if (($DomainsToCheckForGroups[$DomainNumber] -ne '') -and ($DomainsToCheckForGroups[$DomainNumber] -ine $UserDomain)) {
                    Write-Host "    $($DomainsToCheckForGroups[$DomainNumber]) (mailbox user object membership across trusts)"
                    $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$($DomainsToCheckForGroups[$DomainNumber])")
                    $Search.filter = '(objectclass=user)'
                    try {
                        $UserAccount = ([ADSI]"$(($Search.FindOne()).path)")
                    } catch {
                        Write-Host "      Error connecting to $($DomainsToCheckForGroups[$DomainNumber]). Removing domain from list."
                        Write-Host '      If this error is permanent, check AD trust and firewall config. Consider using parameter DomainsToCheckForGroups.'
                        $DomainsToCheckForGroups[$DomainNumber] = ''
                        continue
                    }
                    if ($LdapFilterSIDs -eq '') {
                        continue
                    }

                    $Search.filter = "(&(objectclass=foreignsecurityprincipal)$LdapFilterSIDs)"
                    foreach ($fsp in $Search.FindAll()) {
                        if (($fsp.path -ne '') -and ($null -ne $fsp.path)) {
                            # Foreign Security Principals do not have the tokenGroups attribute
                            # We need to switch to another, slower search method
                            #
                            # member:1.2.840.113556.1.4.1941:= (LDAP_MATCHING_RULE_IN_CHAIN) only returns domain local groups from the domain defined in searchroot
                            # A Foreign Security Principal ist created in each (sub)domain, in which it is granted permissions,
                            # and it can only be member of a domain local group - so we set the searchroot to the (sub)domain of the Foreign Security Principal.
                            $Search.searchroot = New-Object System.DirectoryServices.DirectoryEntry("GC://$((($fsp.path -split ',DC=')[1..999] -join '.'))")                            
                            $Search.filter = "(member:1.2.840.113556.1.4.1941:=$($fsp.Properties.distinguishedname))"

                            foreach ($group in $Search.findall()) {
                                $translated = $null
                                $sid = New-Object System.Security.Principal.SecurityIdentifier($group.properties.objectsid[0], 0)
                                try {
                                    $translated = $sid.Translate('System.Security.Principal.NTAccount').ToString()
                                } catch {
                                    try {
                                        $adObject = ([ADSI]("LDAP://$((($fsp.path -split ',DC=')[1..999] -join '.'))/<SID=" + $sid.ToString() + '>'))
                                        $translated = $adObject.Properties['samAccountName'][0].ToString()
                                    } catch {
                                    }
                                }
                                $objTrans = New-Object -ComObject 'NameTranslate'
                                $objNT = $objTrans.GetType()
                                $objNT.InvokeMember('Init', 'InvokeMethod', $Null, $objTrans, (3, $Null)) # 3 = ADS_NAME_INITTYPE_GC
                                try {
                                    $objNT.InvokeMember('Set', 'InvokeMethod', $Null, $objTrans, (12, $sid.tostring())) # 12 = SIDORSIDHISTORY
                                    $Groups += ($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 3)) -replace '\\', ' ' # 3 = NT4NAME, 4 = DISPLAYNAME
                                    Write-Host "      [$($Groups[-1])]" -NoNewline
                                    try {
                                        $x = (($Groups[-1] -split ' ')[0] + ' ' + ($objNT.InvokeMember('Get', 'InvokeMethod', $Null, $objTrans, 4))) # 3 = NT4NAME, 4 = DISPLAYNAME
                                        if ($x -ine $Groups[-1]) {
                                            $Groups += $x
                                            Write-Host ", [$x]"
                                        } else {
                                            Write-Host
                                        }
                                    } catch {
                                        Write-Host
                                    }
                                } catch {
                                    continue
                                }
                            }
                        }
                    }
                }
            }
        } else {
            Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox'
        }


        Write-Host '  Get SMTP addresses'
        $CurrentMailboxSMTPAddresses = @()
        if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
            $ADPropsCurrentMailbox.proxyaddresses | ForEach-Object {
                if ([string]$_ -ilike 'smtp:*') {
                    $CurrentMailboxSMTPAddresses += [string]$_ -ireplace 'smtp:', ''
                    Write-Host ('    [' + ([string]$_ -ireplace 'smtp:', '') + ']')
                }
            }
        } else {
            $CurrentMailboxSMTPAddresses += $($MailAddresses[$AccountNumberRunning])
            Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox'
            Write-Host '    Using mailbox name as single known SMTP address'
        }

        Write-Host '  Process common signatures'
        foreach ($Signature in $SignatureFilesCommon.GetEnumerator()) {
            Set-Signatures
        }

        Write-Host '  Process group signatures'
        $OutlookWebHash = @{}
        if (($($LegacyExchangeDNs[$AccountNumberRunning]) -ne '')) {
            foreach ($x in $SignatureFilesGroupFilePart.GetEnumerator()) {
                $Groups | ForEach-Object {
                    if ($x.Value.tolower().Contains('[' + $_.tolower() + ']')) {
                        $OutlookWebHash.add($x.Name, $SignatureFilesGroup[$x.Name])
                    }
                }
            }
            foreach ($Signature in $OutlookWebHash.GetEnumerator()) {
                Set-Signatures
            }
        } else {
            $CurrentMailboxSMTPAddresses += $($MailAddresses[$AccountNumberRunning])
            Write-Host '    Skipping, as mailbox has no legacyExchangeDN and is assumed not to be an Exchange mailbox'
        }

        Write-Host '  Process mail address specific signatures'
        $SignatureHash = @{}
        foreach ($x in $SignatureFilesMailboxFilePart.GetEnumerator()) {
            foreach ($y in $CurrentMailboxSMTPAddresses) {
                if ($x.Value.tolower().contains('[' + $y.tolower() + ']')) {
                    $SignatureHash.add($x.Name, $SignatureFilesMailbox[$x.Name])
                }
            }
        }
        foreach ($Signature in $SignatureHash.GetEnumerator()) {
            Set-Signatures
        }
    }

    # Outlook Web Access
    if ($ADPropsCurrentMailbox.mail -ieq $ADPropsCurrentUser.mail) {
        Write-Host '  Setting Outlook Web signature'
        # if the mailbox of the currenlty logged on user is part of his default Outlook Profile, copy the signature to OWA
        for ($j = 0; $j -lt $MailAddresses.count; $j++) {
            if ($MailAddresses[$j] -ieq [string]$ADPropsCurrentUser.mail) {
                if ($RegistryPaths[$j] -like ('*\Outlook\Profiles\' + $OutlookDefaultProfile + '\9375CFF0413111d3B88A00104B2A6676\*')) {
                    try {
                        $TempNewSig = Get-ItemPropertyValue -LiteralPath $RegistryPaths[$j] -Name 'New Signature'
                    } catch {
                        $TempNewSig = ''
                    }
                    try {
                        $TempReplySig = Get-ItemPropertyValue -LiteralPath $RegistryPaths[$j] -Name 'Reply-Forward Signature'
                    } catch {
                        $TempReplySig = ''
                    }
                    if (($TempNewSig -eq '') -and ($TempReplySig -eq '')) {
                        Write-Host '    No default signatures defined, nothing to do'
                        $TempOWASigFile = $null
                        $TempOWASigSetNew = $null
                        $TempOWASigSetReply = $null
                    }

                    if (($TempNewSig -ne '') -and ($TempReplySig -eq '')) {
                        Write-Host '    Signature for new mails found'
                        $TempOWASigFile = $TempNewSig
                        $TempOWASigSetNew = 'True'
                        $TempOWASigSetReply = 'False'
                    }

                    if (($TempNewSig -eq '') -and ($TempReplySig -ne '')) {
                        Write-Host '    Default signature for reply/forward found'
                        $TempOWASigFile = $TempReplySig
                        $TempOWASigSetNew = 'False'
                        $TempOWASigSetReply = 'True'
                    }


                    if ((($TempNewSig -ne '') -and ($TempReplySig -ne '')) -and ($TempNewSig -ine $TempReplySig)) {
                        Write-Host '    Different default signatures for new and reply/forward found, using new signature'
                        $TempOWASigFile = $TempNewSig
                        $TempOWASigSetNew = 'True'
                        $TempOWASigSetReply = 'False'
                    }

                    if ((($TempNewSig -ne '') -and ($TempReplySig -ne '')) -and ($TempNewSig -ieq $TempReplySig)) {
                        Write-Host '    Same default signature for new and reply/forward'
                        $TempOWASigFile = $TempNewSig
                        $TempOWASigSetNew = 'True'
                        $TempOWASigSetReply = 'True'
                    }
                    if (($null -ne $TempOWASigFile) -and ($TempOWASigFile -ne '')) {
                        try {
                            Import-Module -Name '.\Microsoft.Exchange.WebServices.dll'
                            $exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
                            $exchService.UseDefaultCredentials = $true
                            $exchService.AutodiscoverUrl($ADPropsCurrentUser.mail)
                            $folderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $($ADPropsCurrentUser.mail))
                            #Specify the Root folder where the FAI Item is
                            $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($exchService, 'OWA.UserOptions', $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
                            $hsHtmlSignature = (Get-Content -LiteralPath (Join-Path -Path $SignaturePaths[0] -ChildPath ($TempOWASigFile + '.htm')) -Raw).ToString()
                            $stTextSig = (Get-Content -LiteralPath (Join-Path -Path $SignaturePaths[0] -ChildPath ($TempOWASigFile + '.txt')) -Raw).ToString()

                            $OutlookWebHash = @{}
                            # Keys are case sensitive when setting them
                            $OutlookWebHash.Add('signaturehtml', $hsHtmlSignature)
                            $OutlookWebHash.Add('signaturetext', $stTextSig)
                            $OutlookWebHash.Add('signaturetextonmobile', $stTextSig)
                            $OutlookWebHash.Add('autoaddsignature', $TempOWASigSetNew)
                            $OutlookWebHash.Add('autoaddsignatureonmobile', $TempOWASigSetNew)
                            $OutlookWebHash.Add('autoaddsignatureonreply', $TempOWASigSetReply)

                            foreach ($OutlookWebHashKey in $OutlookWebHash.Keys) {
                                if ($UsrConfig.Dictionary.ContainsKey($OutlookWebHashKey)) {
                                    $UsrConfig.Dictionary[$OutlookWebHashKey] = $OutlookWebHash.$OutlookWebHashKey
                                } else {
                                    $UsrConfig.Dictionary.Add($OutlookWebHashKey, $OutlookWebHash.$OutlookWebHashKey)
                                }
                            }

                            $UsrConfig.Update()
                        } catch {
                            Write-Host '    Error setting Outlook Web signature, please contact your administrator'
                        }

                        Remove-Module -Name '.\Microsoft.Exchange.WebServices.dll' -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
}


# Quit word, as all signatures have been edited
$COMWord.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($COMWord) | Out-Null
Remove-Variable COMWord


# Delete old signatures created by this script, which are no longer available in $SignatureTemplatePath
# We check all local signatures for a specific marker in HTML code, so we don't touch user created signatures
Write-Host 'Removing old signatures created by this script, which are no longer centrally available'
$SignaturePaths | ForEach-Object {
    Get-ChildItem $_ -Filter '*.htm' -File | ForEach-Object {
        if ((Get-Content -LiteralPath $_.fullname -Raw) -like ('*' + $HTMLMarkerTag + '*')) {
            if (($_.name -notin $global:SignatureFilesDone) -and ($_.name -notin $SignatureFilesCommon.values) -and ($_.name -notin $SignatureFilesMailbox.Values) -and ($_.name -notin $SignatureFilesGroup.Values)) {
                Write-Host ("  '" + $([System.IO.Path]::ChangeExtension($_.fullname, '')) + "*'")
                Remove-Item -LiteralPath $_.fullname -Force -ErrorAction silentlycontinue
                Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($_.fullname, '.rtf'))) -Force -ErrorAction silentlycontinue
                Remove-Item -LiteralPath ($([System.IO.Path]::ChangeExtension($_.fullname, '.txt'))) -Force -ErrorAction silentlycontinue
            }
        }
    }
}