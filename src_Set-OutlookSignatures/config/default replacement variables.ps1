<#
    This file allows defining custom replacement variables for Set-OutlookSignatures

    This script is executed as a whole once for each mailbox.
    It allows for complex replacement variable handling (complex string transformations, retrieving information from web services and databases, etc.).
    Important when the final text value of a variable contains another variable: Variables are not replaced in the order they are defined in this file,
      but alphabetically using the sort order culture 127 (invariant).

    Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.

    Replacement variable names are not case sensitive.

    A variable defined in this file overrides the definition of the same variable defined earlier in the software.


    See the help and support center (https://set-outlooksignatures.com/help) for more examples, such as:
    Allowed tags (https://set-outlooksignatures.com/details#allowed-tags)
    How to work with INI files (https://set-outlooksignatures.com/details#how-to-work-with-ini-files)
    Replacement variables (https://set-outlooksignatures.com/details#replacement-variables)
    Photos from Active Directory (https://set-outlooksignatures.com/details#photos-account-pictures-user-image-from-active-directory-or-entra-id)
    Delete images when attribute is empty, variable content based on group membership (https://set-outlooksignatures.com/faq#delete-images-when-attribute-is-empty-variable-content-based-on-group-membership)
    How to avoid blank lines when replacement variables return an empty string (https://set-outlooksignatures.com/faq#how-to-avoid-blank-lines-when-replacement-variables-return-an-empty-string)
#>


<#
    What is the recommended approach for custom configuration files?
    You should not change the default configuration file '.\config\default replacement variable.ps1', as it might be changed in a future release of Set-OutlookSignatures. In this case, you would have to sort out the changes yourself.

    The following steps are recommended:
        1. Create a new custom configuration file in a separate folder.
        2. The first step in the new custom configuration file should be to load the default configuration file:
            # Loading default replacement variables shipped with Set-OutlookSignatures
            . ([System.Management.Automation.ScriptBlock]::Create((ConvertEncoding -InFile $(Join-Path -Path $(Get-Location).ProviderPath -ChildPath '/config/default replacement variables.ps1') -InIsHtml $false)))
        3. After importing the default configuration file, existing replacement variables can be altered with custom definitions and new replacement variables can be added.
        4. Instead of altering existing replacement variables, it is recommended to create new replacement variables with modified content.
        5. Start Set-OutlookSignatures with the parameter 'ReplacementVariableConfigFile' pointing to the new custom configuration file.


    To simplify signature design in limited space, shorter versions of replacement variables are made available automatically:
        - CurrentUser -> U
          Example: '$CurrentUserVariableX$' is also available as '$UVariableX$'
        - CurrentUserManager -> UM
          Example: '$CurrentUserManagerVariableX$' is also available as '$UMVariableX$'
        - CurrentMailbox -> M
          Example: '$CurrentMailboxVariableX$' is also available as '$MVariableX$'
        - CurrentMailboxManager -> MM
          Example: '$CurrentMailboxManagerVariableX$' is also available as '$MMVariableX$'
#>


# Basic replacement variables
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    foreach ($ReplacementVariableAttributePair in @(
            , @('GivenName', 'givenName')
            , @('Surname', 'sn')
            , @('Department', 'department')
            , @('Title', 'title')
            , @('StreetAddress', 'streetAddress')
            , @('PostalCode', 'postalCode')
            , @('Location', 'l')
            , @('Country', 'co')
            , @('State', 'st')
            , @('Telephone', 'telephoneNumber')
            , @('Fax', 'facsimileTelephoneNumber')
            , @('Mobile', 'mobile')
            , @('Mail', 'mail')
            , @('Photo', 'thumbnailPhoto')
            , @('PhotoDeleteEmpty', 'thumbnailPhoto')
            , @('ExtAttr1', 'extensionAttribute1')
            , @('ExtAttr2', 'extensionAttribute2')
            , @('ExtAttr3', 'extensionAttribute3')
            , @('ExtAttr4', 'extensionAttribute4')
            , @('ExtAttr5', 'extensionAttribute5')
            , @('ExtAttr6', 'extensionAttribute6')
            , @('ExtAttr7', 'extensionAttribute7')
            , @('ExtAttr8', 'extensionAttribute8')
            , @('ExtAttr9', 'extensionAttribute9')
            , @('ExtAttr10', 'extensionAttribute10')
            , @('ExtAttr11', 'extensionAttribute11')
            , @('ExtAttr12', 'extensionAttribute12')
            , @('ExtAttr13', 'extensionAttribute13')
            , @('ExtAttr14', 'extensionAttribute14')
            , @('ExtAttr15', 'extensionAttribute15')
            , @('Office', 'physicalDeliveryOfficeName')
            , @('Company', 'company')
            , @('MailNickname', 'mailNickname')
            , @('DisplayName', 'displayName')
        )
    ) {
        $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttributePair[0])`$"] = $(
            if ($ReplacementVariableAttributePair[0] -iin @('Photo', 'PhotoDeleteEmpty')) {
                (Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).($ReplacementVariableAttributePair[1])
            } else {
                [string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).($ReplacementVariableAttributePair[1])
            }
        )
    }
}


<#
    Sample code: Full user name including honorific and academic titles
    $CurrentUserNameWithHonorifics$, $CurrentUserManagerNameWithHonorifics$, $CurrentMailboxNameWithHonorifics$, $CurrentMailboxManagerNameWithHonorifics$

    According to standards in German speaking countries:
      "<custom AD attribute 'honorificPrefix'> <standard AD attribute 'givenname'> <standard AD attribute 'surname'>, <custom AD attribute 'honorificSuffix'>"
        If one or more attributes are not set, unnecessary whitespaces and commas are avoided

    Examples:
      Mag. Dr. John Doe, BA MA PhD
      Dr. John Doe
      John Doe, PhD
      John Doe

    Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.
#>
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    $ReplaceHash["`$$($ReplacementVariableNamespace)NameWithHonorifics`$"] = @(
        @(
            (
                @(
                    @(
                        [string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).honorificPrefix
                        [string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).givenname
                        [string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).sn
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                ) -join ' '
            )
            [string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).honorificSuffix
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    ) -join ', '
}


<#
    Sample code: Take salutation or gender pronouns string from Extension Attribute 3
    $CurrentUserSalutation$, $CurrentUserManagerSalutation$, $CurrentMailboxSalutation$, $CurrentMailboxManagerSalutation$
    $CurrentUserGenderPronouns$, $CurrentUserManagerGenderPronouns$, $CurrentMailboxGenderPronouns$, $CurrentMailboxManagerGenderPronouns$

    Format
      If ExtensionAttribute3 is not empty or whitespace, put it in brackets and add a leading space
        Examples: " (Mr.)", " (Ms.)", " (she/her)"
      Else: '' (emtpy string)

    Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.
#>
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    $ReplaceHash["`$$($ReplacementVariableNamespace)Salutation`$"] = $ReplaceHash["`$$($ReplacementVariableNamespace)GenderPronouns`$"] = $(
        if ([string]::IsNullOrWhiteSpace([string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).extensionattribute3)) {
            $null
        } else {
            " ($([string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).extensionattribute3))"
        }
    )
}


<#
    Sample code: Avoid blank lines in signature if replacement variables are empty

    Details: https://set-outlooksignatures.com/faq#how-to-avoid-blank-lines-when-replacement-variables-return-an-empty-string
#>
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    foreach ($ReplacementVariableAttribute in @('Telephone', 'Mobile', 'Fax')) {
        $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)-prefix-noempty`$"] = $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)-noempty`$"] = $(
            if (-not $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)`$"]) {
                ''
            } else {
                $(if ($UseHtmTemplates) { '<br>' } else { "`n" }) + "$($ReplacementVariableAttribute): "
            }
        )
    }
}


<#
    Create $Current[User|Manager|Mailbox|MailboxManager][Telephone|Fax|Mobile]-[E164|INTERNATIONAL|NATIONAL|RFC3966]$ replacement variables
    FormatPhoneNumber: Format phone number in different formats

    Examples
    FormatPhoneNumber -Number $ReplaceHash['$CurrentUserTelephone$'] -Country $ReplaceHash['$CurrentUserCountry$'] -Format 'INTERNATIONAL'
    FormatPhoneNumber -Number $ReplaceHash['$CurrentUserTelephone$'] -Country $ReplaceHash['$CurrentUserCountry$'] -Format 'RFC3966'

    Parameters
    Number
        The phone number to format or parse, as a string. Can include country code or be in local format.
        Extensions can only be detected reliably when marked with common indicators such as "ext", "ext.", "x", "x.", ";ext=", ",", or ";".
        There is comprehensive public information about country codes and national destination codes, but not on how
        carriers actually handle numbers they assign. Service numbers, short numbers, portable numbers make automatic extension detection practically impossible.
    Country
        Either a two-letter ISO country code (e.g., "AT", "US") or full English country name (e.g., "Austria", "United States").
        Required when the phone number does not include a country code such as +43 or +1.
        Country codes starting with 00 ('+0043 ...') can only be interpreted correctly if the Country parameter is specified.
    Format
        Desired phone number format.
        Examples are based on two numbers:
        '+1 305 418 9136,56', which is '305 418 9136 ext 56' with country set to 'US'.
        '+43 50 123456,7890', which is '050 123456 ext 7890' with country set to 'AT'.
        Format is one of the following:
        E164
            International format used for carrier routing. Not intended to be displayed to end users.
            Examples (note the missing extension):
            +13054189136
            +4350123456
        INTERNATIONAL
            Displaying numbers to users in a global context (e.g., contact lists, websites).
            Examples:
            +1 305-418-9136 ext. 56
            +43 50 123 456 ext. 7890
        NATIONAL
            Local format as dialed within the country, no country code.
            Examples:
            (305) 418-9136 ext. 56
            050 123 456 ext. 7890
        RFC3966
            Embedding phone numbers in hyperlinks (tel:+43-1-23456789) or machine-readable formats.
            Examples:
            tel:+1-305-418-9136;ext=56
            tel:+43-50-123-456;ext=7890
        CUSTOM
            Useful when you need to extract parts of the phone number for custom formatting.
            Returns an object with the following properties:
            CountryCode (int), NationalDestinationCode (string), SubscriberNumber (string), Extension (string),
            ParseResult (a PhoneNumber object), OriginalInput (string), ErrorMessage (string)
            Examples:
            CountryCode 1, NationDesitionCode 305, SubscriberNumber 4189136, Extension 56
            CountryCode 43, NationalDestinationCode 50, SubscriberNumber 123456, Extension 7890
#>
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    foreach ($ReplacementVariableAttribute in @('Telephone', 'Fax', 'Mobile')) {
        foreach ($ReplacementVariableAttributeFormat in @('E164', 'INTERNATIONAL', 'NATIONAL', 'RFC3966')) {
            if ($ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)`$"]) {
                $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)-$($ReplacementVariableAttributeFormat)`$"] = FormatPhoneNumber -Number $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)`$"] -Country $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"] -Format $ReplacementVariableAttributeFormat
            } else {
                $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)-$($ReplacementVariableAttributeFormat)`$"] = $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)`$"]
            }
        }
    }
}


<#
    Example: Custom formatting in a (technically wrong) style often seen in German speaking countries
    '+1 305 418 9136,56' -> '+1 (0) 305 4189136 DW 56'
    '+43 50 123456,7890' -> '+43 (0) 50 123456 DW 7890'
#>
<#
    foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
        foreach ($ReplacementVariableAttribute in @('Telephone', 'Fax', 'Mobile')) {
            , (FormatPhoneNumber -Number $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)`$"] -Country $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"] -Format CUSTOM) | ForEach-Object {
                $ReplaceHash["`$$($ReplacementVariableNamespace)$($ReplacementVariableAttribute)-CustomGermanFormat`$"] = $(
                    if ($_.ErrorMessage) {
                        $_.OriginalInput
                    } else {
                        @(
                            @(
                                "+$($_.CountryCode)"
                                '(0)'
                                "$($_.NationalDestinationCode)"
                                "$($_.SubscriberNumber)"
                                "$(if ($_.Extension) { "DW $($_.Extension)" } else { '' } )"
                            ) | Where-Object { $_ }
                        ) -join ' '
                    }
                )
            }
        }
    }
#>


<#
    Sample code: Create vCard QR codes and save the images in the following replacement variables:
    $CurrentUserCustomImage1$, $CurrentUserManagerCustomImage1$, $CurrentMailboxCustomImage1$, $CurrentMailboxManagerCustomImage1$

    You are not limited to vCard, you can create any QR code content you like.

    Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.
#>
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    $QRCodeContent = @(
        @(
            @(
                'BEGIN:VCARD'
                'VERSION:2.1'
                "N:$($ReplaceHash["`$$($ReplacementVariableNamespace)Surname`$"]);$($ReplaceHash["`$$($ReplacementVariableNamespace)GivenName`$"]);;$([string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).honorificPrefix);$([string](Get-Variable -Name "ADProps$($ReplacementVariableNamespace)" -ValueOnly).honorificSuffix)"
                "TITLE:$($ReplaceHash["`$$($ReplacementVariableNamespace)Title`$"])"
                "ORG:$($ReplaceHash["`$$($ReplacementVariableNamespace)Company`$"])"
                "EMAIL;WORK;INTERNET:$($ReplaceHash["`$$($ReplacementVariableNamespace)Mail`$"])"
                "TEL;WORK;VOICE:$($ReplaceHash["`$$($ReplacementVariableNamespace)Telephone-RFC3966`$"] -ireplace 'tel:', '' -ireplace ';ext=', ',')"
                "TEL;WORK;CELL:$($ReplaceHash["`$$($ReplacementVariableNamespace)Mobile-RFC3966$"] -ireplace 'tel:', '' -ireplace ';ext=', ',')"
                "ADR;WORK:;;$($ReplaceHash["`$$($ReplacementVariableNamespace)StreetAddress`$"]);$($ReplaceHash["`$$($ReplacementVariableNamespace)Location`$"]);$($ReplaceHash["`$$($ReplacementVariableNamespace)State`$"]);$($ReplaceHash["`$$($ReplacementVariableNamespace)Postalcode`$"]);$($ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"])"
                'END:VCARD'
            ) | ForEach-Object { $_.trim() }
        ) | Where-Object { (-not [string]::IsNullOrWhiteSpace($_)) -and (-not $_.EndsWith(':')) }
    ) -join ("`r`n")

    if ($QRCodeContent -notmatch '\r\nN:.*\r\n') { $QRCodeContent = 'https://set-outlooksignatures.com' }

    $ReplaceHash["`$$($ReplacementVariableNamespace)CustomImage1`$"] = ((New-Object -TypeName QRCoder.PngByteQRCode -ArgumentList ((New-Object -TypeName QRCoder.QRCodeGenerator).CreateQrCode($QRCodeContent, 'L', $true))).GetGraphic(20, [byte[]]@(0, 0, 0), [byte[]]@(255, 255, 255), $false))
}


<#
    Format an address according to country specific rules
#>
# Create $Current[User|Manager|Mailbox|MailboxManager]PostalAddress$ replacement variables
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    $FormatPostAddressOptions = @{
        # Address components as described in https://github.com/OpenCageData/address-formatting/blob/master/conf/components.yaml
        Components      = @{
            attention = @(
                @(
                    @(
                        "$($ReplaceHash["`$$($ReplacementVariableNamespace)GivenName`$"]) $($ReplaceHash["`$$($ReplacementVariableNamespace)Surname`$"])"
                        "$($ReplaceHash["`$$($ReplacementVariableNamespace)Department`$"])"
                        "$($ReplaceHash["`$$($ReplacementVariableNamespace)Company`$"])"
                    ) | ForEach-Object { $_.trim() }
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            ) -join [System.Environment]::NewLine
            road      = $ReplaceHash["`$$($ReplacementVariableNamespace)StreetAddress`$"]
            city      = $ReplaceHash["`$$($ReplacementVariableNamespace)Location`$"]
            postcode  = $ReplaceHash["`$$($ReplacementVariableNamespace)Postalcode`$"]
            state     = $ReplaceHash["`$$($ReplacementVariableNamespace)State`$"]
            country   = $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"]
        }

        # Country as two-letter ISO country code (e.g., "AT", "US") or full English country name (e.g., "Austria", "United States")
        #   Needed to choose correct address format rules
        Country         = Resolve-Country -InputString $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"] -ReturnType 'cca2' -FallbackValue 'AT'

        # Shorten address components ("St." instead of "Street", "Rd." instead of "Road", etc.)
        Abbreviate      = $false

        # Only return known parts of the address, omit unknown parts
        #   When disabled, unknown parts are added the the "attention" component
        OnlyAddress     = $false

        # Use a custom address template instead of the predefined ones
        #   Predefined templates: https://github.com/OpenCageData/address-formatting/blob/master/conf/countries/worldwide.yaml
        AddressTemplate = $null
    }

    if ($UseHtmTemplates) {
        $ReplaceHash["`$$($ReplacementVariableNamespace)PostalAddress`$"] = [System.Net.WebUtility]::HtmlEncode((Format-PostalAddress @FormatPostAddressOptions)) -replace '\r?\n', '<br />' # Converts paragraphs to line breaks
    } else {
        $ReplaceHash["`$$($ReplacementVariableNamespace)PostalAddress`$"] = (Format-PostalAddress @FormatPostAddressOptions) -replace '\r?\n', "`n" # Converts paragraphs to line breaks
    }
}

# Company name and address only
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    $FormatPostAddressOptions = @{
        # Address components as described in https://github.com/OpenCageData/address-formatting/blob/master/conf/components.yaml
        Components      = @{
            attention = $ReplaceHash["`$$($ReplacementVariableNamespace)Company`$"]
            road      = $ReplaceHash["`$$($ReplacementVariableNamespace)StreetAddress`$"]
            city      = $ReplaceHash["`$$($ReplacementVariableNamespace)Location`$"]
            postcode  = $ReplaceHash["`$$($ReplacementVariableNamespace)Postalcode`$"]
            state     = $ReplaceHash["`$$($ReplacementVariableNamespace)State`$"]
            country   = $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"]
        }

        # Country as two-letter ISO country code (e.g., "AT", "US") or full English country name (e.g., "Austria", "United States")
        #   Needed to choose correct address format rules
        Country         = Resolve-Country -InputString $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"] -ReturnType 'cca2' -FallbackValue 'AT'

        # Shorten address components ("St." instead of "Street", "Rd." instead of "Road", etc.)
        Abbreviate      = $false

        # Only return known parts of the address, omit unknown parts
        #   When disabled, unknown parts are added the the "attention" component
        OnlyAddress     = $true

        # Use a custom address template instead of the predefined ones
        #   Predefined templates: https://github.com/OpenCageData/address-formatting/blob/master/conf/countries/worldwide.yaml
        AddressTemplate = $null
    }

    if ($UseHtmTemplates) {
        $ReplaceHash["`$$($ReplacementVariableNamespace)PostalAddressCompany`$"] = [System.Net.WebUtility]::HtmlEncode((Format-PostalAddress @FormatPostAddressOptions)) -replace '\r?\n', '<br />' # Converts paragraphs to line breaks
    } else {
        $ReplaceHash["`$$($ReplacementVariableNamespace)PostalAddressCompany`$"] = (Format-PostalAddress @FormatPostAddressOptions) -replace '\r?\n', "`n" # Converts paragraphs to line breaks
    }
}

# Address only
foreach ($ReplacementVariableNamespace in @('CurrentUser', 'CurrentUserManager', 'CurrentMailbox', 'CurrentMailboxManager')) {
    $FormatPostAddressOptions = @{
        # Address components as described in https://github.com/OpenCageData/address-formatting/blob/master/conf/components.yaml
        Components      = @{
            attention = ''
            road      = $ReplaceHash["`$$($ReplacementVariableNamespace)StreetAddress`$"]
            city      = $ReplaceHash["`$$($ReplacementVariableNamespace)Location`$"]
            postcode  = $ReplaceHash["`$$($ReplacementVariableNamespace)Postalcode`$"]
            state     = $ReplaceHash["`$$($ReplacementVariableNamespace)State`$"]
            country   = $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"]
        }

        # Country as two-letter ISO country code (e.g., "AT", "US") or full English country name (e.g., "Austria", "United States")
        #   Needed to choose correct address format rules
        Country         = Resolve-Country -InputString $ReplaceHash["`$$($ReplacementVariableNamespace)Country`$"] -ReturnType 'cca2' -FallbackValue 'AT'

        # Shorten address components ("St." instead of "Street", "Rd." instead of "Road", etc.)
        Abbreviate      = $false

        # Only return known parts of the address, omit unknown parts
        #   When disabled, unknown parts are added the the "attention" component
        OnlyAddress     = $true

        # Use a custom address template instead of the predefined ones
        #   Predefined templates: https://github.com/OpenCageData/address-formatting/blob/master/conf/countries/worldwide.yaml
        AddressTemplate = $null
    }

    if ($UseHtmTemplates) {
        $ReplaceHash["`$$($ReplacementVariableNamespace)PostalAddressNoCompany`$"] = [System.Net.WebUtility]::HtmlEncode((Format-PostalAddress @FormatPostAddressOptions)) -replace '\r?\n', '<br />' # Converts paragraphs to line breaks
    } else {
        $ReplaceHash["`$$($ReplacementVariableNamespace)PostalAddressNoCompany`$"] = (Format-PostalAddress @FormatPostAddressOptions) -replace '\r?\n', "`n" # Converts paragraphs to line breaks
    }
}
