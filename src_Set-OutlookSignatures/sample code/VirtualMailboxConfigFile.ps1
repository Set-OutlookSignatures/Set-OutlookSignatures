<#
This sample code shows how to define the virtual mailboxes and additional signature INI entries for use in Set-OutlookSignatures.

You have to adapt it to fit your environment.
The sample code is written in a generic way, which allows for easy adaption.

Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers fee-based support for this and other open source code.

Virtual mailboxes are mailboxes that are not available in Outlook but are treated by Set-OutlookSignatures as if they were.
This is an option for scenarios where you want to deploy signatures with not only the $CurrentUser...$ but also
   $CurrentMailbox...$ replacement variables for mailboxes that have not been added to Outlook, such as in Send As or
   Send On Behalf scenarios, where users often only change the from address but do not add the mailbox to Outlook.

This script is executed as a whole once per Set-OutlookSignatures run.

Attention: The configuration file is executed as part of Set-OutlookSignatures.ps1 and is not checked for any harmful content. Please only allow qualified technicians write access to this file, only use it to to define replacement variables and test it thoroughly.

A variable defined in this file overrides the definition of the same variable defined earlier in the software.

$ADPropsMailboxes is an array containing the properties of each mailbox Set-OutlookSignatures has found before.
$ADPropsMailboxManagers is an array containing the properties of the manager of each mailbox Set-OutlookSignatures has found before.
  The index number of the two arrays match: $ADPropsMailboxManagers[0] contains the properties of the manager of the mailbox $ADPropsMailboxes[0].
  The arrays are not unique, so the same mailbox or the same manager can appear multiple times (as in different Outlook profiles, for example).
  Available attributes:
    See '.\config\default replacement variables' for examples.
    '$ADPropsMailboxes[0] | fl *' lists all available attributes.
    GroupsSIDs is only available for $ADPropsMailboxes, contact us if you need it for $ADPropsMailboxManagers.

$ADPropsCurrentUser is an object containing the properties of the currently logged-in user.
$ADPropsCurrentUserManager is an object containing the properties of the manager of the currently logged-in user.
  The attributes are identical to the attributes of $ADPropsMailboxes and $ADPropsMailboxManagers.
    GroupsSIDs is only available for $ADPropsCurrentUser, contact us if you need it for $ADPropsCurrentUserManager.

To add a virtual mailbox, add the SMTP address to the $VirtualMailboxesToAdd array:
  $VirtualMailboxesToAdd += 'first.last@example.com'
Virtual mailboxes are added to the list of mailboxes in the sequence as they have been added to $VirtualMailboxesToAdd.
Only those virtual mailboxes are added that are not already in the list of mailboxes.

To add additional signature INI entries, add them to the $SignatureIniAdditionalLines string:
  $SignatureIniAdditionalLines += "[template file.docx]"
  $SignatureIniAdditionalLines += "option A"
  $SignatureIniAdditionalLines += "option B"

To add additional OOF INI entries, add them to the $OOFIniAdditionalLines string:
  $OOFIniAdditionalLines += "[template file.docx]"
  $OOFIniAdditionalLines += "option A"
  $OOFIniAdditionalLines += "option B"
#>


# Example: Always add mailbox a@example.com
Write-Host '      Always add mailbox a@example.com and assign it template "External formal Delegate.docx"'
@('a@example.com') | ForEach-Object {
    $VirtualMailboxesToAdd += $_

    # Add an additional signature INI entry
    #   Applies the template to the mailbox but only when it is not the one of the logged-in user
    #   Sets an individual signature name
    $SignatureIniAdditionalLines += '[External formal Delegate.docx]'
    $SignatureIniAdditionalLines += "$($_)"
    $SignatureIniAdditionalLines += "-CURRENTUSER:$($_)"
    $SignatureIniAdditionalLines += "OutlookSignatureName = External formal Delegate $($_)"
}
Write-Host '        Done'


# Example: If the mailbox b@example.com is in Outlook, add mailbox c@example.com
Write-Host '      If mailbox b@example.com is in Outlook: Add virtual mailbox c@example.com'
if ($ADPropsMailboxes.proxyaddresses -icontains 'smtp:b@example.com') {
    Write-Host '        Condition met'
    $VirtualMailboxesToAdd += 'c@example.com'
} else {
    Write-Host '        Condition not met'
}


# Example: All users with the manager c@example.com get added the virtual mailbox d@example.com
Write-Host '      All users with the manager c@example.com: Add virtual mailbox d@example.com'
if ($ADPropsCurrentUserManager.proxyaddresses -icontains 'smtp:c@example.com') {
    Write-Host '        Condition met'
    $VirtualMailboxesToAdd += 'd@example.com'
} else {
    Write-Host '        Condition not met'
}


# Example: If the current user is a member of the Entra group e@example.com, add the virtual mailbox f@example.com
Write-Host '      Current user is member of the Entra group e@example.com: Add virtual mailbox f@example.com'
if ($ADPropsCurrentUser.GroupsSIDs -icontains $(ResolveToSid('EntraID e@example.com'))) {
    Write-Host '        Condition met'
    $VirtualMailboxesToAdd += 'f@example.com'
} else {
    Write-Host '        Condition not met'
}


# Example: Use data from Export-RecipientPermissions stored on SharePoint Online
#   This is especially helpful if you want to automate as much of the process as possible.
#   Visit https://github.com/Export-RecipientPermissions for details about Export-RecipientPermissions.
#
#   Find all entries where the current user has SendAs or SendOnBehalf permissions
#   and add the mailboxes of the grantors to the list of virtual mailboxes.
#
# For best results use Export-RecipientPermissions with the following settings:
#   Enable 'ExpandGroups' so that groups granted a permission are resolved to their members.
#   Optionally, you can prepare the export of Export-RecipientPermissions so that the code in this file runs faster:
#     Only export SendAs and SendOnBehalf permissions by setting the according parameters.
#     Only export entries for grantors that are mailboxes by defining a GrantorFilter.
#     Only export allow permissions by defining an ExportFileFilter.
#     Modify the export to remove SendAs permissions when there is also a SendOnBehalf permission for the same combination of grantor and trustee.
#
# Attention: Code is commented out because a non-accessible path will cause Set-OutlookSignatures to exit.
<#
Write-Host '      SendAs and SendOnBehalf from Export-RecipientPermissions'
Write-Host '        Downloading file from SharePoint Online'
$ExportRecipientPermissionsFile = 'https://example.sharepoint.com/sites/library/path/file.csv'
ConvertPath ([ref]$ExportRecipientPermissionsFile)
. $CheckPathScriptblock -CheckPathRefPath ([ref]$ExportRecipientPermissionsFile) -ExpectedPathType 'Leaf'
Write-Host '        Importing and filtering'

$ExportRecipientPermissions = Import-Csv -LiteralPath $ExportRecipientPermissionsFile -Encoding utf8 -Delimiter ';'

# Get SendAs and SendOnBehalf allows for current user only and save grantors for SendAs and SendOnBehalf in separate arrays
#   Separate the arrays because Exchange prioritizes SendAs over SendOnBehalf
$ExportRecipientPermissionsFilteredSendAsGrantors = @()
$ExportRecipientPermissionsFilteredSendOnBehalfGrantors = @()

foreach ($ExportRecipientPermission in $ExportRecipientPermissions) {
    if (
        $("smtp:$($ExportRecipientPermission.'Trustee Primary SMTP')" -iin $ADPropsCurrentUser.proxyaddresses) -and
        $($ExportRecipientPermission.'Permission' -iin @( 'SendAs', 'SendOnBehalf')) -and
        $($ExportRecipientPermission.'Allow/Deny' -ieq 'Allow') -and
        $($ExportRecipientPermission.'Grantor Recipient Type' -inotlike '*group')
    ) {
        if ($ExportRecipientPermission.'Permission' -ieq 'SendAs') {
            $ExportRecipientPermissionsFilteredSendAsGrantors += $ExportRecipientPermission.'Grantor Primary SMTP'
        } elseif ($ExportRecipientPermission.'Permission' -ieq 'SendOnBehalf') {
            $ExportRecipientPermissionsFilteredSendOnBehalfGrantors += $ExportRecipientPermission.'Grantor Primary SMTP'
        }
    }
}

# As Exchange prioritizes SendAs over SendOnBehalf, remove SendOnBehalf grantors that are also SendAs grantors
$ExportRecipientPermissionsFilteredSendOnBehalfGrantors = @(
    $ExportRecipientPermissionsFilteredSendOnBehalfGrantors | Where-Object {
        $_ -inotin $ExportRecipientPermissionsFilteredSendAsGrantors
    }
)

# Add virtual mailboxes and additional INI lines for SendAs
$ExportRecipientPermissionsFilteredSendAsGrantors | ForEach-Object {
    Write-Host "          Found $($_) (SendAs)"
    $VirtualMailboxesToAdd += $_
    $SignatureIniAdditionalLines += '[External formal Delegate.docx]'
    $SignatureIniAdditionalLines += "$($_)"
    $SignatureIniAdditionalLines += "-CURRENTUSER:$($_)"
    $SignatureIniAdditionalLines += "OutlookSignatureName = External formal SendAs $($_)"
}

# Add virtual mailboxes and additional INI lines for SendOnBehalf
$ExportRecipientPermissionsFilteredSendOnBehalfGrantors | ForEach-Object {
    Write-Host "          Found $($_) (SendOnBehalf)"
    $VirtualMailboxesToAdd += $_
    $SignatureIniAdditionalLines += '[External formal Delegate SendOnBehalf.docx]'
    $SignatureIniAdditionalLines += "$($_)"
    $SignatureIniAdditionalLines += "-CURRENTUSER:$($_)"
    $SignatureIniAdditionalLines += "OutlookSignatureName = External formal SendOnBehalf $($_)"
}
#>