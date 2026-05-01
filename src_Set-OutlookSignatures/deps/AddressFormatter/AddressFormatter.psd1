# Module Manifest for AddressFormatter
@{
    # Version and Metadata
    ModuleVersion     = '1.0.0'
    GUID              = '6602b7c2-0465-4496-829b-e7aedf30b8b3'
    Author            = 'Markus Gruber'
    Copyright         = '(c) 2025--present Markus Gruber. All rights reserved.'

    # Function to export (only Format-PostalAddress is exposed to the user)
    FunctionsToExport = 'Format-PostalAddress'

    # Dependencies
    nestedModules     = @(
        'nestedModules\powershell-yaml\powershell-yaml.psm1'
    )

    # Root module script
    RootModule        = 'AddressFormatter.psm1'
}