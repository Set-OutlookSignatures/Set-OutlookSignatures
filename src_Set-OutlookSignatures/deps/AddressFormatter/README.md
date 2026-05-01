# AddressFormatter
Address formatting for PowerShell using the templates from https://github.com/OpenCageData/address-formatting.

Works cross-platform: PowerShell 5.1 on Windows, PowerShell (pwsh) 7+ on Windows, Linux, and macOS

Based on Perl implementation https://metacpan.org/dist/Geo-Address-Formatter.

# Usage
1. Import module  
   `Import-Module 'c:\your_modules_path\AddressFormatter\AddressFormatter.psd1'`
2. Format address using `Format-PostalAddress`
   ```powershell
   $FormatPostAddressOptions = @{
       # Address components as described in https://github.com/OpenCageData/address-formatting/blob/master/conf/components.yaml
       Components = @{
       attention = 'Text for attention line'
       road      = 'Name of the road'
       city      = 'Name of the city'
       postcode  = 'Postcode'
       state     = 'Name of the state'
       country   = 'Name of the country'
    }

    # Country as two-letter ISO country code (e.g., "AT", "US") is needed to choose correct address format rules
    Country = 'AT'

    # Shorten address components ("St." instead of "Street", "Rd." instead of "Road", etc.)
    #   Depends on Country attribute
    Abbreviate = $false

    # Only return known parts of the address, omit unknown parts
    #   When disabled, unknown parts are added the the "attention" component
    OnlyAddress = $false

    # Use a custom address template instead of the predefined ones
    #   Predefined templates: https://github.com/OpenCageData/address-formatting/blob/master/conf/countries/worldwide.yaml
    AddressTemplate = $null
   }

   Format-PostalAddress @FormatPostAddressOptions
   ```

# Usage examples
Note the subtle country specific formatting differences in the following examples, such as where the house number and the postcode is placed.

## Austrian Presidential Office
```powershell
$FormatPostAddressOptions = @{
    # Address components as described in https://github.com/OpenCageData/address-formatting/blob/master/conf/components.yaml
    Components = @{
        attention = @(
            'Bürger:innenservice'
            'Österreichische Präsidentschaftskanzlei'
        ) -join [System.Environment]::NewLine
        house        = 'Hofburg'
        house_number = 1
        road         = 'Ballhausplatz'
        city         = 'Wien'
        postcode     = '1010'
        state        = ''
        country      = 'Austria'
    }
    Country    = 'AT'
}

Format-PostalAddress @FormatPostAddressOptions
```

The PowerShell commands above give the following result:
```
Bürger:innenservice
Österreichische Präsidentschaftskanzlei
Hofburg
Ballhausplatz 1
1010 Wien
Austria
```
## USA White House
```powershell
$FormatPostAddressOptions = @{
    # Address components as described in https://github.com/OpenCageData/address-formatting/blob/master/conf/components.yaml
    Components = @{
        house        = 'The White House'
        house_number = 1600
        road         = 'Pennsylvania Avenue, N.W.'
        city         = 'Washington'
        postcode     = '20500'
        state        = 'Washington, DC'
        country      = 'USA'
    }
    Country    = 'US'
}

Format-PostalAddress @FormatPostAddressOptions
```

The PowerShell commands above give the following result:
```
The White House
1600 Pennsylvania Avenue, N.W.
Washington, DC 20500
United States of America
```

Note that the country name 'USA' has been automatically corrected to 'United States of America'.

# Admin tasks
## Update address templates and test cases
Address templates and test cases are stored in the folder '`subModules\OpenCageData\address-formatting`'. This folder is a git submodule, which is basically a clone of the repository '`https://github.com/OpenCageData/address-formatting`'.

To update the templates and test cases, run '`git submodule update`'.

## Run tests
Run `.\tests\tests.ps1` to run the tests. Only errors and a summary will be logged.

Example output:
```
Start script @2026-03-06T15:17:40+01:00@

Import modules
  AddressFormatter
  powershell-yaml

Submodule OpenCageData/address-formatting
  Commit 064d82b, dated 2026-02-03T17:32:05+01:00

Running test cases
  465 test cases from 256 files

Test results
  Passed: 465/465 (100,00 %)
  Failed: 0/465 (0,00 %)

End script @2026-03-06T15:17:50+01:00@
```
