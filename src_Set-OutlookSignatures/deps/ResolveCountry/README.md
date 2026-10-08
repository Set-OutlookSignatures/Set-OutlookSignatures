## ResolveCountry PowerShell Module

### Overview
  
ResolveCountry is a PowerShell module for resolving country names, country codes, localized names, and common user-entered variants into standardized country information.  
The module combines official country data, native names, translated names, alternative spellings, and fuzzy matching to reliably identify countries even when input data contains typos, abbreviations, localized names, or inconsistent formatting.  
Typical use cases include:
- Normalizing country values from Active Directory, Entra ID, HR systems, CRM platforms, and CSV imports.
- Converting country names to ISO 3166-1 country codes.
- Validating and standardizing user-entered country information.
- Supporting address formatting, phone number normalization, and geographic data processing.
- Building automation workflows that require consistent country identifiers.  
To provide fast lookups, the module maintains an in-memory cache and a compressed local data store while remaining fully compatible with both Windows PowerShell 5.1 and PowerShell 7+.

### Used By
  
ResolveCountry is used by [Set-OutlookSignatures](https://set-outlooksignatures.com) for country normalization and ISO code resolution, but can be used independently in any PowerShell automation scenario requiring reliable country matching and standardization.

### Acknowledgements
  
This module uses the comprehensive country dataset provided by the [**mledoze/countries**](https://github.com/mledoze/countries) repository. Thank you to the contributors of that project for maintaining such high-quality, normalized geographical data.

### Compatibility
  
Fully compatible with **PowerShell 5.1** (Windows) and **PowerShell 7+** (Cross-platform: Windows, macOS, Linux).

### Features
- **High Performance:** Caches data in an in-memory hashtable at the module-scope immediately upon import.
- **Disk Caching:** Compresses lookup arrays into a ResolveCountry.Data.json.gz file locally to avoid re-downloading standard data.
- **Fuzzy Matching:** Uses a C# Levenshtein distance compiler for rapid string normalization and typo-correction.
- **Versatile Outputs:** Can return ISO 3166-1 alpha-2 (cca2), alpha-3 (cca3), numeric codes (ccn3), native names, or the entire JSON object.

### Installation & Setup
- Create a folder named ResolveCountry in your PowerShell Modules directory (e.g., C:\Program Files\WindowsPowerShell\Modules\ResolveCountry for PS5.1 or ~\Documents\PowerShell\Modules\ResolveCountry for PS7).
- Place ResolveCountry.psm1 inside this folder.
- Import the module into your session:

```powershell
Import-Module ResolveCountry
```

_Note: The first time you import the module, it will automatically download the remote dataset and generate the compressed country_lookup.json.gz cache. Subsequent imports will read directly from the disk cache._

### Usage

#### Resolve-Country

Takes a query string and searches for the corresponding country.

**Basic lookup (defaults to returning ISO cca2 code):**

```powershell
Resolve-Country "United States of America"
# Output: US

Resolve-Country "Deutschland"
# Output: DE
```

**Handling typos using fuzzy matching:**

```powershell
Resolve-Country "Swtizerland"
# Output: CH

Resolve-Country "uni tedStates aMerica"
# Output: US
```

**Returning different data types:**

```powershell
Resolve-Country "Japan" -ReturnType 'cca3'
# Output: JPN

Resolve-Country "Austria" -ReturnType 'name.native.official'
# Output: Republik Österreich
```

#### Custom Countries & Aliases

You can extend the built-in country data with custom countries and aliases.

**1. Define custom country data and aliases**

```powershell
$customCountries = @{
    'XX' = [pscustomobject]@{
        cca2 = 'XX'; cca3 = 'XXX'; ccn3 = '999'
        name = [pscustomobject]@{
            common   = 'Testlandia'
            official = 'Republik Testland'
        }
    }
}

$customAliases = @{
    'Niederösterreich' = 'AT'
    'NÖ'               = 'AT'
    'Testi'            = 'XX'
}
```

**2. Register them in memory**

```powershell
Set-CountryLookupConfig -CustomCountries $customCountries -CustomAliases $customAliases
```

**3. Use them**

```powershell
Resolve-Country 'Niederösterreich'
# Output: AT

Resolve-Country 'Testi' -ReturnType 'name.official'
# Output: Republik Testland

'Niederösterreich', 'NÖ', 'Testi' | Resolve-Country
# Output: AT, AT, XX
```

#### Update-ResolveCountryData

Forces a manual update of the local disk cache by polling the upstream repository again.

```powershell
Update-ResolveCountryData
```
