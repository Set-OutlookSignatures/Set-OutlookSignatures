# Script to embed locally available files into single HTML file as Base64 strings
# Can only embed files stored on the local computer


# Input HTML file or folder
# If folder, all .htm and .html files directly in this folder are considered
# Every single HTML file must be UTF-8 encoded.
$inputPath = '.\source'


function main {
    Write-Host 'Script started.'

    if (($ExecutionContext.SessionState.LanguageMode) -eq 'FullLanguage') {
        $fullLanguageMode = $true
    } else {
        $fullLanguageMode = $false
        Write-Host "  This PowerShell session is in $($ExecutionContext.SessionState.LanguageMode) mode, not FullLanguage mode."
        Write-Host '  Base64 conversion not possible. Exiting.'
        Write-Host 'Script completed.'
        exit
    }

    $inputFiles = @()
    if (Test-Path $inputPath -PathType leaf) {
        $inputFiles += (Get-ChildItem $inputPath -File).FullName
    } elseif (Test-Path -LiteralPath $inputPath -PathType Container) {
        Get-ChildItem -Include '*.html', '*.htm' -Exclude '* - Single File.html', '* - Single File.htm' -LiteralPath $inputPath -Depth 0 -File | ForEach-Object {
            $inputFiles += $_.FullName
        }
    } else {
        Write-Host "  Folder or file `"$inputPath`" not found, exiting."
        Write-Host 'Script completed.'
        exit
    }

    Write-Host "  Found $($inputFiles.count) files in `"$inputPath`"."

    $currentFileCount = 0
    foreach ($inputFile in $inputFiles) {
        $inputFile = Get-ChildItem -LiteralPath $inputFile
        $currentFileCount++
        $html = Get-Content -LiteralPath $inputFile -Raw -Encoding UTF8
        Write-Host "    File $currentFileCount`: $inputfile"

        if ($html.Contains([char]0xfffd)) {
            Write-Host '      File is not UTF-8 encoded or contains byte sequences not valid in UTF-8, ignoring file.'
            continue
        }
        
        $src = @()
        ([regex]'(?i)src="(.*?)"').Matches($html) |  ForEach-Object {
            $src += $_.Groups[0].Value
            $src += (Join-Path -Path (Split-Path -Path $inputFile -Parent) -ChildPath ([uri]::UnEscapeDataString($_.Groups[1].Value)))
        }

        Write-Host "      Found $($src.count / 2) `"src=`" tags."

        for ($i = 0; $i -lt $src.count; $i = $i + 2) {
            Write-Host "        Tag $(($i / 2) + 1): " -NoNewline
            if ($src[$i].StartsWith('src="data:')) {
                Write-Host "$($src[$i].substring(0,50))[...] is already a data URI. Ignoring tag."
            } elseif (Test-Path -LiteralPath $src[$i + 1] -PathType leaf) {
                Write-Host "$($src[$i]) is available locally, " -NoNewline
                $fmt = $null
                switch ((Get-ChildItem -LiteralPath $src[$i + 1]).Extension) {
                    '.apng' { $fmt = 'data:image/apng;base64,' }
                    '.avif' { $fmt = 'data:image/avif;base64,' }
                    '.gif' { $fmt = 'data:image/gif;base64,' }
                    '.jpg' { $fmt = 'data:image/jpeg;base64,' }
                    '.jpeg' { $fmt = 'data:image/jpeg;base64,' }
                    '.jfif' { $fmt = 'data:image/jpeg;base64,' }
                    '.pjpeg' { $fmt = 'data:image/jpeg;base64,' }
                    '.pjp' { $fmt = 'data:image/jpeg;base64,' }
                    '.png' { $fmt = 'data:image/png;base64,' }
                    '.svg' { $fmt = 'data:image/svg+xml;base64,' }
                    '.webp' { $fmt = 'data:image/webp;base64,' }
                    '.css' { $fmt = 'data:text/css;base64,' }
                    '.less' { $fmt = 'data:text/css;base64,' }
                    '.js' { $fmt = 'data:text/javascript;base64,' }
                    '.otf' { $fmt = 'data:font/otf;base64,' }
                    '.sfnt' { $fmt = 'data:font/sfnt;base64,' }
                    '.ttf' { $fmt = 'data:font/ttf;base64,' }
                    '.woff' { $fmt = 'data:font/woff;base64,' }
                    '.woff2' { $fmt = 'data:font/woff2;base64,' }
                }
                if ($fmt) {
                    Write-Host 'embedding as base64.'
                    $html = $html.replace( `
                            $src[$i], `
                        ('src="' + $fmt + [Convert]::ToBase64String([IO.File]::ReadAllBytes($src[$i + 1])) + '"') `
                    )

                } else {
                    Write-Host "but $((Get-ChildItem -LiteralPath $src[$i+1]).Extension) is not a supported extension. Ignoring."
                }
            } else {
                Write-Host "$($src[$i]) is not available locally. Ignoring tag."
            }
        }
    
        $outputFile = (Join-Path -Path (Split-Path -Path $inputFile -Parent) -ChildPath ($inputFile.BaseName + ' - Single File' + $inputFile.Extension))
        Write-Host "      Writing `"$outputFile`"."
        $html | Out-File -LiteralPath $outputFile -Force -Encoding utf8
    }

    Write-Host 'Script completed.'
}


main