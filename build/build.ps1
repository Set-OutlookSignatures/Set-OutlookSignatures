function main {
    Write-Output 'Basics'
    Set-Location $env:GITHUB_WORKSPACE

    & choco.exe install pandoc --no-progress
    refreshenv

    if ($env:RELEASETAG) {
        $ReleaseTag = $env:RELEASETAG
    } else {
        $ReleaseTag = ($env:GITHUB_REF -replace 'refs/tags/', '')
        "RELEASETAG=$ReleaseTag" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
    }

    if ($env:RELEASEFILE) {
        $ReleaseFile = $env:RELEASEFILE
    } else {
        $ReleaseFile = ($env:GITHUB_REPOSITORY -split '/')[1] + '_' + $ReleaseTag + '.zip'
        "RELEASEFILE=$ReleaseFile" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
    }

    if ($env:RELEASENAME) {
        $ReleaseName = $env:RELEASENAME
    } else {
        $ReleaseName = "Release $ReleaseTag"
        "RELEASENAME=$ReleaseName" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
    }

    $BuildDir = $('./build/' + ($env:GITHUB_REPOSITORY -split '/')[1]) + '_' + $ReleaseTag
    New-Item $BuildDir -ItemType 'directory' | Out-Null

    Write-Output "BuildDir: $BuildDir"
    Write-Output "ReleaseFile: $ReleaseFile"
    Write-Output "ReleaseName: $ReleaseName"
    Write-Output "ReleaseTag: $ReleaseTag"


    Write-Output 'Copy basic files'
    Set-Location $env:GITHUB_WORKSPACE

    Copy-Item '.\src\*' $BuildDir -Recurse
    Copy-Item '.\LICENSE.txt' "$BuildDir\docs\LICENSE.txt" -Force


    Write-Output 'Convert markdown files to HTML and copy them'
    Set-Location $env:GITHUB_WORKSPACE

    @(
        ('.\docs\CHANGELOG.md', "$BuildDir\docs\CHANGELOG.html"),
        ('.\docs\CODE_OF_CONDUCT.md', "$BuildDir\docs\CODE_OF_CONDUCT.html"),
        ('.\docs\CONTRIBUTING.md', "$BuildDir\docs\CONTRIBUTING.html"),
        ('.\docs\Implementation approach.md', "$BuildDir\docs\Implementation approach.html"),
        ('.\docs\README.md', "$BuildDir\docs\README.html")
    ) | ForEach-Object {
        & pandoc.exe $($_[0]) --resource-path=".;docs" -f gfm -t html --self-contained -H .\build\pandoc_header.html --css .\build\pandoc_css_empty.css --metadata pagetitle="$(([System.IO.FileInfo]"$($_[0])").basename) - Set-OutlookSignatures" -o $($_[1])
    }


    Write-Output 'Update version number in script'
    Set-Location $env:GITHUB_WORKSPACE

    Set-Location $BuildDir

    ((Get-Content Set-OutlookSignatures.ps1 -Raw) -replace 'xxxVersionStringxxx', $ReleaseTag) | Set-Content Set-OutlookSignatures.ps1


    Write-Output 'Create file hashes and place them in file hashes.txt'
    Set-Location $env:GITHUB_WORKSPACE

    Set-Location $BuildDir

    Remove-Item 'hashes.txt' -Force

    $Hashes = ForEach ($File in (Get-ChildItem -File -Recurse)) {
        Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256 | Select-Object @{N = 'PathRelative'; E = { Resolve-Path -LiteralPath $file.FullName -Relative } }, Algorithm, Hash
    }

    $Hashes | Export-Csv hashes.txt


    Write-Output 'Create release file'
    Set-Location $env:GITHUB_WORKSPACE

    Compress-Archive $BuildDir $ReleaseFile


    Write-Output 'Output additional information for release'
    Set-Location $env:GITHUB_WORKSPACE

    $Changelog = '.\docs\changelog.md'
    $ChangeLogLines = Get-Content $Changelog
    $ChangelogStartline = $null
    $ChangelogEndline = $null
    $ReleaseTagDate = $(Get-Date -Format 'yyyy-MM-dd')
    for ($i = 0; $i -lt $ChangeLogLines.count; $i++) {
        if (-not $ChangelogStartline) {
            if ($ChangeLogLines[$i] -match ("^##\s*(\[$ReleaseTag\] - $ReleaseTagDate\s*|$ReleaseTag - $ReleaseTagDate\s*|$ReleaseTag - $ReleaseTagDate$)|>$ReleaseTag</a> - $ReleaseTagDate\s*")) {
                $ChangelogStartline = $i
                continue
            }
        } else {
            if (($ChangeLogLines[$i]).startswith('## ')) {
                $ChangelogEndline = $i - 1
                break
            }
        }
    }
    if (-not $ChangelogStartline) {
        $ReleaseMarkdown = "# **Tag '$ReleaseTag - $ReleaseTagDate' not found in '$Changelog', using first entry.**`r`n"
        $ChangelogStartline = $null
        $ChangelogEndline = $null
        for ($i = 0; $i -lt $ChangeLogLines.count; $i++) {
            if (-not $ChangelogStartline) {
                if (($ChangeLogLines[$i]).startswith('## ')) {
                    $ChangelogStartline = $i
                }
            } else {
                if (($ChangeLogLines[$i]).startswith('## ')) {
                    $ChangelogEndline = $i - 1
                    break
                }
            }
        }
    } else {
        if (-not $Changelogendline) { $ChangelogEndline = $ChangelogLines.count - 1 }
    }
    for ($i = $ChangelogStartline; $i -le $ChangelogEndline; $i++) {
        $ChangeLogLines[$i] = $ChangeLogLines[$i] -replace '^##', '#'
    }
    $ReleaseMarkdown = $ReleaseMarkdown + ($($ChangeLogLines[$ChangelogStartline..$ChangelogEndline]) -join "`r`n")

    if ($RegExMatches = [regex]::matches($ReleaseMarkdown, '\[(.*?)\]')) {
        for ($i = 0; $i -lt $ChangeLogLines.count; $i++) {
            foreach ($m in $RegExMatches) {
                if ($ChangeLogLines[$i].StartsWith("$($m.value):")) {
                    $ReleaseMarkdown = $ReleaseMarkdown + "`r`n$($ChangeLogLines[$i])"
                }
            }
        }
    }
    
    $ReleaseMarkdown = $ReleaseMarkdown + @"
`r`n# File hashes
- SHA256 hash of '$ReleaseFile': $((Get-FileHash $ReleaseFile -Algorithm SHA256).hash)
- See 'hashes.txt' in '$ReleaseFile' for hash value of every single file in the release.
"@

    Write-Output 'ReleaseMarkdown:'
    Write-Output $ReleaseMarkdown
    $ReleaseMarkdown | Out-File -FilePath .\build\CHANGELOG.md -Encoding utf8 -Force
}

if ($env:GITHUB_WORKSPACE) {
    main
} else {
    throw 'This script is designed to run as part of a GitHub workflow, and does not work elsewhere.'
}
