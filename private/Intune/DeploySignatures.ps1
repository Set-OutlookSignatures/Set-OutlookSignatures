##[String]$location = Split-Path -Parent $PSCommandPath
[String]$temp = [environment]::getfolderpath('TEMP')

function DownloadGitHubRepository 
{ 
    param( 
       [Parameter(Mandatory=$True)] 
       [string] $Name, 
        
       [Parameter(Mandatory=$False)]
       [string] $RepositoryZipUrl = "",

       [Parameter(Mandatory=$False)] 
       [string] $Owner = "Set-OutlookSignatures", 

       [Parameter(Mandatory=$False)] 
       [string] $RepoName = "Set-OutlookSignatures", 

       [Parameter(Mandatory=$False)] 
       [string] $Branch = "master", 

       [Parameter(Mandatory=$False)] 
       [string] $Location = "C:\temp"
    ) 
     
    # Force to create a zip file 
    $ZipFile = "$location\$Name.zip"
    New-Item $ZipFile -ItemType File -Force
 
    # download the zip 
    #$RepositoryZipUrl = "https://github.com/[Owner]/[repoName]/archive/[Branch].zip", 
    #$RepositoryZipUrl = "https://github.com/$Owner/$RepoName/archive/$Branch.zip"
    $RepositoryZipUrl = "https://github.com/alltimeuk/EmailSignatures/archive/refs/tags/1.0.0.zip"
    Write-Host 'Starting download from GitHub'
    Invoke-RestMethod -Uri $RepositoryZipUrl -OutFile $ZipFile
    Write-Host 'Download finished'
 
    #Extract Zip File
    Write-Host 'Starting unzipping of $Name.zip'
    Expand-Archive -Path $ZipFile -DestinationPath $location -Force
    Get-ChildItem  -path $location
    Write-Host 'Unzip finished here: $Location'
     
    # remove the zip file
    #Remove-Item -Path $ZipFile -Force
}

#Download
DownloadGitHubRepository -Name "OutlookSignatures" -Owner "alltimeuk" -RepoName "EmailSignatures" -branch "main" -location "$temp"

#Run
.\$temp\src_set-OutlookSignatures\Set-OutlookSignatures.ps1 -graphonly true -SignatureTemplatePath .\$temp\private\Signatures -SignatureIniPath .\$temp\private\Signatures\_Signatures.ini -SetCurrentUserOOFMessage false -CreateRtfSignatures true -CreateTxtSignatures true -DisableRoamingSignatures false -MirrorLocalSignaturesToCloud true