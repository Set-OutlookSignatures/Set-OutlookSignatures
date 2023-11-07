##[String]$location = Split-Path -Parent $PSCommandPath
[String]$temp = [environment]::getfolderpath('TEMP')
New-Item -Path $temp -Name "set-outlooksignatures" -ItemType "directory" -erroraction SilentlyContinue
$destination = $temp + "\set-outlooksignatures"



function DownloadGitHubRepository 
{ 
    param( 
       [Parameter(Mandatory=$True)] 
       [string] $Name, 
         
       [Parameter(Mandatory=$True)] 
       [string] $Author, 

       [Parameter(Mandatory=$False)] 
       [string] $RepositoryZipUrl = "https://github.com/alltimeuk/emailsignatures/archive/master.zip", 
         
       [Parameter(Mandatory=$False)] 
       [string] $Branch = "master", 
         
       [Parameter(Mandatory=$False)] 
       [string] $Location = "C:\temp"
    ) 
     
    # Force to create a zip file 
    $ZipFile = "$location\$Name.zip"
    New-Item $ZipFile -ItemType File -Force
 
    # download the zip 
    Write-Host 'Starting download from GitHub'
    Invoke-RestMethod -Uri $RepositoryZipUrl -OutFile $ZipFile
    Write-Host 'Download finished'
 
    #Extract Zip File
    Write-Host 'Starting unzipping of $Name.zip'
    Expand-Archive -Path $ZipFile -DestinationPath $location -Force
    Write-Host 'Unzip finished here: $Location'
     
    # remove the zip file
    Remove-Item -Path $ZipFile -Force
}

#Download
DownloadGitHubRepository -Name "Set-OutlookSignatures" -Author "Simon Jackson @ Alltime Technologies Ltd" -RepositoryZipUrl "https://github.com/alltimeuk/EmailSignatures/archive/refs/heads/main.zip" -location "$destination"

#Run
.\$destination\src_set-OutlookSignatures\Set-OutlookSignatures.ps1 -graphonly true -SignatureTemplatePath .\$destination\private\Signatures -SignatureIniPath .\$destination\private\Signatures\_Signatures.ini -SetCurrentUserOOFMessage false -CreateRtfSignatures true -CreateTxtSignatures true -DisableRoamingSignatures false -MirrorLocalSignaturesToCloud true