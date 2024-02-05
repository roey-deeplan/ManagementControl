$AppCatalogURL = "" # https://alomone.sharepoint.com example
$SppkgFile = Get-ChildItem -Path "../sharepoint/solution" #"../sharepoint/solution/management-control.sppkg"
$AppFilePath = "../sharepoint/solution/" + [string]$SppkgFile
Write-Host $AppFilePath
$UserName = ""
$Password = ""

$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $UserName, $SecurePassword

# Connect to SharePoint Online App Catalog site
try {
    Connect-PnPOnline -Url $AppCatalogURL -Credentials $Cred
} catch {
    Write-Host "Error connecting to SharePoint Online: $_"
    exit
}

$content = Get-ChildItem -Path "../src"

# Webpart - Deploy to all sites, Extention - Deploy to one site
if ([string]$content[0] -eq "webparts") {
    Add-PnPApp -Path $AppFilePath -Scope Tenant -Publish -SkipFeatureDeployment
    Write-Host "Deployed webpart to all sites"
} else {
    Add-PnPApp -Path $AppFilePath -Scope Tenant -Publish
    Write-Host "Deployed extention"
}
