
$AppCatalogURL = "https://projects1.sharepoint.com"
$AppFilePath = "./sharepoint/solution/management-control.sppkg"

#Connect to SharePoint Online App Catalog site
Connect-PnPOnline -Url $AppCatalogURL 

#Add App to App catalog - upload app to sharepoint online app catalog using powershell
$App = Add-PnPApp -Path $AppFilePath -SkipFeatureDeployment -Publish
 
# Check if $App is not null before trying to publish
if ($App -ne $null) {
    Publish-PnPApp -Identity $App.Id -Scope Tenant
    Write-Host "Success to add the app to the app catalog."
} else {
    Write-Host "Failed to add the app to the app catalog."
}