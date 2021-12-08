#Parameters
#https://dwpstage.sharepoint.com
$AppCatalogURL = "https://8p5g5n.sharepoint.com/sites/appcatalog"
#$AppCatalogURL = "https://dwpstage.sharepoint.com/sites/appcatalog"
$AppFilePath = "C:\Users\adminpen.arpula\spfxclientsideprojects\spfx-ecb-extension\sharepoint\solution\spfx-ecb-extension.sppkg"
 
#Connect to SharePoint Online App Catalog site
Connect-PnPOnline -Url $AppCatalogURL -UseWebLogin 
 
#Add App to App catalog - upload app to sharepoint online app catalog using powershell
#$App = Add-PnPApp -Path $AppFilePath

#$App =Add-PnPApp -Path $AppFilePath -Scope Tenant -Publish

$App =Add-PnPApp -Path $AppFilePath -Publish -SkipFeatureDeployment -Overwrite

'Added ' + $App.Id
 
#Deploy App to the Tenant
#Publish-PnPApp -Identity $App.Id -Scope Tenant


