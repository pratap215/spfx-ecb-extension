$AppCatalogURL = "https://8p5g5n.sharepoint.com/sites/appcatalog"
$AppName = "spfx-ecb-extension-client-side-solution"

Connect-PnPOnline -Url $AppCatalogURL -UseWebLogin 

$App = Get-PnPApp -Scope Tenant | Where {$_.Title -eq $AppName}

'Removing '  + $App.Id
 
#Remove an App from App Catalog Site
Remove-PnPApp -Identity $App.Id

'Removed ############################################################################################'  
Write-Output "   "
Write-Output "   "

gulp bundle --ship;
Write-Output 'gulp bundle completed-------------------------------------------------------------------'
Write-Output "   "
gulp package-solution --ship;
Write-Output "gulp package-solution --ship completed ***************************************************"
Write-Output "   "

Write-Output "(((((((((((((((Starting App catalog Deploy))))))))))))))))))"


#$AppCatalogURL = "https://dwpstage.sharepoint.com/sites/appcatalog"
$AppFilePath = "C:\Users\adminpen.arpula\spfxclientsideprojects\spfx-ecb-extension\sharepoint\solution\spfx-ecb-extension.sppkg"
 
#Connect to SharePoint Online App Catalog site
#Connect-PnPOnline -Url $AppCatalogURL -UseWebLogin 
 
#Add App to App catalog - upload app to sharepoint online app catalog using powershell
#$App = Add-PnPApp -Path $AppFilePath

#$App =Add-PnPApp -Path $AppFilePath -Scope Tenant -Publish

$App =Add-PnPApp -Path $AppFilePath -Publish -SkipFeatureDeployment -Overwrite

'Added ' + $App.Id

Write-Output "(((((((((((((End App catalog Deploy)))))))))))))))))))))))"
Write-Output "   "
