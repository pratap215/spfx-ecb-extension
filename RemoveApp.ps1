#Parameters
#$AppCatalogSiteURL = "https://crescent.sharepoint.com/sites/Apps"
#AppCatalogSiteURL
#$AppCatalogSiteURL = "https://8p5g5n.sharepoint.com/sites/appcatalog"
$AppCatalogSiteURL = "https://8p5g5n.sharepoint.com/sites/appcatalog"
#$AppFilePath = "C:\Users\adminpen.arpula\spfxclientsideprojects\react-application-machine-translations\sharepoint\solution\machine-translation-extension.sppkg"
$AppName = "spfx-ecb-extension-client-side-solution"
  
#Connect to SharePoint Online App Catalog site
Connect-PnPOnline -Url $AppCatalogSiteURL -UseWebLogin 
 
#Get the from tenant App catalog
$App = Get-PnPApp -Scope Tenant | Where {$_.Title -eq $AppName}

'Removed '  + $App.Id
 
#Remove an App from App Catalog Site
Remove-PnPApp -Identity $App.Id

#Remove-PnPApp -Identity 87CC9434-00C1-4C88-B50C-DD8A7888F9B8


#Read more: https://www.sharepointdiary.com/2019/09/sharepoint-online-how-to-remove-app.html#ixzz7CLXkHspj