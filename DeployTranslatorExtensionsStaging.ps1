cls
$AppCatalogURL = "https://dwpstage.sharepoint.com/sites/appcatalog"
$AppFilePath_1 = "C:\Users\adminpen.arpula\spfxclientsideprojects\spfx-ecb-extension\sharepoint\solution\spfx-ecb-extension.sppkg"
$AppFilePath_2 = "C:\Users\adminpen.arpula\spfxclientsideprojects\react-application-machine-translations\sharepoint\solution\machine-translation-extension.sppkg"


Connect-PnPOnline -Url $AppCatalogURL -UseWebLogin 

try
{

$AppName_1 = "spfx-ecb-extension-client-side-solution"
$App1 = Get-PnPApp -Scope Tenant | Where {$_.Title -eq $AppName_1}

if (-not ($App1.Id -eq $null))
{
'Removing '  + $App1.Id
Remove-PnPApp -Identity $App1.Id
'Removed '  
}

}
catch
{
    Write-Output $_
}

Write-Output "   "

try
{

$AppName_2 = "machine-translation-extension-client-side-solution"
$App2 = Get-PnPApp -Scope Tenant | Where {$_.Title -eq $AppName_2}

if (-not ($App2.Id -eq $null))
{
'Removing '  + $App2.Id
Remove-PnPApp -Identity $App2.Id
'Removed '  
}

}
catch
{
    Write-Output $_
}

Write-Output "   "


try
{

Write-Output "Starting App catalog Deploy"

$AppOne =Add-PnPApp -Path $AppFilePath_1 -Publish -SkipFeatureDeployment -Overwrite

'Added ' + $AppOne.Id

$AppTwo =Add-PnPApp -Path $AppFilePath_2 -Publish -SkipFeatureDeployment -Overwrite

'Added ' + $AppTwo.Id

Write-Output "End App catalog Deploy"
Write-Output "   "

}
catch
{
    Write-Output $_
}
