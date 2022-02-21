<#
 .Synopsis
 Runs a sync to update Azure AD attributes with the corresponding value from the SharePoint/Delve user profile.
 Based heavily upon https://github.com/merill/m365-gender-pronoun-kit
#>

# Input bindings are passed in via param block.
param($Timer)

# The following should be provided by environment variables. This way when you setup the Azure function you can 
# provide these items via App Settings (and the certificate via the keyvault which then can be used as an app setting.)
$Url = $env:Url
if ($null -eq $Url) { throw "Missing Url env var" }
$Tenant = $env:Tenant
if ($null -eq $Tenant) { throw "Missing Tenant env var" }
$AzureClientId = $env:AzureClientId
if ($null -eq $AzureClientId) { throw "Missing AzureClientId env var" }
$ClientId = $env:ClientId
if ($null -eq $ClientId) { throw "Missing ClientId env var" }
$CertificateBase64Encoded = $env:CertificateBase64Encoded
if ($null -eq $CertificateBase64Encoded) { throw "Missing CertificateBase64Encoded env var" }
$ClientSecret = $env:ClientSecret
if ($null -eq $ClientSecret) { throw "Missing ClientSecret env var" }
$INCATTRS = $env:INCATTRS       # should be comma separated string of extensionAttribute1=UserProfileAttrName pairs. e.g. extensionAttribute1=Pronouns,extensionAttribute2=WorkingHours
if ($null -eq $INCATTRS) { throw "Missing INCATTRS env var" }


$tokenConnection = Connect-PnPOnline -Tenant $Tenant -Url $Url -ClientId $AzureClientId -CertificateBase64Encoded $CertificateBase64Encoded -ReturnConnection
$accesstoken = Get-PnPGraphAccessToken -Connection $tokenConnection

function Invoke-Graph{
    param(
        $Uri,
        [ValidateSet('PATCH', 'GET', 'POST')] $Method,
        $Body,
        [ValidateSet('v1.0', 'beta')] $ApiVersion = "v1.0"
    )
    if ($null -eq $accesstoken) {
      $accesstoken = Get-PnPGraphAccessToken -Connection $tokenConnection
    }
    if($Uri.StartsWith('https')){
        $graphUri = $Uri
    }
    else {
        $graphUri = 'https://graph.microsoft.com/{0}/{1}' -f $ApiVersion, $Uri
    }

    $res = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri $graphUri -Method $Method -Body $Body -ContentType 'application/json'
    Write-Host $res
    Write-Output $res
}

# parse args in to list of aad/sharepoint pairs:
# e.g. extensionAttribute1=ACHPronouns,extensionAttribute2=ACHWorkingHours,extensionAttribute3=ACHNativeLand
$VALIDATTR = @('extensionAttribute1', 'extensionAttribute2', 'extensionAttribute3', 'extensionAttribute4', 'extensionAttribute5',
'extensionAttribute6', 'extensionAttribute7', 'extensionAttribute8', 'extensionAttribute9', 'extensionAttribute10',
'extensionAttribute11', 'extensionAttribute12', 'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15')
$attrs = [System.Collections.Generic.List[PSObject]]::new()
foreach ( $param in $INCATTRS.split(",") )
{
    $pair = $param.Split("=")
    if ($pair[0] -and $VALIDATTR.Contains($pair[0])) {

    } else {
      Throw 'Invalid attribute pairing.'
    }
    $attrs.Add($pair)
}
$hasPnP = (Get-Module PnP.PowerShell -ListAvailable).Length
if ($hasPnP -eq 0) {
    Write-Host "This script requires PnP PowerShell. Please include it in requirements.psd1"
    Throw "Missing PnP.Powershell module."
}

$pnpConnection = Connect-PnPOnline -Url $Url -ClientId $ClientId -ClientSecret $ClientSecret -ReturnConnection

$aadUsers = (Invoke-Graph -Uri 'users?$select=id,userPrincipalName,onPremisesExtensionAttributes&$top=999' -Method GET)
do{
    foreach ($aadUser in $aadUsers.value){
        if ($aadUser.UserPrincipalName.Contains("#EXT#@")) { continue }
        Write-Host "Checking $($aadUser.UserPrincipalName)"
        $pnpUser = Get-PnPUserProfileProperty -Account $aadUser.UserPrincipalName -Connection $pnpConnection
        $body = @{onPremisesExtensionAttributes = @{}}
        foreach ($pair in $attrs) {
          $extAttr = $pair[0]
          $pnpAttr = $pair[1]
          $aadValue = $aadUser.onPremisesExtensionAttributes."$extAttr"
          $pnpValue = $pnpUser.UserProfileProperties."$pnpAttr"
          Write-Host "aad: $aadValue pnp: $pnpValue"
          if($pnpValue -eq '') {$pnpValue = $null}
          if($aadValue -ne $pnpValue){
              $body["onPremisesExtensionAttributes"]["$extAttr"] = $pnpValue
          }
        }
        if ($body["onPremisesExtensionAttributes"].count -gt 0) {
          Write-Host "... Updating"
          Invoke-Graph -Uri "users/$($aadUser.id)" -Method PATCH -Body (ConvertTo-Json $body -Depth 3)
        }
    }
    if($null -ne $aadUsers.'@odata.nextLink') { $aadUsers = Invoke-Graph -Uri $aadUsers.'@odata.nextLink' -Method GET }
} while ($null -ne $aadUsers.'@odata.nextLink')


# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()
# Write an information log with the current time.
Write-Host "Timered sync function ran! TIME: $currentUTCtime"
