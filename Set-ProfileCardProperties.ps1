<#
 .Synopsis
 Creates the properties in the Microsoft 365 Profile Card for the specified tenant
Based heavily upon https://github.com/merill/m365-gender-pronoun-kit
 #>
 param(
    [ValidateSet('extensionAttribute1', 'extensionAttribute2', 'extensionAttribute3', 'extensionAttribute4', 'extensionAttribute5',
    'extensionAttribute6', 'extensionAttribute7', 'extensionAttribute8', 'extensionAttribute9', 'extensionAttribute10',
    'extensionAttribute11', 'extensionAttribute12', 'extensionAttribute13', 'extensionAttribute14', 'extensionAttribute15')]
    [System.String]$PronounAttribute = "extensionAttribute1",
    [System.String]$WorkingHoursAttribute = "extensionAttribute2",
    [switch] $Step1,
    [switch] $Step2,
    [switch] $Force,
    [switch] $Get
)
$ErrorActionPreference = 'Stop'

$hasGraph = (Get-Module Microsoft.Graph.Authentication -ListAvailable).Length
if ($hasGraph -eq 0) {
    Write-Host "This script requires Microsoft Graph PowerShell, trying to install"
    Find-Module Microsoft.Graph.Authentication
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
}

Connect-MgGraph -Scopes 'User.ReadWrite.All'

$uri = "https://graph.microsoft.com/beta/organization/$((Get-MgContext).TenantId)/settings/profileCardProperties"

$profileCard = Invoke-MgGraphRequest -Uri $uri -Method GET

if(($Step1 -or $Step2) -and !$force -and $profileCard.value.length -gt 0){
    Write-Warning "Existing profile card found."
    Write-Host (ConvertTo-Json $profileCard.value -Depth 5)
    Write-Error 'No changes were made because an existing profile card was found.'
    Write-Error 'Use the -Step1 parameter to overwrite the current profile card.'
    exit
}

if ($Step1) {
  $directoryPropertyName = $PronounAttribute -replace 'extension', 'custom'
  $displayName = "Pronouns"
  $method = "PATCH"
} elseif ($Step2) {
  $directoryPropertyName = $WorkingHoursAttribute -replace 'extension', 'custom'
  $displayName = "Working Hours"
  $method = "POST"
} else {
  $Get = $true
}

if ($Get) {
  Write-Host "Current Card values:"
  Write-Host (ConvertTo-Json $profileCard.value -Depth 5)
  exit
}

$body = @{
    directoryPropertyName = $directoryPropertyName
    annotations = @(
        @{
            displayName = $displayName
        }
    )
}

$bodyJson = ConvertTo-Json $body -Depth 3
Invoke-MgGraphRequest -Uri $uri -Method $method -Body $bodyJson
Write-Host 'Pronoun profile card has been set succesfully.'
