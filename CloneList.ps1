#
#   This script has serious isues with XML replacing string portion - DO NOT USE IN PRODUCTION ENVIRONMENT
#

Clear-Host
 
$templateFile = ".\Orders.xml"
$oldListName  = "Sales"
$newListName  = "Orders"

$settings = Get-Content -Path .\settings.json | ConvertFrom-Json

$conn = Connect-PnPOnline -Url                  "https://$($settings.SPORootSite)/sites/$($settings.SPOSite)"                             `
                          -ClientId             "$($settings.client_id)"                                                                  `
                          -CertificatePassword  (ConvertTo-SecureString -String $($settings.certificate_password) -AsPlainText -Force)    `
                          -CertificatePath      ".\$($settings.entra_applicationname).pfx"                                                `
                          -Tenant               $settings.tenant_domain -ReturnConnection

Write-Host "- Extracting list schema from '$oldListName'..."

Get-PnPSiteTemplate -Connection $conn -Handlers Lists -ListsToExtract $oldListName -Out $templateFile

Write-Host "- Renaming references in template from '$oldListName' to '$newListName'..."

#  this portion is really poorly implemented - it just replaces all occurrences of the old list name with the new list name, 
#  which can lead to unintended consequences if the old list name appears in other contexts within the XML. 
#  it needs to safely rename list name only - currently renames columns, views, etc - DOT NOT USE IN PRODUCTION ENVIRONMENT

# $xmlContent = Get-Content -Path $templateFile -Raw
# $xmlContent = $xmlContent -replace $oldListName, $newListName
# Set-Content -Path $templateFile -Value $xmlContent

# Write-Host "Creating new list..."

# Invoke-PnPSiteTemplate -Connection $conn -Path $templateFile

# Write-Host "List cloned successfully!`n"