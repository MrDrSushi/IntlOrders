clear-host

if ( (Test-Path -Path ".\settings.json") -eq $false )
{
    write-error ">> settings.json not found!`n"
    break
}
else
{
    $settings = Get-Content -Path .\settings.json | ConvertFrom-Json
}

# Register the app + generate cert

$result = Register-PnPEntraIDApp -ApplicationName      $($settings.entra_applicationname)    `
                                 -Tenant               $($settings.tenant_domain)            `
                                 -CertificatePassword  $(ConvertTo-SecureString -String ($settings.certificate_password) -AsPlainText -Force) `
                                 -OutPath              ".\" `
                                 #-SharePointApplicationPermissions "Sites.FullControl.All"
# Output Results

$result | Format-List ClientId, Tenant, CertificateBase64