clear-host

$RootSite = "tenant-name-goes-here.sharepoint.com"
$Site     = "HR"

$Token_Body = @{
                    "tenant"        = "--- your tenant guid goes here ---"
                    "grant_type"    = "client_credentials"
                    "client_id"     = "--- your azure application id goes here ---"
                    "client_secret" = "--- your azure application secret goes here ---"
                    "resource"      = "https://graph.microsoft.com/"
               }

$Token_Params = @{
                    "URI"         = "https://login.microsoftonline.com/$($Token_Body.tenant)/oauth2/token"
                    "Body"        = $Token_Body
                    "ContentType" = "application/x-www-form-urlencoded"
                    "Method"      = "POST"
                 }

$Token_GraphAPI = Invoke-RestMethod @Token_Params


#   Site ID for $site

$requestSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($RootSite):/sites/$($Site)" `
                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                 -ContentType "application/json; charset=utf-8" -Method GET

$siteID = $requestSite.id.Split(",")[1]



#   User Information List ID

$requestUsersList = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists?`$filter=DisplayName eq 'User Information List'"  `
                                      -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}  `
                                      -ContentType "application/json; charset=utf-8" -Method GET

$usersListID = $requestUsersList.value.id


#   Getting the list of users from the SharePoint User Information List

$requestUsers = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($usersListID)/items?`$expand=fields(`$select=id,IsSiteAdmin,Deleted,SipAddress)"  `
                                  -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}  `
                                  -ContentType "application/json; charset=utf-8" -Method GET


#   The final product is just a list of IDs, we won't need anything else when creating new records for a field person

$users = $requestUsers.value.fields | ? { $_.IsSiteAdmin -eq $false -and $_.Deleted -eq $false -and $_.SipAddress  -ne $null } | select id