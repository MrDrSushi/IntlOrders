#
#    Graph Batch Delete - ALL items from SharePoint List
#    Uses your existing settings.json
#

clear-host

if (-not (Test-Path -Path ".\settings.json")) {
    Write-Host ">> settings.json not found!" -ForegroundColor Red
    break
}

$settings = Get-Content -Path .\settings.json | ConvertFrom-Json

$startTime = Get-Date

# ====================== TOKEN ======================
$Token_Body = @{
    grant_type    = "client_credentials"
    client_id     = $settings.client_id
    client_secret = $settings.client_secret
    resource      = "https://graph.microsoft.com/"
}

$Token_Params = @{
    URI         = "https://login.microsoftonline.com/$($settings.tenant_domain)/oauth2/token"
    Body        = $Token_Body
    ContentType = "application/x-www-form-urlencoded"
    Method      = "POST"
}

$Token_GraphAPI       = Invoke-RestMethod @Token_Params
$Token_ExpirationTime = (Get-Date).AddSeconds($Token_GraphAPI.expires_in)

# ====================== SITE & LIST ID ======================
$siteReq = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($settings.SPORootSite):/sites/$($settings.SPOSite)" `
                             -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                             -Method GET

$siteID = $siteReq.id.Split(",")[1]

$listReq = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$($settings.SPOList)" `
                             -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                             -Method GET

$listID = $listReq.id

Write-Host "Target List: $($settings.SPOList)  |  SiteID: $siteID  |  ListID: $listID" -ForegroundColor Green

# ====================== DELETION LOOP (PAGINATED) ======================
$deletedTotal = 0
$batchCurrent = 0
$nextLink     = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items?`$top=500&`$select=id"

while ($nextLink) {

    # Token refresh
    if ((Get-Date) -ge $Token_ExpirationTime) {
        Write-Host "`n`t »» Issuing new token ..." -ForegroundColor Yellow
        $Token_GraphAPI       = Invoke-RestMethod @Token_Params
        $Token_ExpirationTime = (Get-Date).AddSeconds($Token_GraphAPI.expires_in)
        Write-Host "`t »» New token issued!" -ForegroundColor Green
    }

    # Get next page of items
    $itemsResp = Invoke-RestMethod -Uri $nextLink `
                                   -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                   -Method GET

    $items = $itemsResp.value
    if ($items.Count -eq 0) { break }

    $nextLink = $itemsResp.'@odata.nextLink'

    # Build batch of up to 20 DELETE requests
    $requests = @()
    $batchIndex = 0

    foreach ($item in $items) {
        $batchIndex++
        $requests += @{
            id      = "$batchIndex"
            url     = "/sites/$siteID/lists/$listID/items/$($item.id)"
            method  = "DELETE"
            headers = @{ "Content-Type" = "application/json" }
        }

        if ($requests.Count -eq 20) {
            $batchCurrent++
            $payload = @{ requests = $requests } | ConvertTo-Json -Depth 4

            $timeRequest = Measure-Command {
                $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                                          -Body $payload `
                                          -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                          -ContentType "application/json" `
                                          -Method POST
            }

            $deletedTotal += 20
            $requests = @()

            # Progress
            Write-Host "════  Batch $($batchCurrent.ToString('000'))  |  Deleted: $deletedTotal items  |  Request time: $($timeRequest.TotalSeconds.ToString('0.000'))s" -ForegroundColor Cyan
        }
    }

    # Send any remaining items (< 20)
    if ($requests.Count -gt 0) {
        $batchCurrent++
        $payload = @{ requests = $requests } | ConvertTo-Json -Depth 4

        $timeRequest = Measure-Command {
            $null = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/`$batch" `
                                      -Body $payload `
                                      -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                      -ContentType "application/json" `
                                      -Method POST
        }

        $deletedTotal += $requests.Count
        Write-Host "════  Batch $($batchCurrent.ToString('000'))  |  Deleted: $deletedTotal items  |  Request time: $($timeRequest.TotalSeconds.ToString('0.000'))s" -ForegroundColor Cyan
    }
}

$endTime   = Get-Date
$totalTime = $endTime - $startTime

Write-Host "`n════════════════»»  DELETION COMPLETE!" -ForegroundColor Green
Write-Host "Total items permanently deleted : $deletedTotal" -ForegroundColor White
Write-Host "Total runtime                   : $($totalTime.ToString('hh\:mm\:ss\.fff'))" -ForegroundColor White