clear-host

if ( (Test-Path -Path ".\settings.json") -eq $false )
{
    ">> settings.json not found!"
    break
}
else
{
    $settings = Get-Content -Path .\settings.json | ConvertFrom-Json
}

$startTime = Get-Date

$Token_Body = @{
                    "tenant"        = $settings.tenant
                    "grant_type"    = "client_credentials"
                    "client_id"     = $settings.client_id
                    "client_secret" = $settings.client_secret
                    "resource"      = "https://graph.microsoft.com/"
               }

$Token_Params = @{
                    "URI"         = "https://login.microsoftonline.com/$($Token_Body.tenant)/oauth2/token"
                    "Body"        = $Token_Body
                    "ContentType" = "application/x-www-form-urlencoded"
                    "Method"      = "POST"
                 }

$Token_GraphAPI       = Invoke-RestMethod @Token_Params
$Token_ExpirationTime = (Get-Date).AddSeconds($Token_GraphAPI.expires_in)

#   Site ID for $settings.SPOSite

$requestSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($settings.SPORootSite):/sites/$($settings.SPOSite)" `
                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                 -ContentType "application/json; charset=utf-8" -Method GET

if ($null -ne $requestSite)
{
    $siteID = $requestSite.id.Split(",")[1]
}
else
{
    ">> Site '$($settings.SPOSite)' not found!"
    break
}

#   the List ID for $settings.SPOList

$requestList = Invoke-RestMethod -Uri  "https://graph.microsoft.com/v1.0/sites/$($siteId)/lists/$($settings.SPOList)" `
                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                 -ContentType "application/json; charset=utf-8" -Method GET

if ($null -ne $requestList)
{
    $listID = $requestList.id
}
else
{
    ">> List '$($settings.SPOList)' not found!"
    break
}

#   deletion process begins from the first ID available

$requestID = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($listID)/items?`$top=1" `
                               -ContentType "application/json"  `
                               -Method Get -Headers @{ Authorization = "Bearer $($Token_GraphAPI.access_token)" }

if ($null -eq $requestID)
{
    "- There are no more items!"
    exit
}

#   where to start and finish (SPO List item IDs)

$itemDeleteStart = $requestID.value.id
$itemDeleteEnd   = 5000

#   self-contained, useful variables for the deletion loop

$batchTotal = [math]::ceiling( ($itemDeleteEnd - $itemDeleteStart) / 20 )
$leadingZeroes = "d" + $($batchTotal).ToString().Length
$batchCurrent = $itemsQueued = $index = $dependsOn = 0
$showOutput = $false
$payload = $requests = $request = @()


#
#   Performs the deletion of items from "itemDeleteStart" to "itemDeleteEnd"
#

$itemDeleteStart..$itemDeleteEnd | % {

    $index++
    $itemsQueued++
    $dependsOn++

    $request = @{
                    id      = $_
                    url     = "/sites/$siteId/lists/$listId/items/$_"
                    headers = @{ "Content-Type" = "application/json" }
                    method  = "DELETE"
                }

    if ($dependsOn -gt 1)
    {
        $request.Add( "dependsOn", @($_-1) )
    }

    $requests += $request


    if ($itemsQueued -eq 20)
    {
        $batchCurrent++
        $showOutput = $true
        $dependsOn = 0

        $payload = @{ requests = $requests } | ConvertTo-Json -Depth 4

        #
        #  Renews the token when it expires
        #

        if ((Get-Date) -ge $Token_ExpirationTime)
        {
            "`n`t »» Issuing new token ... "

            $Token_GraphAPI       = Invoke-RestMethod @Token_Params
            $Token_ExpirationTime = (Get-Date).AddSeconds($Token_GraphAPI.expires_in)

            "`n`t »» New token issued!"
        }

        $batchRequest = $null

        $timeRequest = Measure-Command {
            $batchRequest = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                                              -ContentType "application/json" `
                                              -Body $payload -Method Post `
                                              -Headers @{ Authorization = "Bearer $($Token_GraphAPI.access_token)" }

            if ($batchRequest -EQ $null)
            {
                #
                # TO-DO:  needs some better handling for failed requests - for now just aborting the process
                #
                "`n -- Error - Aborting process"
                break
            }
        }

        $itemsQueued = 0
        $requests = @()
    }


    if ($timeRequest -ne $null)
    {
        $timeTotal += $timeRequest
    }


    if ($showOutput)
    {
        if ($timeRetry -eq $null)
        {
            "════  Batch {0:$($leadingZeroes)} of {1} `t`t`t request: {2:d1}m:{3:d2}s.{4:d3}ms `t`t time: {5:d2}h:{6:d2}m:{7:d2}s.{8:d3}ms  `n" -f $batchCurrent, $batchTotal,   $timeRequest.Minutes, $timeRequest.Seconds, $timeRequest.Milliseconds,   $timeTotal.Hours, $timeTotal.Minutes, $timeTotal.Seconds, $timeTotal.Milliseconds
        }
        else
        {
            $timeTotal.Add($timeRetry)

            "════  Batch {0:$($leadingZeroes)} of {1} `t`t`t request: {2:d1}m:{3:d2}s.{4:d3}ms `t`t`t retry: {5:d1}m:{6:d2}s.{7:d3}ms `t`t time: {8:d2}h:{9:d2}m:{9:d2}s.{10:d3}ms  `n" -f $batchCurrent, $batchTotal,   $timeRequest.Minutes, $timeRequest.Seconds, $timeRequest.Milliseconds,  $timeRetry.Minutes, $timeRetry.Seconds, $timeRetry.Milliseconds,  $timeTotal.Hours, $timeTotal.Minutes, $timeTotal.Seconds, $timeTotal.Milliseconds
        }

        $timeRequest = $null
        $showOutput = $false
    }

}

$endTime = Get-Date
$totalTime = $endTime - $startTime

"`n`n════════════════»»  Total runtime: {0:d2}h:{1:d2}min:{2:d2}s.{3:d3}ms" -f $totalTime.Hours, $totalTime.Minutes, $totalTime.Seconds, $totalTime.Milliseconds