#
#     General Settings for the script execution
#

clear-host

$RunSettings_SoftRun      = $false
$RunSettings_TotalRecords = 5

# future implementation
#
# $RunSettings_LocalRegion = New-Object System.Globalization.CultureInfo($settings.LocalRegion)
# $RunSettings_LocalRegion.NumberFormat.NumberDecimalSeparator = $settings.DecimalSeparator
# $RunSettings_LocalRegion.NumberFormat.NumberGroupSeparator   = $settings.GroupSeparator

if ( Test-Path -Path ".\settings.csv" )
{
    ">> settings.json not found!"
    break
}
else
{
    $settings = Get-Content -Path .\settings.json | ConvertFrom-Json
}


#region ══════════════════════════════════════════════════════════════════════════════════════[ CSV File Loading ]

if ( (Test-Path -Path ".\world-data-Airports.csv") -and ($null -eq $airports) )
{
    $airports = Import-Csv ".\world-data-Airports.csv" -Encoding UTF8 | ? { $_.AirportType -in @("Small Airport", "Medium Airport", "Large Airport") }
    #$airports_countries = $airports | Select-Object country , iso2 | Sort-Object -Unique country
}
elseif ((Test-Path -Path ".\world-data-Airports.csv" ) -eq $false)
{
    ">> World-data-Airports.csv not found!"
    break
}

if ( (Test-Path -Path ".\World-data-Locations.csv") -and ($null -eq $locations) )
{
    $locations = Import-Csv ".\World-data-Locations.csv" -Encoding UTF8
    $locations_countries = $locations | Select-Object country , iso2 | Sort-Object -Unique country
}
elseif ((Test-Path -Path ".\World-data-Locations.csv") -eq $false)
{
    ">> Wolrd-data-Locations.csv not found!"
    break
}

if ( (Test-Path -Path ".\world-data-Ports.csv") -and ($null -eq $ports) )
{
    $ports = Import-Csv ".\world-data-Ports.csv" -Encoding UTF8
    #$ports_countries = $ports | Select-Object country | Sort-Object -Unique country
}
elseif ( (Test-Path -Path ".\world-data-Ports.csv") -eq $false )
{
    ">> Wolrd-data-Ports.csv not found!"
    break
}

#endregion ═══════════════════════════════════════════════════════════════════════════════════════════════════════

#region ══════════════════════════════════════════════════════════════════════════════════════[ Random Comments ]

function Get-ShipmentComment
{
    $Notes1 = @(
                "The", "All", "The following"
               )

    $Notes2 = @(
                "goods", "items", "contents", "stock", "supplied", "stated", "products", "merchandise", "wares", "articles"
               )

    $Notes3 = @(
                "shall be", "shall not be", "will be", "will not be", "are", "are not", "were", "were not", "should be", "should not be"
               )

    $Notes4 = @(
                "listed", "weighted", "sealed", "packaged", "refrigerated", "unrefrigerated", "described", "identified", "screened", "inspected", "authorized",
                "cleared", "supervised", "monitored", "tracked", "documented", "perishable"
               )

    $Notes5 = @(
                "on loading", "on unloading", "on arrival", "on departure", "on inspection", "after inspection", "before inspection", "during inspection",
                "uppon customs arrival", "before customs arrival", "inform customs", "GPS monitored", "freight is controlled",
                "before arrival", "after arrival", "check the instructions", "contact the head office", "observe the preservation"
               )

    $Notes6 = @(
                "exercise attention", "do not open", "keep away from direct sunglight", "keep refrigerated", "storage and stability to be followed",
                "handle with care", "avoid stacking", "maintain under prescribed temperature", "do not stack", "keep refrigerated at all times",
                "inspection will be guided", "to hold for inspection", "to remain under supervision", "contact customs", "do not contact customs", " "
               )

    $Notes7 = @(
                "package count by SKU",
                "for collection or prepaid",
                "total weight, cube, carton, and pallet count"
                "transport is monitored",
                "transport is supervised",
                "proceed directly to customs",
                "do not leave customs",
                "free from customs",
                "stay under customs custody",
                "should not remain under customs custody",
                "mark freight terms",
                "certify and states the information",
                "may be imposed for marking",
                "tracking subject to further reschedule",
                "unloading may be followed by inspection",
                "dot not leave customs",
                "customs cleared",
                "customs will determine inspection"
                "fragile shipment upon arrival",
                "follows government authorization",
                "additional information attached",
                "assess the risks and eliminate or minimise them",
                "use and store safely",
                "dangerous",
                "radioactive",
                "samples",
                " "
               )


    $Notes8 = @(
                "perishable", "sensitive", "flamable", "dangerous", "radioactive", "electrical", "military grade",
                "concealed", "self-contained", "hazardous", "enclosed", "unsafe", "frozen", "pharmaceuticals", " "
               )


    $comments      = $null
    $comments_node = $null

    switch ( (Get-Random -Minimum 1 -Maximum 9) )
    {
        #
        # sentences by numbers of words: 8, 7, 6, 5, and 4
        #

        {$_ -eq 8}
        {
            $comments = (Get-Random -InputObject $Notes1) + " " + `
                        (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes3) + " " + `
                        (Get-Random -InputObject $Notes4) + " - " + `
                        (Get-Random -InputObject $Notes5)


            $comments_node = (Get-Random -InputObject $Notes6)

            if ($comments_node -ne " ")
            {
                $comments += ", $comments_node"
            }

            $comments_node = (Get-Random -InputObject $Notes7)

            if ($comments_node -ne " ")
            {
                $comments += " - $comments_node"
            }

            $comments_node = (Get-Random -InputObject $Notes8)

            if ($comments_node -ne " ")
            {
                $comments += " - $comments_node"
            }
        }

        {$_ -eq 7}
        {
            $comments = (Get-Random -InputObject $Notes1) + " " + `
                        (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes3) + " " + `
                        (Get-Random -InputObject $Notes4) + " - " + `
                        (Get-Random -InputObject $Notes5)

            $comments_node = (Get-Random -InputObject $Notes6)

            if ($comments_node -ne " ")
            {
                $comments += ", $comments_node"
            }

            $comments_node = (Get-Random -InputObject $Notes7)

            if ($comments_node -ne " ")
            {
                $comments += " - $comments_node"
            }
        }

        {$_ -eq 6}
        {
            $comments = (Get-Random -InputObject $Notes1) + " " + `
                        (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes3) + " " + `
                        (Get-Random -InputObject $Notes4) + " - " + `
                        (Get-Random -InputObject $Notes5)

            $comments_node = (Get-Random -InputObject $Notes6)

            if ($comments_node -ne " ")
            {
                $comments += ", $comments_node"
            }
        }

        {$_ -eq 5}
        {
            $comments = (Get-Random -InputObject $Notes1) + " " + `
                        (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes3) + " " + `
                        (Get-Random -InputObject $Notes4) + " - " + `
                        (Get-Random -InputObject $Notes5)
        }

        {$_ -eq 4}
        {
            $comments = (Get-Random -InputObject $Notes1) + " " + `
                        (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes3) + " " + `
                        (Get-Random -InputObject $Notes4)
        }

        #
        # complex creations - random counts
        #

        {$_ -eq 3}
        {
            $comments = (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes4)
        }

        {$_ -eq 2}
        {
            $comments = (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes4)

            $comments_node = (Get-Random -InputObject $Notes8)

            if ($comments_node -ne " ")
            {
                $comments += " - $comments_node"
            }

        }

        {$_ -eq 1}
        {
            $comments = (Get-Random -InputObject $Notes2) + " " + `
                        (Get-Random -InputObject $Notes4)

            $comments_node = (Get-Random -InputObject $Notes8)

            if ($comments_node -ne " ")
            {
                $comments += " - $comments_node"
            }

        }
    }

    return $comments
}

#endregion ══════════════════════════════════════════════════════════════════════════════════════════════════════

#region ══════════════════════════════════════════════════════════════════════════════════════[ Additional Types ]

$DateMin = Get-Date -year 2000 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
$DateMax = Get-Date

$ItemType = @(
               "Air Fryer",
               "Airplane",
               "Airsoft Guns",
               "Alexandrite",
               "Aluminium",
               "Apples",
               "Aquamarine",
               "Armored Car",
               "Artichokes",
               "Avocado",
               "Axes",
               "Bacon",
               "Baggels",
               "Bags",
               "Bananas",
               "Batteries",
               "Beans",
               "Beers",
               "Blackberries",
               "Blenders",
               "Blowdryer",
               "Blueberries",
               "Boots",
               "Bread",
               "Brocoli",
               "Bronze",
               "Cabbages",
               "Carrots",
               "Cassette Tapes",
               "Chainsaws",
               "Charcoal",
               "Cherries",
               "Chicken",
               "Chickens",
               "Chocolate",
               "Cigarretes",
               "Cigars",
               "Coal",
               "Cocoa",
               "Coconut",
               "Coffee Beans",
               "Computer",
               "Corn",
               "Cosmetics",
               "Cows",
               "Dates",
               "Diamonds",
               "Dishwashers",
               "Dog Bowls",
               "Dolls",
               "Drones",
               "Dryers",
               "DVD Discs",
               "Dynamite",
               "Eggs",
               "Emeralds",
               "Erasers",
               "Flip flops",
               "Floppy Disks",
               "Forklifts",
               "Forks",
               "Freezers",
               "Fruits",
               "Gold",
               "Grapefruit",
               "Grapes",
               "Grenades",
               "Guava",
               "Hats",
               "High Heels",
               "Horses",
               "Ice Cream",
               "Insecticides",
               "Iron",
               "Jade",
               "Kiwi",
               "Knives",
               "Laptops",
               "Lemons",
               "Limes",
               "Lobsters",
               "Lychees",
               "Machine Guns",
               "Mangoes",
               "Meat",
               "Medical Equipment",
               "Medicine",
               "Microwaves",
               "Mint",
               "Missiles",
               "Mobile Phones",
               "Moccasins",
               "Motorcycles",
               "Mouse Traps",
               "Night Vision Goggles",
               "Notebooks",
               "Oatmeal",
               "Olive Oil",
               "Olives",
               "Onions",
               "Oolong",
               "Orange Juice",
               "Oranges",
               "Oregano",
               "Oysters",
               "Peaches",
               "Peanuts",
               "Pencils",
               "Pens",
               "Peppers",
               "Perfumes",
               "Pinneapples",
               "Pistols",
               "Pork",
               "Potatoes",
               "Prunes",
               "Quahog",
               "Quail",
               "Quail Eggs",
               "Quandong",
               "Quark",
               "Quartz",
               "Quesadilla",
               "Queso Dip",
               "Quiche",
               "Quinoa",
               "Radios",
               "Raisins",
               "Raspberries",
               "Recycling Material",
               "Refrigerators",
               "Revolvers",
               "Rice",
               "Rifles",
               "Roses",
               "Rubber",
               "Ruby",
               "Rulers",
               "Salmon",
               "Salt",
               "Salted Nuts",
               "Sandals",
               "Sapphires",
               "Sardines",
               "Sheeps",
               "Shoes",
               "Shotguns",
               "Silver",
               "Snickers",
               "Solar Panels",
               "Spoons",
               "Steak",
               "Steel",
               "Strawberries",
               "Tablets",
               "Tanks",
               "Tea Leaves",
               "Teddy Bears",
               "Textitles",
               "Tomatoes",
               "Topaz",
               "Tourmaline",
               "T-Shirts",
               "Tumblers",
               "Tuna",
               "Turquoise",
               "TVs",
               "Ube",
               "Udon",
               "Umbrellas",
               "Unagi",
               "Unsalted Nuts",
               "Vanilla",
               "Vases",
               "Veal",
               "Vegetable Oil",
               "Vegetable Soup",
               "Vegetables",
               "Velvet Beans",
               "Venison",
               "VHS Tapes",
               "Video-games",
               "Vienna Sausages",
               "Vinegar",
               "Vinyl Records",
               "Vodka",
               "Washing Machine",
               "Watches",
               "Water Bottles",
               "Watermelon",
               "Wine",
               "Wood",
               "Xanthan Gum",
               "X-Ray Machine",
               "Xylitol",
               "Yams",
               "Yeast",
               "Yellowfin Tuna",
               "Yogurt",
               "Zircom"
             )

$SalesChannel = @(
                     "Internet",
                     "Phone",
                     "Sales Rep"
                 )

$OrderPriority = @(
                    "High",
                    "Medium",
                    "Low"
                  )

$ShippingMethod = @(
                    "Air",
                    "Sea",
                    "Land"
                   )

$Sector = @(
            "NGO",
            "Public",
            "Private"
           )

$AirlineNames = @(
                    "Air France Cargo",
                    "Alaska Air Cargo",
                    "American Airlines Cargo",
                    "American Airlines Freight",
                    "Asiana Airlines Cargo",
                    "Atlas Air",
                    "British Airways Cargo",
                    "British Airways World Cargo",
                    "Cargo Garuda Indonesia",
                    "Cargolux",
                    "Caribbean Airlines",
                    "Cathay Pacific Cargo",
                    "China Airlines",
                    "China Southern Airlines Cargo",
                    "Czech Airlines Cargo",
                    "Delta Airlines Cargo",
                    "DHL Aviation",
                    "Dragon Air Cargo",
                    "Emirates SkyCargo",
                    "Etihad Airways Cargo",
                    "EVA Air Cargo",
                    "FedEx Express",
                    "Gol Transportes Aï¿½reos",
                    "Gulf Air Cargo",
                    "Hainan Airlines Cargo",
                    "Iberia Cargo",
                    "Japan Airlines Cargo",
                    "Kenya Airways Cargo",
                    "KLM Cargo",
                    "Korean Air Cargo",
                    "Kuwait Airways Cargo",
                    "LOT Polish Airlines Cargo",
                    "Lufthansa Cargo",
                    "Pakistan Intl Airlines Cargo",
                    "Philippine Airlines Cargo",
                    "Polar Air Cargo",
                    "Qantas Freight",
                    "Qatar Airways Cargo",
                    "SAS Cargo Group",
                    "Shenzhen Airlines Cargo",
                    "Sichuan Airlines Cargo",
                    "South African Airways",
                    "SriLankan Cargo",
                    "Sudan Airways",
                    "Swiss WorldCargo",
                    "Thai Airways Cargo",
                    "Turkish Airlines",
                    "Turkish Cargo",
                    "United Airlines Cargo",
                    "UPS Airlines",
                    "Virgin Atlantic Cargo",
                    "Virgin Australia Cargo",
                    "WestJet Cargo"
            )

$VesselName = @(
                "Antwerpen Express",
                "Basle Express",
                "Budapest Express",
                "Cosco Belgium",
                "Cosco Houston",
                "Cosco Japan",
                "Cosco Oceania",
                "Cosco Pacific",
                "Cosco Taicang",
                "Cscl Bohai Sea",
                "Cscl Jupiter",
                "Cscl Mars",
                "Cscl Mercury",
                "Cscl Nepture",
                "Cscl Saturn",
                "Cscl Star",
                "Cscl Uranus",
                "Cscl Venus",
                "Cyprus Cape Martin",
                "Ebba Maersk",
                "Edith Maersk",
                "Eleonora Maersk",
                "Elly Maersk",
                "Emma Maersk",
                "Essen Express",
                "Estelle Maersk",
                "Eugen Maersk",
                "Evelyn Maersk",
                "France CMA CGM Fidelio",
                "Germany CMA CGM Orfeo",
                "Hamburg Express",
                "Hong Kong Express",
                "Leverkusen Express",
                "Liberia Aegiali",
                "Liberia As Rafaela",
                "Liberia Bomar Rossi",
                "Liberia Cala Paguro",
                "Liberia E R Felixstowe",
                "Liberia E R France",
                "Liberia Emirates Dana",
                "Liberia Emirates Wafa",
                "Liberia Emirates Wasl",
                "Liberia Gsl Africa",
                "Liberia Gsl Valerie",
                "Liberia Hansa Breitenburg",
                "Liberia Ikaria",
                "Ludwigshafen Express",
                "Madrid Express",
                "Malta A Idefix",
                "Malta Adrian Schulte",
                "Malta Cma Cgm Coral",
                "Marshall Islands Baltic Bridge",
                "Marshall Islands Baltic West",
                "Marshall Islands Cape Fawley",
                "Msc Filomena",
                "Nagoya Express",
                "New York Express",
                "Panama Akinada Bridge",
                "Panama Cosco Africa",
                "Panama Hakata Seoul",
                "Paris Express",
                "Portugal Actuaria",
                "Portugal Bernadette",
                "Portugal Conti Courage",
                "Shanghai Express",
                "Singapore Apl Columbus",
                "Singapore Apl Jeddah",
                "Singapore Asiatic King",
                "Singapore Asiatic Moon",
                "Singapore Asiatic Neptune",
                "Singapore Ever United",
                "Singapore Green Earth",
                "Singapore Green Pole",
                "Singapore Green Sea",
                "Singapore Interasia Heritage",
                "Singapore Jitra Bhum",
                "South Korea Hyundai Goodwill",
                "Southampton Express",
                "Thailand Jaru Bhum",
                "Ulsan Express",
                "Vienna Express"
               )

$FreightTerms = @(
                   "Prepaid",
                   "Collect",
                   "Elsewhere"
                 )

#endregion ═══════════════════════════════════════════════════════════════════════════════════════════════════════

#region ══════════════════════════════════════════════════════════════════════════════════════[ Graph Token, Site, List, and Users ]

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

#   Site ID for $settings.SPORootSite

$requestSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($settings.SPORootSite):/sites/$($settings.SPOSite)" `
                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"} `
                                 -ContentType "application/json; charset=utf-8" -Method GET

if ($null -ne $requestSite)
{
    $siteID = $requestSite.id.Split(",")[1]
}
else
{
    ">> Site '$($settings.SPORootSite)' not found!"
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

#
#   IMPORTANT NOTICE
#   ================
#
#   SharePoint Online keeps a list of users under a list called "User Information List" and is stored under the following URL:
#
#           https://TENANT.sharepoint.com/sites/SITE/_catalogs/users/
#
#   I'm using the 'Display Name' for filtering the list, and this value will vary according to the language in use by your SharePoint,
#   make sure to check the url above on your tenant to get the description matching the value from your SharePoint in the code below.
#
#   Currently there is no way to obtain the list by its physical name by doing something like: "$filter=Name eq 'users'" (you'll get HTTP 400 Bad Request)
#


$requestUsersList = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists?`$filter=DisplayName eq 'User Information List'"  `
                                      -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}  `
                                      -ContentType "application/json; charset=utf-8" -Method GET


if ($requestUsersList -ne $null)
{
    $usersListID = $requestUsersList.value.id
}
else
{
    ">> SharePoint Online 'User Information List' not found!"
    break
}

#   Getting the list of users from the SharePoint User Information List for the given site

$requestUsers = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($usersListID)/items?`$expand=fields(`$select=id,IsSiteAdmin,Deleted,SipAddress)"  `
                                  -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}  `
                                  -ContentType "application/json; charset=utf-8" -Method GET

#   The filter below removes the following:
#
#   IsSiteAdmin:   SPO site admins (my preference but you can include them if you want)
#   Deleted:       brings back only active users with an account and permissions granted to the site
#   SipAddress:    the most important filter, it will excludes the system accounts such as "SHAREPOINT\system", "NT Service\SPSearch", "Everyone", and other groups
#                  it is redudant to include "Fields/ContentType -eq "Person", filtering the data with SipAddress will limit the results only to user accounts
#
#   The final product is just a list of IDs, we won't need anything else when creating new records

$Users = $requestUsers.value.fields | ? { $_.IsSiteAdmin -eq $false -and $_.Deleted -eq $false -and $_.SipAddress  -ne $null } | select id

#endregion ═════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════


#
#   Field Preparation (randomized data)
#

$requestExecuted = $false
$leadingZeroes   = "d" + $([math]::ceiling($RunSettings_TotalRecords / 20).ToString().Length)
$dependsOn       = 0
$batch_Counter   = 0
$load            = @{ requests = @() }
$payload         = $null

$timeRecord  = 0
$timePayload = 0
$timeRequest = 0
$timeTotal   = 0

for ($loop = 1; $loop -le $RunSettings_TotalRecords; $loop++)
{
    #
    #   Current Progress
    #

    $timeRecord = Measure-Command {

        $date = New-Object DateTime (Get-Random -min $DateMin.ticks -max $DateMax.ticks)

        $field_ItemType     = Get-Random -InputObject $ItemType
        $field_ItemSKU      = (New-Guid).ToString()
        $field_Sector       = Get-Random -InputObject $Sector
        $field_Confidential = Get-Random -InputObject @($false, $true)

        $field_OrderID       = Get-Random -Minimum 100000000 -Maximum 999999999
        $field_OrderPriority = Get-Random -InputObject $OrderPriority
        $field_OrderDate     = "{0:yyyy-MM-ddTHH:mm:ss.mmm}" -f $date

        $field_UnitsSold = Get-Random -Minimum 1 -Maximum 10000
        $field_UnitPrice = [math]::round( (Get-Random -Minimum 1.00 -Maximum 10000.00) , 2 )
        $field_UnitCost  = [math]::round( $field_UnitPrice * (Get-Random -Minimum 10 -Maximum 100) / 100 , 2 )

        $field_TotalRevenue = [math]::round( $field_UnitsSold * $field_UnitPrice , 2 )
        $field_TotalCost    = [math]::round( $field_UnitsSold * $field_UnitCost , 2 )
        $field_TotalProfit  = [math]::round( $field_TotalRevenue - $field_TotalCost , 2 )

        $field_Containers   = Get-Random -Minimum 0 -Maximum ( [int]([math]::Truncate( $field_UnitsSold / 1000 * 10 ) + 1 ) )
        $field_FreightTerms = Get-Random -InputObject $FreightTerms
        $field_SalesChannel = Get-Random -InputObject $SalesChannel

        $field_SalesCoordinator   = (Get-Random -InputObject $Users).Id
        $field_SalesPerson        = (Get-Random -InputObject $Users).Id
        $field_PaymentCoordinator = (Get-Random -InputObject $Users).Id
        $field_ShippingForeman    = (Get-Random -InputObject $Users).Id

        $field_ShippingInsured   = Get-Random -InputObject @($false, $true)
        $field_ShippingDate      = "{0:yyyy-MM-ddTHH:mm:ss.mmm}" -f $date.AddSeconds( (Get-Random -Minimum 1000000 -Maximum 5300000) )

        $field_ShippingMethod    = Get-Random -InputObject $ShippingMethod

        $field_VesselNameOrID    = ""
        $field_PortOfOrigin      = ""
        $field_PortOfOriginName  = ""
        $field_PortOfDestiny     = ""
        $field_PortOfDestinyName = ""

        switch ($field_ShippingMethod)
        {
            "Air"
            {
                $origin  = Get-Random -InputObject $airports
                $destiny = Get-Random -InputObject $airports

                $field_VesselNameOrID = Get-Random -InputObject $AirlineNames

                $field_PortOfOrigin      = $origin.Municipality  + ", " + $origin.Country
                $field_PortOfOriginName  = $origin.AirportName

                $field_PortOfDestiny     = $destiny.Municipality + ", " + $destiny.Country
                $field_PortOfDestinyName = $destiny.AirportName
            }

            "Land"
            {
                $country = (Get-Random -InputObject $locations_countries).Country

                $origin  = (Get-Random -InputObject ($locations | ? { $_.Country -eq $country }))
                $destiny = (Get-Random -InputObject ($locations | ? { $_.Country -eq $country }))

                $field_PortOfOrigin   = $origin.City  + ", " + $origin.Country
                $field_PortOfDestiny  = $destiny.City + ", " + $destiny.Country
            }

            "Sea"
            {
                $field_VesselNameOrID = Get-Random -InputObject $VesselName

                $origin  = (Get-Random -InputObject $ports)
                $destiny = (Get-Random -InputObject $ports)

                $field_PortOfOrigin      = $origin.Country
                $field_PortOfOriginName  = $origin.PortName

                $field_PortOfDestiny     = $destiny.Country
                $field_PortOfDestinyName = $destiny.PortName
            }
        }

        $field_ShippingNotes = Get-ShipmentComment
        $field_Comments      = Get-ShipmentComment
        $field_Title         = "{0}, {1}" -f  $field_ItemType, $field_UnitsSold.ToString("###,###,###,###")    #$field_UnitsSold.ToString("N0", $RegionNZ)

    }

    #
    #   adds the current record to the body of requests
    #

    $load.requests += @{
                             id      = $loop
                             url     = "sites/$($siteID)/lists/$($listID)/items"
                             method  = "POST"
                             headers = @{ "content-type" = "application/json" }
                             body    = @{ fields = @{
                                                        "ItemType"      = $field_ItemType
                                                        "ItemSKU"       = $field_ItemSKU
                                                        "Sector"        = $field_Sector
                                                        "Confidential"  = $field_Confidential

                                                        "OrderID"       = $field_OrderID
                                                        "OrderPriority" = $field_OrderPriority
                                                        "OrderDate"     = $field_OrderDate

                                                        "UnitsSold"     = $field_UnitsSold
                                                        "UnitPrice"     = $field_UnitPrice
                                                        "UnitCost"      = $field_UnitCost

                                                        "TotalRevenue"  = $field_TotalRevenue
                                                        "TotalCost"     = $field_TotalCost
                                                        "TotalProfit"   = $field_TotalProfit

                                                        "Containers"    = $field_Containers
                                                        "FreightTerms"  = $field_FreightTerms
                                                        "SalesChannel"  = $field_SalesChannel

                                                        "SalesCoordinatorLookupId"   = $field_SalesCoordinator
                                                        "SalesPersonLookupId"        = $field_SalesPerson
                                                        "PaymentCoordinatorLookupId" = $field_PaymentCoordinator
                                                        "ShippingForemanLookupId"    = $field_ShippingForeman

                                                        "ShippingInsured"   = $field_ShippingInsured
                                                        "ShippingDate"      = $field_ShippingDate
                                                        "ShippingMethod"    = $field_ShippingMethod

                                                        "VesselNameOrID"    = $field_VesselNameOrID
                                                        "PortOfOrigin"      = $field_PortOfOrigin
                                                        "PortOfOriginName"  = $field_PortOfOriginName
                                                        "PortOfDestiny"     = $field_PortOfDestiny
                                                        "PortOfDestinyName" = $field_PortOfDestinyName

                                                        "ShippingNotes"     = $field_ShippingNotes
                                                        "Comments"          = $field_Comments
                                                        "Title"             = $field_Title
                                                    }
                                        }
                       }

    #
    #   increments the dependsOn indexer, when it is greather than 1, dependency will be created
    #

    $dependsOn++

    #
    #   creates depencies for requests with more than 1 item
    #

    if ($dependsOn -gt 1)
    {
        $load.requests[ $load.requests.Count-1 ].Add( "dependsOn", @( ($loop-1).ToString() ) )
    }

    #
    #   Submits the request to the Graph API endpoint
    #   The request is limited to a maximum of 20 records
    #

    if ( ($load.requests.Count -eq 20) -or ($load.requests.count -gt 0 -and $loop -eq $RunSettings_TotalRecords) )
    {
        $batch_Counter++
        $dependsOn = 0

        $payload = ConvertTo-Json $load -Depth 4

        try
        {
            #
            #   Renews the token when its time has expired
            #

            if ( ((Get-Date) -ge $Token_ExpirationTime) -and (!$RunSettings_SoftRun))
            {
                "`n`t >> Issuing new token ... "

                $Token_GraphAPI       = Invoke-RestMethod @Token_Params
                $Token_ExpirationTime = (Get-Date).AddSeconds($Token_GraphAPI.expires_in)

                "`n`t >> New token issued!"
            }

            #
            #   Send the batch request out to the endpoint
            #

            $timeRequest = Measure-Command -Expression {
                if (!$RunSettings_SoftRun)
                {
                    $request = $null
                    $request = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/$batch' `
                                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}  `
                                                 -ContentType "application/json; charset=utf-8"  `
                                                 -Body $payload  `
                                                 -Method Post

                    if ($request -eq $null)
                    {
                        # TO-DO:  failed requests are not being correctly handled here (to be improved later on)

                        "`n -- Error - Aborting process"
                        break
                    }
                }
            }

            $requestExecuted = $true
        }
        catch
        {
                #
                #   Something went wrong, we will display the batch output followed by its Exception Message (to be improved later on)
                #

                "`n>>>>  Batch {0:$($leadingZeroes)} of {1} `t`t`t :: Exception: {2} `n" -f $batch_Counter, [math]::ceiling($RunSettings_TotalRecords / 20), $_.Exception.Message
            }

        $load = @{ requests = @() }
    }

    #
    #   Computing the total time:
    #
    #     - payload time = time spent during the data loop to create the payload
    #

    $timePayload += $timeRecord    # formula: the sum of each individual record to assemble the limit of 20 (or inferior when it is the end)

    #
    #   Display the Output
    #
    #      - batch:    current batch number and total of batches to complete
    #
    #      - payload:  total spent to assemble the payload (looping)
    #
    #      - request:  time spent for the Graph API request
    #
    #      - time:     total time (current time spent)
    #
    #      - records:  the current number of processed and total records to be completed
    #

    if ($requestExecuted)
    {
        $timeTotal += $timePayload + $timeRequest

        "════  Batch {0:$($leadingZeroes)} of {1} `t`t`t payload: {2:d2}s.{3:d3}ms `t`t request: {4:d1}m:{5:d2}s.{6:d3}ms `t`t time: {7:d2}h:{8:d2}m:{9:d2}s.{10:d3}ms  `t`t {11,10} / {12} records `n" -f $batch_Counter, [math]::ceiling($RunSettings_TotalRecords / 20),  $timePayload.Seconds, $timePayload.Milliseconds,   $timeRequest.Minutes, $timeRequest.Seconds, $timeRequest.Milliseconds,   $timeTotal.Hours, $timeTotal.Minutes, $timeTotal.Seconds, $timeTotal.Milliseconds, $loop, $RunSettings_TotalRecords

        #
        #   shows anything wrong (response is not HTTP 201)
        #

        if (!$RunSettings_SoftRun)
        {
            forEach($response in $request.responses)
            {
                if ($response.status -ne 201)
                {
                    " `t ID: {0}"                -f $response.id
                    " `t HTTP {0} - {1}, {2} `n" -f $response.status, $response.body.error.code, $response.body.error.message
                }
            }
            "`n"
        }

        $timePayload = 0
        $requestExecuted = $false
    }

}