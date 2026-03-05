clear-host

#
#   General Settings for the script execution
#

$RunSettings_SoftRun      = $false
$RunSettings_TotalRecords = 10

#
#   future implementation
#
#   $RunSettings_LocalRegion = New-Object System.Globalization.CultureInfo($settings.LocalRegion)
#   $RunSettings_LocalRegion.NumberFormat.NumberDecimalSeparator = $settings.DecimalSeparator
#   $RunSettings_LocalRegion.NumberFormat.NumberGroupSeparator   = $settings.GroupSeparator
#

if ( (Test-Path -Path ".\settings.json") -eq $false )
{
    write-error ">> settings.json not found!`n"
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
    write-error ">> World-data-Airports.csv not found!`n"
    break
}

if ( (Test-Path -Path ".\world-data-Locations.csv") -and ($null -eq $locations) )
{
    $locations = Import-Csv ".\world-data-Locations.csv" -Encoding UTF8
    $locations_countries = $locations | Select-Object country , iso2 | Sort-Object -Unique country
}
elseif ((Test-Path -Path ".\world-data-Locations.csv") -eq $false)
{
    write-error ">> world-data-Locations.csv not found!`n"
    break
}

if ( (Test-Path -Path ".\world-data-Ports.csv") -and ($null -eq $ports) )
{
    $ports = Import-Csv ".\world-data-Ports.csv" -Encoding UTF8
    #$ports_countries = $ports | Select-Object country | Sort-Object -Unique country
}
elseif ( (Test-Path -Path ".\world-data-Ports.csv") -eq $false )
{
    write-error ">> wolrd-data-Ports.csv not found!`n"
    break
}

#endregion ═══════════════════════════════════════════════════════════════════════════════════════════════════════

#region ══════════════════════════════════════════════════════════════════════════════════════[ Random Comments ]

function Get-ShipmentComment {
    
    $shippingstuff = Get-Random -InputObject @('Box','Cargo', 'Container', 'Crate', 'Package', 'Pallet', 'Stash')

    $subjects      = @(
        "The cargo", "$($shippingstuff) #$(Get-Random -Minimum 1 -Maximum 9999)", "Experimental samples", "The 'merchandise'", 
        "A suspiciously heavy $($shippingstuff.ToLower())", "The $($shippingstuff.ToLower()) of dreams", "The CEO's 'private' $($shippingstuff.ToLower())", "A $($shippingstuff.ToLower()) emitting low-frequency humming",
        "The prototype", "An unmarked black $($shippingstuff.ToLower())", "A $($shippingstuff.ToLower()) of vintage mannequins", "The 'non-hazardous' (mostly) sludge",
        "A crate labeled 'DO NOT EAT'", "The office's backup supply of snacks", "A collection of antique accordions",
        "The diplomatic pouch", "Box #666", "A frozen block of unknown origin", "The shipment of rubber ducks", "The gravity-defying cube"
    )
    $status        = @(
        "is currently", "is definitely not", "has been legally declared", "appears to be", "is technically", 
        "is stubbornly remaining", "was accidentally classified as", "is vibrating into", "has been temporarily promoted to",
        "is officially haunting", "is masquerading as", "is theoretically", "is slowly becoming", "is refusing to acknowledge",
        "has been disavowed by", "is oscillating between", "is being guarded by", "is strictly prohibited from"
    )
    $actions       = @(
        "vibrated", "soaked in mystery", "blessed by the night shift", "ignored for three days", "stacked precariously", 
        "documented via interpretive dance", "used as a temporary coffee table", "marinated in warehouse humidity",
        "scrubbed with a toothbrush", "shouted at by the manager", "covered in sticky notes", "photographed for evidence",
        "balanced on a single toothpick", "baptized in spilled energy drink", "lost and then found in the rafters",
        "integrated into the warehouse's ecosystem", "subjected to a stern talking-to", "wrapped in excessive bubble wrap"
    )
    $timing        = @(
        "during the solar eclipse", "at the crack of noon", "while the supervisor was 'fishing'", "exactly at the wrong time", 
        "during the Great Coffee Shortage of 2026", "at precisely 3:00 AM", "while the internet was down", 
        "during the office holiday party", "right before the inspector arrived", "at the exact moment of a power surge",
        "during a heated argument about pizza toppings", "in the middle of a heavy thunderstorm", "during a suspiciously quiet moment",
        "while everyone was watching a cat video", "right as the forklift battery died"
    )
    $safety        = @(
        "HANDLE WITH GLOVES", "DO NOT FEED", "KEEP UPRIGHT (mostly)", "FRAGILE LIKE MY EGO", "SMELL AT YOUR OWN RISK", 
        "DO NOT STACK ON THE INTERN", "STAY 50 FEET BACK", "WEAR A HAZMAT SUIT", "AUTHORIZED PERSONNEL ONLY", 
        "DO NOT LOOK DIRECTLY AT IT", "SHAKE WELL BEFORE OPENING", "DO NOT EXPOSE TO OXYGEN", "CONTAINS SHARP TRUTHS",
        "AVOID EYE CONTACT", "SECURE WITH ADHESIVE TAPE AND PRAYER", "HIGHLY REACTIVE TO SARCASM"
    )
    $customs       = @(
        "cleared by a very sleepy officer", "held for 'investigation'", "lost in the paperwork void", "bribed through with high-fives", 
        "flagged for excessive weirdness", "re-routed via a dimension we don't recognize", "denied entry by a confused pigeon",
        "approved by a robot named Gary", "missing 47 necessary stamps", "subjected to a 24-hour interrogation",
        "currently being used as a paperweight in sector 7", "cleared but with a very judgmental look", "lost in the 'miscellaneous' pile",
        "accidentally exported to the moon", "confiscated by the fashion police"
    )
    $warnings      = @(
        "contains ghosts", "highly caffeinated", "slightly glowing", "unpredictable in rain", "honestly, just be careful", 
        "may contain traces of glitter", "vibrates when spoken to", "prone to spontaneous combustion", "leaks concentrated sadness",
        "attracts wild raccoons", "tastes like purple", "may cause mild hallucinations", "has its own gravitational pull",
        "is sentient on Tuesdays", "sounds like a choir of bees", "is suspiciously cold to the touch"
    )
    $complaints    = @(
        "The forklift is haunted.", "I'm not paid enough for this.", "Someone ate my labeled yogurt.", 
        "The warehouse cat has claimed this as a bed.", "It's too early for this.", "The printer is screaming again.",
        "The lights are flickering in Morse code.", "The breakroom microwave smells like burnt hair.", "My boots are squeaking.",
        "The roof is leaking green liquid.", "Someone replaced the water in the cooler with tea.", "The walls are sweating.",
        "I've forgotten what sunlight looks like.", "The vending machine took my last dollar.", "The radio only plays polka."
    )
    $sounds        = @(
        "making a ticking sound", "whistling 'Despacito'", "purring loudly", "emitting a faint static noise", 
        "occasionally shouting in Latin", "humming an 80s power ballad", "clucking like a nervous chicken",
        "making a sound like tearing silk", "whispering secrets about the manager", "beeping in an irregular rhythm",
        "thumping like a heartbeat", "giggling softly when moved", "echoing with the sound of distant waves"
    )
    $smells        = @(
        "smelling faintly of wet dog", "scented like 'New Car' and regret", "smelling like a campfire in a rainstorm", 
        "scented with expensive cologne and sulfur", "smelling like old library books", "smelling of ozone and ozone-adjacent things",
        "scented with fresh cinnamon and fear", "smelling like a damp basement", "scented like a very old sandwich",
        "smelling like 'The Ocean' (but the scary part)", "smelling like burnt toast and rubber"
    )
    $conditions    = @(
        "under a leaking roof", "in a localized gravity anomaly", "surrounded by suspicious pigeons", 
        "while covered in 'Property of Area 51' stickers", "in a puddle of unknown blue liquid", "resting on a bed of hay",
        "inside a giant Ziploc bag", "while being used as a doorstop", "hidden behind a stack of empty pallets",
        "in the path of a very aggressive Roomba", "suspended by thin pieces of dental floss", "surrounded by safety cones"
    )
    $signature     = @(
        "- Signed, Steve ($(Get-Random -InputObject @('day','night'))) shift", "- Logged by the AI that's replacing us", 
        "- Per the request of the Shadow Government", "- Sent from my Smart-Toaster", "- Dictated but not read",
        "- From the desk of a very tired human", "- Automatically generated by the Chaos Protocol", 
        "- Verified by the Warehouse Gremlin", "- XOXO, Logistics Dept.", "- Written in the dark"
    )

    # --- Structural Templates ---
    $templates = @(
        { "$((Get-Random $subjects)) $((Get-Random $status)) $((Get-Random $actions)) $((Get-Random $timing))." },
        { "Alert: $((Get-Random $subjects)) is $((Get-Random $sounds)) and $((Get-Random $smells)). $((Get-Random $safety))!" },
        { "$((Get-Random $complaints)) $((Get-Random $subjects)) $((Get-Random $status)) $((Get-Random $conditions))." },
        { "Reference #$((Get-Random -Minimum 1 -Maximum 9999)): $((Get-Random $subjects)) was $((Get-Random $customs)) $((Get-Random $timing)). $((Get-Random $signature))." },
        { 
            $para = "$((Get-Random $subjects)) $((Get-Random $status)) $((Get-Random $actions)) " +
                    "$((Get-Random $conditions)). Note: it is $((Get-Random $sounds)). " +
                    "Customs update: $((Get-Random $customs)). " +
                    "$((Get-Random $safety)). $((Get-Random $warnings)). $((Get-Random $signature))."
            $para
        }
    )

    $result = &(Get-Random -InputObject $templates)

    if ((Get-Random -Minimum 1 -Maximum 100) -le 15) 
    {
        $result += " (P.S. $((Get-Random $complaints)))"
    }

    return $result
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
                    "tenant"        = $settings.tenant_domain
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

$requestSite = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($settings.SPORootSite):/sites/$($settings.SPOSite)"  `
                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}                              `
                                 -ContentType "application/json; charset=utf-8"                                                      `
                                 -Method GET


if ($null -ne $requestSite)
{
    $siteID = $requestSite.id.Split(",")[1]
}
else
{
    write-error ">> Site '$($settings.SPORootSite)' not found!`n"
    break
}

#   the List ID for $settings.SPOList

$requestList = Invoke-RestMethod -Uri  "https://graph.microsoft.com/v1.0/sites/$($siteId)/lists/$($settings.SPOList)"  `
                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}                `
                                 -ContentType "application/json; charset=utf-8"                                        `
                                 -Method GET

if ($null -ne $requestList)
{
    $listID = $requestList.id
}
else
{
    write-error ">> List '$($settings.SPOList)' not found!`n"
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
                                      -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}                                          `
                                      -ContentType "application/json; charset=utf-8"                                                                  `
                                      -Method GET


if ($null -ne $requestUsersList)
{
    $usersListID = $requestUsersList.value.id
}
else
{
    write-error ">> SharePoint Online 'User Information List' not found!`n"
    break
}

#   Getting the list of users from the SharePoint User Information List for the given site

$requestUsers = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($siteID)/lists/$($usersListID)/items?`$expand=fields(`$select=id,IsSiteAdmin,Deleted,SipAddress)"  `
                                  -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}                                                                            `
                                  -ContentType "application/json; charset=utf-8"                                                                                                    `
                                  -Method GET

#
#   The filter below removes the following:
#
#   IsSiteAdmin:   SPO site admins (my preference but you can include them if you want)
#   Deleted:       brings back only active users with an account and permissions granted to the site
#   SipAddress:    the most important filter, it will excludes the system accounts such as "SHAREPOINT\system", "NT Service\SPSearch", "Everyone", and other groups
#                  it is redudant to include "Fields/ContentType -eq "Person", filtering the data with SipAddress will limit the results only to user accounts
#
#   The final product is just a list of IDs, we won't need anything else when creating new items
#

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
        $field_Title         = "{0}, {1}" -f  $field_ItemType, $field_UnitsSold.ToString("###,###,###,###")

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
    #   The request is limited to a maximum of 20 items
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
                    $request = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/$batch'                          `
                                                 -Headers @{"Authorization" = "Bearer $($Token_GraphAPI.access_token)"}  `
                                                 -ContentType "application/json; charset=utf-8"                          `
                                                 -Body $payload                                                          `
                                                 -Method Post

                    if ($request -eq $null)
                    {
                        # TO-DO:  failed requests are not being correctly handled here (to be improved later on)

                        write-error "`n -- Error - Aborting process`n"
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

                write-error "`n>>>>  Batch {0:$($leadingZeroes)} of {1} `t`t`t :: Exception: {2} `n" -f $batch_Counter, [math]::ceiling($RunSettings_TotalRecords / 20), $_.Exception.Message
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