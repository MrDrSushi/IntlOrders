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

Connect-PnPOnline -Url "https://$($settings.SPORootSite)/sites/$($settings.SPOSite)" `
                  -ClientId     $($settings.SPOAddinClientId)  `
                  -ClientSecret $($settings.SPOAddinClientSecret)  `
                  -WarningAction Ignore

if (Get-PnpList -Identity  "$($settings.SPOList)")
{
    Remove-PnPList -Identity "$($settings.SPOList)" -Force:$true -Recycle:$true
}

#region ══════════════════════════════════════════════════════════════════════════════════════[ Additional Types ]

$bingAPIKey = $settings.BingAPIKey

$customFormaterMaps = @"
{
  "`$schema": "http://columnformatting.sharepointpnp.com/columnFormattingSchema.json",
  "elmType": "div",
  "style": {
    "border": "=if( @currentField != '', '2px solid #666666','')",
    "width": "=if( @currentField != '', '128px','')",
    "height": "=if( @currentField != '', '64px','')"
  },
  "children": [
    {
      "elmType": "a",
      "attributes": {
        "href": "=if( @currentField != '', 'https://www.bing.com/maps?where='+'@currentField','')",
        "target": "_blank",
        "title": "@currentField"
      },
      "style": {
        "height": "100%"
      },
      "children": [
        {
          "elmType": "img",
          "attributes": {
            "src": "=if( @currentField != '', 'https://dev.virtualearth.net/REST/v1/Imagery/Map/Road/'+'@currentField'+'?mapSize=128,64'+'&key=$bingAPIKey' , '')"
          }
        }
      ]
    }
  ]
}
"@

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


# Creating the new list

New-PnPList -Title "International Orders" -Url "$($settings.SPOList)" -Template GenericList  -OnQuickLaunch  | Format-Table

# Item Type

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Item Type" -InternalName "ItemType" -Type Choice -Choices $ItemType -AddToDefaultView -Group "Main"


# Item SKU

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Item SKU"  -InternalName "ItemSKU" -Type Text -AddToDefaultView -Group "Main"

    $field
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id  -Values @{MaxLength=36}


# Sector

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sector" -InternalName "Sector" -Type Choice -Choices $Sector -AddToDefaultView -Group "Main"


# Confidential

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Confidential" -InternalName "Confidential" -Type Boolean -AddToDefaultView -Group "Main"


# Order ID

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Order ID" -InternalName "OrderID" -Type Integer -AddToDefaultView -Group "Main"


# Order Priority

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Order Priority" -InternalName "OrderPriority" -Type Choice -Choices $OrderPriority -AddToDefaultView -Group "Main"


# Order Date

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Order Date" -InternalName "OrderDate" -Type DateTime -AddToDefaultView -Group "Main"

    $field
    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("Format","DateOnly")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Units Sold

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Units Sold" -InternalName "UnitsSold" -Type Number -AddToDefaultView -Group "Main"

    $field
    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("Decimals", "0")
    $schema.Field.SetAttribute("Format", "Dropdown")
    $schema.Field.SetAttribute("CommaSeparator","TRUE")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Unit Price

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Unit Price" -InternalName "UnitPrice" -Type Currency -AddToDefaultView -Group "Main"


# Unit Cost

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Unit Cost" -InternalName "UnitCost" -Type Currency -AddToDefaultView -Group "Main"


# Total Revenue

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Total Revenue" -InternalName "TotalRevenue" -Type Currency -AddToDefaultView -Group "Main"


# Total Cost

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Total Cost" -InternalName "TotalCost" -Type Currency -AddToDefaultView -Group "Main"


# Total Profit

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Total Profit" -InternalName "TotalProfit" -Type Currency -AddToDefaultView -Group "Main"


# Containers

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Containers" -InternalName "Containers" -Type Number -AddToDefaultView -Group "Main"

    $field

    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("Decimals", "0")
    $schema.Field.SetAttribute("Format", "Dropdown")
    $schema.Field.SetAttribute("CommaSeparator","TRUE")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Freight Terms

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Freight Terms" -InternalName "FreightTerms" -Type Choice -Choices $FreightTerms -AddToDefaultView -Group "Main"

# Sales Channel

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sales Channel" -InternalName "SalesChannel" -Type Choice -Choices $SalesChannel -AddToDefaultView -Group "Main"


# Sales Coordinator

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sales Coordinator" -InternalName "SalesCoordinator"  -Type User -AddToDefaultView -Group "Main"

    $field

    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Sales Person

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sales Person" -InternalName "SalesPerson" -Type User -AddToDefaultView -Group "Main"

    $field

    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Payment Coordinator

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Payment Coordinator" -InternalName "PaymentCoordinator" -Type User -AddToDefaultView -Group "Main"

    $field

    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Shipping Foreman

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Foreman" -InternalName "ShippingForeman" -Type User -AddToDefaultView -Group "Main"

    $field

    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Shipping Insured

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Insured" -InternalName "ShippingInsured" -Type Boolean  -AddToDefaultView -Group "Main"


# Shipping Date

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Date" -InternalName "ShippingDate" -Type DateTime -AddToDefaultView -Group "Main"

    $field

    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("Format","DateOnly")
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Shipping Method

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Method"  -InternalName "ShippingMethod" -Type Choice -Choices $ShippingMethod -AddToDefaultView -Group "Main"


# Vessel Name

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Vessel Name" -InternalName "VesselNameOrID" -Type Choice -Choices $VesselName -AddToDefaultView -Group "Main"


# Port Of Origin

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Origin" -InternalName "PortOfOrigin" -Type Text -AddToDefaultView -Group "Main"

    $field
    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("CustomFormatter", $customFormaterMaps)
    $schema.Field.SetAttribute("MaxLength", 90)
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Port Of Origin Name

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Origin Name"  -InternalName "PortOfOriginName" -Type Text -AddToDefaultView -Group "Main"

    $field
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{MaxLength=40}


# Port Of Destiny

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Destiny" -InternalName "PortOfDestiny" -Type Text -AddToDefaultView -Group "Main"

    $field
    [xml]$schema = $field.SchemaXml
    $schema.Field.SetAttribute("CustomFormatter", $customFormaterMaps)
    $schema.Field.SetAttribute("MaxLength", 90)
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{SchemaXml=$schema.OuterXml} -UpdateExistingLists


# Port Of Destiny Name

$field = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Destiny Name" -InternalName "PortOfDestinyName" -Type Text -AddToDefaultView -Group "Main"

    $field
    Set-PnPField -List "$($settings.SPOList)" -Identity $field.Id -Values @{MaxLength=40}


# Shipping Notes

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Notes" -InternalName "ShippingNotes" -Type Note -AddToDefaultView -Group "Main"


# Comments

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Comments" -InternalName "Comments" -Type Note -AddToDefaultView -Group "Main"


# Title Field

Set-PnPField -List "$($settings.SPOList)" -Identity Title -Values @{Group="Main"}