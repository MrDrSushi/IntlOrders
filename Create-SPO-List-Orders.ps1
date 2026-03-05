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

Connect-PnPOnline -Url                  "https://$($settings.SPORootSite)/sites/$($settings.SPOSite)" `
                  -ClientId             "$($settings.client_id)"                                      `
                  -CertificatePassword  (ConvertTo-SecureString -String $($settings.certificate_password) -AsPlainText -Force)    `
                  -CertificatePath      ".\$($settings.entra_applicationname).pfx"                                                      `
                  -Tenant               $settings.tenant_domain


if (Get-PnpList -Identity  "$($settings.SPOList)" -ThrowExceptionIfListNotFound:$false -ErrorAction SilentlyContinue)
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

$ShippingNames = @(
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


# ════ Creating the new list

New-PnPList -Title "Orders" -Url "$($settings.SPOList)" -Template GenericList -EnableContentTypes -OnQuickLaunch | Format-Table

Set-PnPList -Identity "$($settings.SPOList)" -EnableAttachments:$false -EnableFolderCreation:$false -EnableVersioning:$false -EnableModeration:$false | Format-Table


# ════ Title

Set-PnPField -List "$($settings.SPOList)" -Identity Title -Values @{Required=$false;Hidden=$true}


# ════ Item Type

$fieldItemType = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Item Type" -InternalName "ItemType" -Type Choice -Choices $ItemType -Required -AddToDefaultView -Group "Main"

    $fieldItemType
    $fieldItemType.DefaultValue = "Air Fryer"
    $fieldItemType.Update()
    $fieldItemType.Context.ExecuteQuery()


# ════ Item SKU

$fieldItemSKU = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Item SKU"  -InternalName "ItemSKU" -Type Text -AddToDefaultView -Group "Main"

    $fieldItemSKU
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldItemSKU.Id  -Values @{MaxLength=36}


# ════ Sector

$fieldSector = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sector" -InternalName "Sector" -Type Choice -Choices $Sector -Required -AddToDefaultView -Group "Main"

    $fieldSector
    $fieldSector.DefaultValue = "Private"
    $fieldSector.Update()
    $fieldSector.Context.ExecuteQuery()


# ════ Confidential

$fieldConfidential = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Confidential" -InternalName "Confidential" -Type Boolean -Required -AddToDefaultView -Group "Main"

    $fieldConfidential
    $fieldConfidential.DefaultValue = 0
    $fieldConfidential.Update()
    $fieldConfidential.Context.ExecuteQuery()


# ════ Order ID

$fieldOrderID = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Order ID" -InternalName "OrderID" -Type Number -Required -AddToDefaultView -Group "Main"

    $fieldOrderID
    [xml]$schemaOrderID = $fieldOrderID.SchemaXml
    $schemaOrderID.Field.SetAttribute("Decimals", "0")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldOrderID.Id -Values @{SchemaXml=$schemaOrderID.OuterXml} -UpdateExistingLists


# ════ Order Priority

$fieldOrderPriority = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Order Priority" -InternalName "OrderPriority" -Type Choice -Choices $OrderPriority -Required -AddToDefaultView -Group "Main"

    $fieldOrderPriority
    $fieldOrderPriority.DefaultValue = "Low"
    $fieldOrderPriority.Update()
    $fieldOrderPriority.Context.ExecuteQuery()


# ════ Order Date

$fieldOrderDate = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Order Date" -InternalName "OrderDate" -Type DateTime -Required -AddToDefaultView -Group "Main"

    $fieldOrderDate
    [xml]$schemaOrderDate = $fieldOrderDate.SchemaXml
    $schemaOrderDate.Field.SetAttribute("Format","DateOnly")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldOrderDate.Id -Values @{SchemaXml=$schemaOrderDate.OuterXml} -UpdateExistingLists

    $fieldOrderDate = Get-PnPField -List "$($settings.SPOList)" -Identity "OrderDate"
    $fieldOrderDate.DefaultFormula = "=NOW()"
    $fieldOrderDate.Update()
    $fieldOrderDate.Context.ExecuteQuery()


# ════ Units Sold

$fieldUnitsSold = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Units Sold" -InternalName "UnitsSold" -Type Currency -Required -AddToDefaultView -Group "Main"

    $fieldUnitsSold
    $fieldUnitsSold.DefaultValue = 0
    $fieldUnitsSold.Update()
    $fieldUnitsSold.Context.ExecuteQuery()


# ════ Unit Price

$fieldUnitPrice = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Unit Price" -InternalName "UnitPrice" -Type Currency -Required -AddToDefaultView -Group "Main"

    $fieldUnitPrice
    $fieldUnitPrice.DefaultValue = 0
    $fieldUnitPrice.Update()
    $fieldUnitPrice.Context.ExecuteQuery()


# ════ Unit Cost

$fieldUnitCost = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Unit Cost" -InternalName "UnitCost" -Type Currency -Required -AddToDefaultView -Group "Main"

    $fieldUnitCost
    $fieldUnitCost.DefaultValue = 0
    $fieldUnitCost.Update()
    $fieldUnitCost.Context.ExecuteQuery()


# ════ Total Revenue

$fieldTotalRevenue = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Total Revenue" -InternalName "TotalRevenue" -Type Currency -Required -AddToDefaultView -Group "Main"

    $fieldTotalRevenue
    $fieldTotalRevenue.DefaultValue = 0
    $fieldTotalRevenue.Update()
    $fieldTotalRevenue.Context.ExecuteQuery()


# ════ Total Cost

$fieldTotalCost = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Total Cost" -InternalName "TotalCost" -Type Currency -Required -AddToDefaultView -Group "Main"

    $fieldTotalCost
    $fieldTotalCost.DefaultValue = 0
    $fieldTotalCost.Update()
    $fieldTotalCost.Context.ExecuteQuery()


# ════ Total Profit

$fieldTotalProfit = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Total Profit" -InternalName "TotalProfit" -Type Currency -Required -AddToDefaultView -Group "Main"

    $fieldTotalProfit
    $fieldTotalProfit.DefaultValue = 0
    $fieldTotalProfit.Update()
    $fieldTotalProfit.Context.ExecuteQuery()


# ════ Containers

$fieldContainers = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Containers" -InternalName "Containers" -Type Number -Required -AddToDefaultView -Group "Main"

    $fieldContainers
    [xml]$schemaContainers = $fieldContainers.SchemaXml
    $schemaContainers.Field.SetAttribute("Decimals", "0")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldContainers.Id -Values @{SchemaXml=$schemaContainers.OuterXml} -UpdateExistingLists

    $fieldContainers = Get-PnPField -List "$($settings.SPOList)" -Identity "Containers"
    $fieldContainers.DefaultValue = 0
    $fieldContainers.Update()
    $fieldContainers.Context.ExecuteQuery()


# ════ Freight Terms

$fieldFreightTerms = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Freight Terms" -InternalName "FreightTerms" -Type Choice -Choices $FreightTerms -Required -AddToDefaultView -Group "Main"

    $fieldFreightTerms
    $fieldFreightTerms.DefaultValue = "Elsewhere"
    $fieldFreightTerms.Update()
    $fieldFreightTerms.Context.ExecuteQuery()


# ════ Sales Channel

$fieldSalesChannel = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sales Channel" -InternalName "SalesChannel" -Type Choice -Choices $SalesChannel -Required -AddToDefaultView -Group "Main"

    $fieldSalesChannel
    $fieldSalesChannel.DefaultValue = "Sales Rep"
    $fieldSalesChannel.Update()
    $fieldSalesChannel.Context.ExecuteQuery()


# ════ Sales Coordinator

$fieldSalesCoordinator = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sales Coordinator" -InternalName "SalesCoordinator"  -Type User -AddToDefaultView -Group "Main"

    $fieldSalesCoordinator
    [xml]$schemaSalesCoordinator = $fieldSalesCoordinator.SchemaXml
    $schemaSalesCoordinator.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldSalesCoordinator.Id -Values @{SchemaXml=$schemaSalesCoordinator.OuterXml} -UpdateExistingLists


# ════ Sales Person

$fieldSalesPerson = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Sales Person" -InternalName "SalesPerson" -Type User -AddToDefaultView -Group "Main"

    $fieldSalesPerson
    [xml]$schemaSalesPerson = $fieldSalesPerson.SchemaXml
    $schemaSalesPerson.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldSalesPerson.Id -Values @{SchemaXml=$schemaSalesPerson.OuterXml} -UpdateExistingLists


# ════ Payment Coordinator

$fieldPaymentCoordinator = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Payment Coordinator" -InternalName "PaymentCoordinator" -Type User -AddToDefaultView -Group "Main"

    $fieldPaymentCoordinator
    [xml]$schemaPaymentCoordinator = $fieldPaymentCoordinator.SchemaXml
    $schemaPaymentCoordinator.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldPaymentCoordinator.Id -Values @{SchemaXml=$schemaPaymentCoordinator.OuterXml} -UpdateExistingLists


# ════ Shipping Foreman

$fieldShippingForeman = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Foreman" -InternalName "ShippingForeman" -Type User -AddToDefaultView -Group "Main"

    $fieldShippingForeman
    [xml]$schemaShippingForeman = $fieldShippingForeman.SchemaXml
    $schemaShippingForeman.Field.SetAttribute("UserDisplayOptions","NamePhoto")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldShippingForeman.Id -Values @{SchemaXml=$schemaShippingForeman.OuterXml} -UpdateExistingLists


# ════ Shipping Insured

$fieldShippingInsured = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Insured" -InternalName "ShippingInsured" -Type Boolean -Required -AddToDefaultView -Group "Main"

    $fieldShippingInsured
    $fieldShippingInsured.DefaultValue = 0
    $fieldShippingInsured.Update()
    $fieldShippingInsured.Context.ExecuteQuery()


# ════ Shipping Date

$fieldShippingDate = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Date" -InternalName "ShippingDate" -Type DateTime -Required -AddToDefaultView -Group "Main"

    $fieldShippingDate
    [xml]$schemaShippingDate = $fieldShippingDate.SchemaXml
    $schemaShippingDate.Field.SetAttribute("Format","DateOnly")
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldShippingDate.Id -Values @{SchemaXml=$schemaShippingDate.OuterXml} -UpdateExistingLists

    $fieldShippingDate = Get-PnPField -List "$($settings.SPOList)" -Identity "ShippingDate"
    $fieldShippingDate.DefaultFormula = "=NOW()+30"
    $fieldShippingDate.Update()
    $fieldShippingDate.Context.ExecuteQuery()


# ════ Shipping Method

$fieldShippingMethod = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Method"  -InternalName "ShippingMethod" -Type Choice -Choices $ShippingMethod -Required -AddToDefaultView -Group "Main"

    $fieldShippingMethod
    $fieldShippingMethod.DefaultValue = "Land"
    $fieldShippingMethod.Update()
    $fieldShippingMethod.Context.ExecuteQuery()


# ════ Vessel Name

$VesselNames = $AirlineNames + $ShippingNames

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Vessel Name" -InternalName "VesselNameOrID" -Type Choice -Choices $VesselNames -AddToDefaultView -Group "Main"


# ════ Port Of Origin

$fieldPortOfOrigin = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Origin" -InternalName "PortOfOrigin" -Type Text -AddToDefaultView -Group "Main"

    $fieldPortOfOrigin
    [xml]$schemaPortOfOrigin = $fieldPortOfOrigin.SchemaXml
    $schemaPortOfOrigin.Field.SetAttribute("CustomFormatter", $customFormaterMaps)
    $schemaPortOfOrigin.Field.SetAttribute("MaxLength", 90)
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldPortOfOrigin.Id -Values @{SchemaXml=$schemaPortOfOrigin.OuterXml} -UpdateExistingLists


# ════ Port Of Origin Name

$fieldPortOfOriginName = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Origin Name"  -InternalName "PortOfOriginName" -Type Text -AddToDefaultView -Group "Main"

    $fieldPortOfOriginName
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldPortOfOriginName.Id -Values @{MaxLength=40}


# ════ Port Of Destiny

$fieldPortOfDestiny = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Destiny" -InternalName "PortOfDestiny" -Type Text -AddToDefaultView -Group "Main"

    $fieldPortOfDestiny
    [xml]$schemaPortOfDestiny = $fieldPortOfDestiny.SchemaXml
    $schemaPortOfDestiny.Field.SetAttribute("CustomFormatter", $customFormaterMaps)
    $schemaPortOfDestiny.Field.SetAttribute("MaxLength", 90)
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldPortOfDestiny.Id -Values @{SchemaXml=$schemaPortOfDestiny.OuterXml} -UpdateExistingLists


# ════ Port Of Destiny Name

$fieldPortOfDestinyName = Add-PnPField -List "$($settings.SPOList)" -DisplayName "Port Of Destiny Name" -InternalName "PortOfDestinyName" -Type Text -AddToDefaultView -Group "Main"

    $fieldPortOfDestinyName
    Set-PnPField -List "$($settings.SPOList)" -Identity $fieldPortOfDestinyName.Id -Values @{MaxLength=40}


# ════ Shipping Notes

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Shipping Notes" -InternalName "ShippingNotes" -Type Note -AddToDefaultView -Group "Main"


# ════ Comments

Add-PnPField -List "$($settings.SPOList)" -DisplayName "Comments" -InternalName "Comments" -Type Note -AddToDefaultView -Group "Main"


# ════ Hiding the unwanted Title field


$view = Get-PnPView -List Orders -Identity "All Items"

$view.ViewFields.Remove("LinkTitle")
$view.Update()
$view.Context.ExecuteQuery()