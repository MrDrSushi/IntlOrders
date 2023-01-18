# **International Orders**

<br>
<br>

## **Purpose**

Working with SharePoint Online isn't easy when you are a developer, creating and maintaining mocking data for tests is quite a job, and that's the whole purpose for this repo, it allows me to quickly create and populate a SharePoint Online list with as many items as I need.

One of the main points is to allow me to create huge lists with thousands of items, like in one particular project where I was tasked to conduct a stress and performance test with a list containing over 1 million items.

There are two mains scripts in this repo for daily use, one to quickly add items to the list called  [GraphBatch-AddItems.ps1](GraphBatch-AddItems.ps1), and another script called [GraphBatch-Delete.ps1](GraphBatch-Delete.ps1) used for the deletion of items.

You might wonder why there is an specific script for the deletion of list items, and the reason is simple, lists that exceed the SharePoint list view threshold can't be deleted from SharePoint, and it is your job to clean up the items to bring the list down for a sucessul removal. The script it is also useful for a quick maintenance to delete any amount of unwanted items.

And last but not least, there is also a script for the creation of the list called [Create-SPO-List-Orders.ps1](Create-SPO-List-Orders.ps1), it creates a brand new list containing all the columns, the script properly configures each column with their length, precisions, data types, and formatting rules.
<br>
<br>

## **Getting Started**

Download all the files or simply clone the repo to your local machine:

```
git clone https://github.com/MrDrSushi/IntlOrders/
```

Create a copy from "settings-example.json" and rename it into "settings.json", update its contents with the actual values matching your environment, keep this new file together with all the scripts, this will be used by all the scripts in order to gain access to your tenant, in the future I will include some level of encryption instead of plain file, for now it is just a small workaround, so keep it safe!
<br>
<br>

## **Script Files**



  \


[Create-SPO-List-Orders.ps1](Create-SPO-List-Orders.ps1)

Use this script to create the SharePoint Online list, if there is already a list with the same name as specified in the settings.json, the existing list will be deleted and sent to the recycle bin.

The script creates the new list, adds all the fields with its particular properties, such as column formatting, length, required, and any additional formats regarding the precision and types.

#

### [GraphBatch-AddItems.ps1](GraphBatch-AddItems.ps1)

#

### [GraphBatch-Delete.ps1](GraphBatch-Delete.ps1)

this is the clean up script, you can use this script for removing all items or just small amounts of items from your list, the only part of the script where you want to update is a variable called:
 ```$itemDeleteEnd = 5000```

In the script this place is located at line number 80 and looks like the following:

```powershell
#   where to start and finish (SPO List item IDs)

$itemDeleteStart = $requestID.value.id
$itemDeleteEnd   = 5000
```

In the example above, the number ```5000``` is the ID for the last item you wanted to be deleted from your list, the script starts from the very first available ID in your list and advances until it reaches the end of the list of a matching item with ```ID: 5000```.

If the first item available on your list is the ```ID: 39```, the script will start from there, sequentially deleting batches of 20 items every time until it reaches an item with matching the ```ID: 5000``` or the end of the list is reached (in case your didn't specify a valid ID).

#

### [Sample-QueryUsers.ps1](Sample-QueryUsers.ps1)

#

### [orders-layout-header.json](orders-layout-header.json)

**under construction** - custom form for the list, this will be the header for the form.

#

### [orders-layout-body.json](orders-layout-body.json)

**under construction** - the body for the custom form for the list - the footer will be added later

#

### [settings-example.json](settings-example.json)

I've opted for a central configuration file to keep me from updating the scripts, this saves time from updating files individually if anything changes on my tenant, make a copy of this file into a new called called "**settings.json**" on your local machine and keep it together with all the scripts, this file will be used by the scripts in order to gain access to your tenant.





### **Data Sources**


[Create-Tables.sql](Create-Tables.sql)

[world-data-Airports.csv](world-data-Airports.csv)

[world-data-Locations.csv](world-data-Locations.csv)

[world-data-Ports.csv](world-data-Ports.csv)

[world-data.xlsx](world-data.xlsx)