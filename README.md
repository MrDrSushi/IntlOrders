# **International Orders**


### **Purpose**
  \
The idea for this project came after different projec


### **Script Files**


The three main script files from this repo are:

- Create-SPO-List-Orders.ps1
- GraphBatch-AddItems.ps1
- GraphBatch-Delete.ps1

The scripts above do not store any personal information about the Azure tenant, endpoints secrets and ids, Bing API key, the names for the SharePoint objects such as root site, list, add-in information.

I've opted to use a central configuration file called "**settings.json**" (the original file is not synced to GitHub for obvious reasons)

  \


[Create-SPO-List-Orders.ps1](Create-SPO-List-Orders.ps1)

Use this script to create the SharePoint Online list, if there is already a list with the same name as specified in the settings.json, the existing list will be deleted and sent to the recycle bin.

The script creates the new list, adds all the fields with its particular properties, such as column formatting, length, required, and any additional formats regarding the precision and types.

#

[GraphBatch-AddItems.ps1](GraphBatch-AddItems.ps1)

#

[GraphBatch-Delete.ps1](GraphBatch-Delete.ps1)

#

[Sample-QueryUsers.ps1](Sample-QueryUsers.ps1)

#

[orders-layout-body.json](orders-layout-body.json)

#

[orders-layout-header.json](orders-layout-header.json)

#

[settings-example.json](settings-example.json)






### **Data Sources**


[Create-Tables.sql](Create-Tables.sql)

[world-data-Airports.csv](world-data-Airports.csv)

[world-data-Locations.csv](world-data-Locations.csv)

[world-data-Ports.csv](world-data-Ports.csv)

[world-data.xlsx](world-data.xlsx)
