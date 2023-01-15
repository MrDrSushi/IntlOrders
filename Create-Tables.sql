if not exists (select * from sysobjects where name = 'Locations' and xtype = 'U')
   create table dbo.Locations
   (
         City              nvarchar(60),
         CityASCII         nvarchar(60),
         Latitude          float,
         Longitude         float,
         Country           nvarchar(50),
         ISO2              nvarchar(3),
         ISO3              nvarchar(3),
         AdminName         nvarchar(60),
         Capital           nvarchar(10),
         Population        int
   )
go

if not exists (select * from sysobjects where name = 'Airports' and xtype = 'U')
   create table dbo.Airports
   (
         AirportType       nvarchar(14),
         AirportName       nvarchar(90),
         Latitude          float,
         Longitude         float,
         ElevationFeet     nvarchar(60),
         Continent         nvarchar(3),
         Country           nvarchar(50),
         ISO2              nvarchar(60),
         ISORegion         nvarchar(10),
         Municipality      nvarchar(50)
   )
go

if not exists (select * from sysobjects where name = 'Ports' and xtype = 'U')
   create table dbo.Ports
   (
      PortName          nvarchar(14),
      AlternateName     nvarchar(100),
      Country           nvarchar(50),
      WaterBody         nvarchar(60),
      HarborSize        nvarchar(10),
      HarborType        nvarchar(20),
      HarborUse         nvarchar(10),
      Railway           nvarchar(10),
      Latitude          float,
      Longitude         float
   )
