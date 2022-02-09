CREATE TABLE [Kortit] (
  [CID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Kortti] VARCHAR (255),
  [Omistaja] LONG ,
  [PVM] DATETIME ,
  [Puumerkki] VARCHAR (255),
  [Muistiinpanot] LONGTEXT 
)
