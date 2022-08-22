CREATE TABLE [Lataukset] (
  [PID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Kortti] LONG ,
  [Voimassa] DATETIME ,
  [Puumerkki] VARCHAR (255),
  [Korttityyppi] VARCHAR (255),
  [KortinArvo] CURRENCY ,
  [Ajankohta] DATETIME 
)
