CREATE TABLE [Korttitilasto] (
  [Tunniste] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [PVM] DATETIME ,
  [Kaikki] LONG ,
  [KkKortit] LONG ,
  [ApKortit] LONG ,
  [KrtKortit] LONG ,
  [OpiskKortit] LONG ,
  [MuuKortit] LONG 
)
