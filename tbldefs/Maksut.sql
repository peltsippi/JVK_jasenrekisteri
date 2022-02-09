﻿CREATE TABLE [Maksut] (
  [PID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Kortti] LONG ,
  [Summa] CURRENCY ,
  [Puumerkki] VARCHAR (255),
  [PVM] DATETIME ,
  [Maksutapa] VARCHAR (255)
)
