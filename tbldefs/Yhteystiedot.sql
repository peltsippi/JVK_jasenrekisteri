CREATE TABLE [Yhteystiedot] (
  [UID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Sukunimi] VARCHAR (50),
  [Etunimi] VARCHAR (50),
  [Sähköpostiosoite] VARCHAR (50),
  [Matkapuhelin] VARCHAR (25),
  [Kaupunki] VARCHAR (50),
  [Jäsenyys] VARCHAR (100),
  [Muistiinpanot] LONGTEXT ,
  [Edustusjasen] BIT 
)
