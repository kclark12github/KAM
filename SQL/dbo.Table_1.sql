CREATE TABLE [dbo].[Values]
(
	[id] INT NOT NULL PRIMARY KEY IDENTITY, 
    [name] NVARCHAR(MAX) NOT NULL, 
    [key] INT NULL, 
    [value] NVARCHAR(MAX) NULL 
)
