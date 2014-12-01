CREATE TABLE [dbo].[Key]
(
	[id] INT NOT NULL PRIMARY KEY IDENTITY, 
	[name] NVARCHAR(MAX) NOT NULL, 
	[default] NVARCHAR(MAX) NULL, 
	[root] NCHAR(80) NOT NULL, 
	[parent] INT NULL, 
	[raw] NVARCHAR(MAX) NULL
)
