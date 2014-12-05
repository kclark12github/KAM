Drop Table [dbo].[Key];
Drop Table [dbo].[Value];
Go
Create Table [dbo].[Key]
(
	[id] INT NOT NULL PRIMARY KEY IDENTITY, 
	[name] NVARCHAR(MAX) NOT NULL, 
	[root] NCHAR(80) NOT NULL, 
	[raw] NVARCHAR(MAX) NULL
);
--Create Index [KeyByName] On [dbo].[Key] ([name]);
Create Table [dbo].[Value]
(
	[id] INT NOT NULL PRIMARY KEY IDENTITY, 
	[name] NVARCHAR(MAX) NOT NULL, 
	[key] INT NULL, 
	[value] NVARCHAR(MAX) NULL 
);
Go