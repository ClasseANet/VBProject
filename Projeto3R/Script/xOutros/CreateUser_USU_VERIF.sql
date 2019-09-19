USE [G3R]
GO
IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'USU_VERIF') DROP USER [USU_VERIF]
GO
IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'DBA') DROP USER [DBA]
GO
CREATE USER [USU_VERIF] FOR LOGIN [USU_VERIF] WITH DEFAULT_SCHEMA=[dbo]
GO
EXEC sp_addrolemember db_datareader, [USU_VERIF]  
go 
EXEC sp_addrolemember db_datawriter , [USU_VERIF]  
go 

