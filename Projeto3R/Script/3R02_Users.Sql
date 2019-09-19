-----------------------------------------------------------------------------------
USE [G3R];
-----------------------------------------------------------------------------------
IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'USU_VERIF') DROP LOGIN [USU_VERIF];
CREATE LOGIN [USU_VERIF] WITH PASSWORD=N'MINOTAURO', DEFAULT_DATABASE=[G3R], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=ON;
--ALTER LOGIN [USU_VERIF] DISABLE;
-----------------------------------------------------------------------------------
IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'USU_VERIF') DROP USER [USU_VERIF];
CREATE USER [USU_VERIF] FOR LOGIN [USU_VERIF] WITH DEFAULT_SCHEMA=[dbo];
GRANT CONNECT TO [USU_VERIF] AS [dbo];
-------------------------------------------------------------------------------------
EXEC [sp_addrolemember] @rolename = 'db_datareader', @membername = 'USU_VERIF';
EXEC [sp_addrolemember] @rolename = 'db_datawriter', @membername = 'USU_VERIF';
EXEC [sp_addrolemember] @rolename = 'db_backupoperator', @membername = 'USU_VERIF';
EXEC sys.sp_addsrvrolemember @loginame = N'USU_VERIF', @rolename = N'sysadmin';
-----------------------------------------------------------------------------------
IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'DBA') DROP LOGIN [DBA];
CREATE LOGIN [DBA] WITH PASSWORD=N'BRASIL', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=ON;
--ALTER LOGIN [DBA] DISABLE;
-------------------------------------------------------------------------------------
IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'DBA') DROP USER [DBA];
CREATE USER [DBA] WITHOUT LOGIN WITH DEFAULT_SCHEMA=[dbo];
GRANT CONNECT TO [DBA] AS [dbo];
-------------------------------------------------------------------------------------
EXEC [sp_addrolemember] @rolename = 'db_owner', @membername = 'DBA';
EXEC [sp_addrolemember] @rolename = 'db_accessadmin', @membername = 'DBA';