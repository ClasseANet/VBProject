
USE [master];
EXEC SP_dropdevice '1G3R';
--EXEC sp_addumpdevice 'disk', '1G3R', 'c:\dump\G3R.bak'
--EXEC sp_addumpdevice 'disk', '1G3R', 'C:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\Backup\G3R.bak'
--EXEC sp_addumpdevice 'disk', 'G3R', 'C:\Arquivos de Programas\Microsoft SQL Server\MSSQL.1\MSSQL\Backup\G3R.bak'
EXEC sp_addumpdevice 'disk', '1G3R', 'C:\Arquivos de Programas\Microsoft SQL Server\MSSQL.1\MSSQL\Backup\1FREGUESIA.bak';
--RESTORE DATABASE [G3R] FROM DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL.1\MSSQL\Backup\G3R.bak'         WITH FILE=1, NOUNLOAD, REPLACE, STATS=10
--RESTORE DATABASE [G3R] FROM DISK = N'C:\Arquivos de Programas\Microsoft SQL Server\MSSQL.1\MSSQL\Backup\G3R.bak' WITH FILE=1, NOUNLOAD, REPLACE, STATS=10

RESTORE DATABASE [G3R] FROM [1G3R] WITH FILE=1,  NOUNLOAD, REPLACE, STATS = 10;

--RESTORE DATABASE [G3R_Freguesia] FROM [1G3R] WITH FILE=1,  NOUNLOAD, REPLACE, STATS = 10
--, MOVE 'G3R_Data' To 'C:\Arquivos de programas\Microsoft SQL Server\MSSQL.1\MSSQL\DATA\G3R_Freguesia.mdf'
--, MOVE 'G3R_Log'  To 'C:\Arquivos de programas\Microsoft SQL Server\MSSQL.1\MSSQL\DATA\G3R_Freguesia.ldf';
/*=========================================================================
=========================================================================*/
USE [G3R];
IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'USU_VERIF') DROP USER [USU_VERIF];
IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'DBA') DROP USER [DBA];
CREATE USER [USU_VERIF] FOR LOGIN [USU_VERIF] WITH DEFAULT_SCHEMA=[dbo];
EXEC [sp_addrolemember] @rolename = 'db_datareader', @membername = 'USU_VERIF';
EXEC [sp_addrolemember] @rolename = 'db_datawriter', @membername = 'USU_VERIF';
EXEC [sp_addrolemember] @rolename = 'db_backupoperator', @membername = 'USU_VERIF';
EXEC sys.sp_addsrvrolemember @loginame = N'USU_VERIF', @rolename = N'sysadmin';
/*=========================================================================
=========================================================================*/

