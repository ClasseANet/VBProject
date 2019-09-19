/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='07', ARQATU='Rev07.sql';
/****************************************************************************
****************************************************************************/

SET IDENTITY_INSERT modulo ON;
INSERT modulo(ID,IDMODU,DSCMODU,SITMODU,MODUPAI,MENUDEFAULT,IDPAI,VBSCRIPT) VALUES('17','BAK','Backup','S','ADMADM','S','2','');
SET IDENTITY_INSERT modulo off;
go
INSERT modulo_sistema(ID,IDMODU,CODSIS,IDPAI,MODUPAI,INDICE,MENU,GRUPOMENU) VALUES('17','BAK','P3R','2',NULL,'0','S','0');
go
EXEC [sp_addrolemember] @rolename = 'db_backupoperator', @membername = 'USU_VERIF';
go
EXEC sys.sp_addsrvrolemember @loginame = N'USU_VERIF', @rolename = N'sysadmin'
go
exec sp_bindefault DF_0, 'OEVENTOAGENDA.FLGCONFIRMADO'
go
ALTER TABLE OEVENTOAGENDA ADD FLGCANCELADO INT;
GO
exec sp_bindefault DF_0, 'OEVENTOAGENDA.FLGCANCELADO'
go

