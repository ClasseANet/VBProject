/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='06', ARQATU='Rev06.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE SMOVEST ADD FLGDELETE INT;
GO
exec sp_bindefault DF_0, 'SMOVEST.FLGDELETE'
go
IF  EXISTS (SELECT T.NAME, C.* FROM sys.COLUMNS C JOIN sys.TABLES T ON T.OBJECT_ID=C.OBJECT_ID WHERE T.NAME='OSESSAO' AND  C.NAME='VALOR')
ALTER TABLE OSESSAO DROP COLUMN VALOR;
GO
IF  EXISTS (SELECT T.NAME, C.* FROM sys.COLUMNS C JOIN sys.TABLES T ON T.OBJECT_ID=C.OBJECT_ID WHERE T.NAME='OSESSAO' AND  C.NAME='HHINI')
ALTER TABLE OSESSAO DROP COLUMN HHINI;
GO
IF  EXISTS (SELECT T.NAME, C.* FROM sys.COLUMNS C JOIN sys.TABLES T ON T.OBJECT_ID=C.OBJECT_ID WHERE T.NAME='OSESSAO' AND  C.NAME='HHFIM')
ALTER TABLE OSESSAO DROP COLUMN HHFIM;
GO
