--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='75', ARQATU='Rev75.sql';
/****************************************************************************
****************************************************************************/
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTPTRATAMENTO' AND C.NAME='FLGDEL')  ALTER TABLE OTPTRATAMENTO ADD FLGDEL varchar(1) NULL  default '0'  
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='SPRODUTO' AND C.NAME='FLGDEL')  ALTER TABLE SPRODUTO ADD FLGDEL varchar(1) NULL default '0' 
go
update OTPTRATAMENTO set FLGDEL='0' Where FLGDEL is null
go
update SPRODUTO set FLGDEL='0' Where FLGDEL is null
go
