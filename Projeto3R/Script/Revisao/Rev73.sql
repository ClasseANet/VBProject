--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='73', ARQATU='Rev73.sql';
/****************************************************************************
****************************************************************************/
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OSESSAO' AND C.NAME='IDITEM')  ALTER TABLE OSESSAO ADD IDITEM int NULL default 0
go
