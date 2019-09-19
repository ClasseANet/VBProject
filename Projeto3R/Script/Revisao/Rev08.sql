/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='08', ARQATU='Rev08.sql';
/****************************************************************************
****************************************************************************/
--SELECT * FROM MODULO
Update modulo
SET VBSCRIPT = '
SUB DB_BACKUP()
   SET NG = CREATEOBJECT("CALENDARIO3R.NG_CALENDARIO")
   WITH NG
      SET .SYS=SYS
      .BACKUP
   END WITH
END SUB'
, IDMODU = 'BAK'
, DSCMODU='Backup'
WHERE IDMODU='BAK'
GO

ALTER TABLE OCLIENTE ADD DSCTPCONHEC varchar(20);