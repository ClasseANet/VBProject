/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='31', ARQATU='Rev31.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE OTPMANIPULO alter column DSCMANIPULO varchar(10)  NULL
go
DELETE OTPMANIPULO Where IDTPMANIPULO>2
go

SET IDENTITY_INSERT OTPMANIPULO on;
INSERT INTO OTPMANIPULO (IDTPMANIPULO, DSCMANIPULO) VALUES (1, 'Pequeno');
INSERT INTO OTPMANIPULO (IDTPMANIPULO, DSCMANIPULO) VALUES (2, 'Grande');
SET IDENTITY_INSERT OTPMANIPULO off;


