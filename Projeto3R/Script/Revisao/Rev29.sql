/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='29', ARQATU='Rev29.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE FLAN ADD FLGEXPORT int
go
exec sp_bindefault DF_0, 'FLAN.FLGEXPORT'
go
update FLAN set FLGEXPORT=0
go
ALTER TABLE FCCORRENTE ADD ATIVO int
go
exec sp_bindefault DF_1, 'FLAN.ATIVO'
go
update FCCORRENTE set ATIVO=1
go
ALTER TABLE COLIGADA ALTER COLUMN TAG varchar(200)
go
