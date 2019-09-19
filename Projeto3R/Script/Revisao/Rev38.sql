/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='38', ARQATU='Rev38.sql';
/****************************************************************************
****************************************************************************/

IF  EXISTS (SELECT * FROM USUARIO WHERE IDUSU='DPIL') DELETE FROM USUARIO WHERE IDUSU= 'LOJA'
go
UPDATE USUARIO SET IDUSU='LOJA' WHERE IDUSU= 'DPIL'
go

