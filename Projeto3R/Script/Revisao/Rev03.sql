/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='03', ARQATU='Rev03.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE [dbo].[OMAQDISPAROS] NOCHECK CONSTRAINT [R_226]
GO
ALTER TABLE [dbo].[OMAQDISPAROS] NOCHECK CONSTRAINT [R_230]
GO