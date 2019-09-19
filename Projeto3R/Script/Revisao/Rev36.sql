
/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='36', ARQATU='Rev36.sql';
/****************************************************************************
****************************************************************************/

UPDATE OSESSAO SET IDMANIPULO =1 WHERE IDMANIPULO =0
GO

ALTER TABLE [dbo].[CITENSVENDA]  DROP CONSTRAINT [FK_CITENSVENDA_CVENDA]
GO 
ALTER TABLE [dbo].[CITENSVENDA]  WITH CHECK ADD  CONSTRAINT [FK_CITENSVENDA_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
ON UPDATE CASCADE
GO

ALTER TABLE [dbo].[CPGTOSVENDA]  DROP CONSTRAINT [FK_CPGTOSVENDA_CVENDA]
GO 
ALTER TABLE [dbo].[CPGTOSVENDA]  WITH CHECK ADD  CONSTRAINT [FK_CPGTOSVENDA_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
ON UPDATE CASCADE
GO

ALTER TABLE [dbo].[RFUNCIONARIO] ADD IDFINGER Int 
GO 
