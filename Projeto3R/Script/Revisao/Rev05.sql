/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='05', ARQATU='Rev05.sql';
/****************************************************************************
****************************************************************************/


IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CPGTOSVENDA_FLAN]') AND parent_object_id = OBJECT_ID(N'[dbo].[FLAN]'))
ALTER TABLE [dbo].[FLAN] DROP CONSTRAINT [FK_CPGTOSVENDA_FLAN]
go

ALTER TABLE [dbo].[FLAN]  WITH NOCHECK ADD  CONSTRAINT [FK_CPGTOSVENDA_FLAN] FOREIGN KEY([IDLOJA], [IDVENDA], [IDPGTO])
REFERENCES [dbo].[CPGTOSVENDA] ([IDLOJA], [IDVENDA], [IDPGTO])
ON UPDATE CASCADE
GO

ALTER TABLE [dbo].[FLAN] NOCHECK CONSTRAINT [FK_CPGTOSVENDA_FLAN]
GO

ALTER TABLE FLAN ADD IDPAI INT
GO
ALTER TABLE FLAN ADD FLGDELETE INT;
GO
exec sp_bindefault DF_0, 'FLAN.FLGDELETE'
go
