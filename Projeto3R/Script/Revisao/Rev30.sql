/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='30', ARQATU='Rev30.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE OCLIENTE ADD IDLOJA0 int
go
ALTER TABLE OCONTATO ADD IDLOJA0 int
go
ALTER TABLE RFUNCIONARIO ADD IDLOJA0 int
go

UPDATE OCLIENTE set IDLOJA0=IDLOJA
go
UPDATE OCONTATO set IDLOJA0=IDLOJA
go
UPDATE RFUNCIONARIO set IDLOJA0=IDLOJA
go

ALTER TABLE [dbo].[OATENDIMENTO]  DROP CONSTRAINT [R_164]
GO
ALTER TABLE [dbo].[OATENDIMENTO]  WITH NOCHECK ADD  CONSTRAINT [R_164] FOREIGN KEY([IDLOJA], [IDCLIENTE])
REFERENCES [dbo].[OCLIENTE] ([IDLOJA], [IDCLIENTE])
GO
ALTER TABLE [dbo].[OATENDIMENTO] NOCHECK CONSTRAINT [R_164]
GO

ALTER TABLE [dbo].[CVENDA]  DROP CONSTRAINT [FK_CVENDA_OCLIENTE]
GO
ALTER TABLE [dbo].[CVENDA]  WITH NOCHECK ADD  CONSTRAINT [FK_CVENDA_OCLIENTE] FOREIGN KEY([IDLOJA], [IDCLIENTE])
REFERENCES [dbo].[OCLIENTE] ([IDLOJA], [IDCLIENTE])
GO
ALTER TABLE [dbo].[CVENDA] NOCHECK CONSTRAINT [FK_CVENDA_OCLIENTE]
go