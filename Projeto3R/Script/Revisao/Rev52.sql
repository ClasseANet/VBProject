--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='52', ARQATU='Rev52.sql';
/****************************************************************************
****************************************************************************/

IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_FLAN_FDESPESA]') AND parent_object_id = OBJECT_ID(N'[dbo].[FLAN]'))
ALTER TABLE [dbo].[FLAN] DROP CONSTRAINT [FK_FLAN_FDESPESA]
GO

ALTER TABLE [dbo].[FLAN]  WITH NOCHECK ADD  CONSTRAINT [FK_FLAN_FDESPESA] FOREIGN KEY([IDLOJA], [IDDESP]) REFERENCES [dbo].[FDESPESA] ([IDLOJA], [IDDESP])ON UPDATE CASCADE
GO

IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OCLIENTE' AND C.NAME='NFE') ALTER TABLE [OCLIENTE] ADD NFE INT NULL
GO
UPDATE OCLIENTE SET NFE=0 WHERE NFE IS NULL
GO
exec sp_bindefault DF_0, 'OCLIENTE.NFE'
go

