--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='70', ARQATU='Rev70.sql';
/****************************************************************************
****************************************************************************/
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OCLIENTE' AND C.NAME='IDFUNC')  ALTER TABLE OCLIENTE ADD IDFUNC int NULL
go
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_EVENTO]') AND parent_object_id = OBJECT_ID(N'[dbo].[OTAREFAEVT]'))
ALTER TABLE [dbo].[OTAREFAEVT] DROP CONSTRAINT [FK_EVENTO]
GO
ALTER TABLE [dbo].[OTAREFAEVT]  WITH NOCHECK ADD  CONSTRAINT [FK_EVENTO] FOREIGN KEY([IDLOJA], [IDEVENTO])
REFERENCES [dbo].[OEVENTOAGENDA] ([IDLOJA], [IDEVENTO])
GO
ALTER TABLE [dbo].[OTAREFAEVT] NOCHECK CONSTRAINT [FK_EVENTO]
GO
