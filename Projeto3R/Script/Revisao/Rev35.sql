
/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='35', ARQATU='Rev35.sql';
/****************************************************************************
****************************************************************************/

UPDATE OSESSAO SET IDMANIPULO =1 WHERE IDMANIPULO =0
GO


IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_AV_CVENDA]') AND parent_object_id = OBJECT_ID(N'[dbo].[OATENDIMENTO_VENDA]'))
ALTER TABLE [dbo].[OATENDIMENTO_VENDA] DROP CONSTRAINT [FK_AV_CVENDA]
GO

ALTER TABLE [dbo].[OATENDIMENTO_VENDA]  WITH CHECK ADD  CONSTRAINT [FK_AV_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
ON UPDATE CASCADE
GO

ALTER TABLE [dbo].[OATENDIMENTO_VENDA] CHECK CONSTRAINT [FK_AV_CVENDA]
GO