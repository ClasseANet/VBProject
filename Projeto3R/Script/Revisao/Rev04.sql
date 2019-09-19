/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='04', ARQATU='Rev04.sql';
/****************************************************************************
****************************************************************************/

/****** Object:  Table [dbo].[GPESQUISA]    Script Date: 09/03/2010 16:46:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[GPESQUISA](
	[IDPESQUISA] [int] IDENTITY(1,1) NOT NULL,
	[CODSIS] [varchar](20) NOT NULL,
	[IDMODU] [varchar](20) NOT NULL,
	[NOMEPESQUISA] [varchar](40) NULL,
	[TIPOPESQUISA] [varchar](20) NULL,
	[DSCPESQUISA] [varchar](200) NULL,
	[ESCOPO] [int] NULL CONSTRAINT [DF_GPESQUISA_ESCOPO]  DEFAULT ((0)),
	[PESQDEFAULT] [int] NULL,
	[PESQSQL] [text] NULL,
	[IDUSU] [varchar](10) NOT NULL,
	[IDCONEXAO] [int] NULL,
	[TAGCAMPOS] [text] NULL,
 CONSTRAINT [PK_GPESQUISA] PRIMARY KEY CLUSTERED 
(
	[IDPESQUISA] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
ALTER TABLE [dbo].[GPESQUISA]  WITH CHECK ADD  CONSTRAINT [FK_GPESQUISA_CODSIS] FOREIGN KEY([CODSIS])
REFERENCES [dbo].[SISTEMA] ([CODSIS])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[GPESQUISA] CHECK CONSTRAINT [FK_GPESQUISA_CODSIS]
GO
ALTER TABLE [dbo].[GPESQUISA]  WITH NOCHECK ADD  CONSTRAINT [FK_GPESQUISA_GCONEXOES] FOREIGN KEY([IDCONEXAO])
REFERENCES [dbo].[GCONEXOES] ([IDCONEXAO])
GO
ALTER TABLE [dbo].[GPESQUISA] NOCHECK CONSTRAINT [FK_GPESQUISA_GCONEXOES]
GO
ALTER TABLE [dbo].[GPESQUISA]  WITH CHECK ADD  CONSTRAINT [FK_GPESQUISA_IDUSU] FOREIGN KEY([IDUSU])
REFERENCES [dbo].[USUARIO] ([IDUSU])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[GPESQUISA] CHECK CONSTRAINT [FK_GPESQUISA_IDUSU]

ALTER TABLE [dbo].[GPESQUISA]  WITH CHECK ADD  CONSTRAINT [FK_GPESQUISA_IDMODU] FOREIGN KEY([IDMODU])
REFERENCES [dbo].[MODULO] ([IDMODU])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[GPESQUISA] CHECK CONSTRAINT [FK_GPESQUISA_IDMODU]
GO