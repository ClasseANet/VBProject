/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='18', ARQATU='Rev18.sql';
/****************************************************************************
****************************************************************************/
Alter Table CVENDA ADD [IDPROMO] [int] NULL;


CREATE TABLE CPROMOCAO
(	[IDPROMO]    int IDENTITY (1,1) NOT NULL,	
	[ALTERSTAMP] int NULL,
	[TIMESTAMP]  datetime NULL,	
	[DSCPROMO]	 varchar(30) NULL,
	[ATIVO]		 int NULL,
	[VLDESC]	 decimal(9,2) NULL ,
	[VALIDADE]   datetime NULL,
 CONSTRAINT [PK_CPROMOCAO] PRIMARY KEY CLUSTERED 
([IDPROMO] ASC)
WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY];

exec sp_bindefault DF_1, 'CPROMOCAO.ATIVO';
exec sp_bindefault DF_1, 'CPROMOCAO.ALTERSTAMP';
exec sp_bindefault DF_Now, 'CPROMOCAO.TIMESTAMP';

ALTER TABLE [dbo].[CVENDA]  WITH NOCHECK ADD  CONSTRAINT [FK_CPROMO] FOREIGN KEY([IDPROMO]) REFERENCES [dbo].[CPROMOCAO] ([IDPROMO])
GO
