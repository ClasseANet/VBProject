
/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='40', ARQATU='Rev40.sql';
/****************************************************************************
****************************************************************************/

alter table OTPMANIPULO ALTER COLUMN DSCMANIPULO VARCHAR(20)
go
alter table OTPSERVICO ALTER COLUMN DSCSERVICO VARCHAR(30)
go
alter table OSALA ADD ATIVO int
go
exec sp_bindefault DF_1, 'OSALA.ATIVO'
go
Update OSALA set ATIVO=1
go
alter table CVENDA ADD FLGBXMANUAL Int
go
exec sp_bindefault DF_0, 'CVENDA.FLGBXMANUAL'
go
Update CVENNDA set FLGBXMANUAL=0
go

Update OEVENTOAGENDA set ScheduleID=1 Where ScheduleID=0
go
ALTER TABLE CPROMOCAO ADD IDPROD INT NULL
go
ALTER TABLE CPROMOCAO ADD QTDPROD DECIMAL(9,2) NULL
go

IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_224]') AND parent_object_id = OBJECT_ID(N'[dbo].[FLAN]'))
ALTER TABLE [dbo].[FLAN] DROP CONSTRAINT [R_224]
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CVENDA_MOVEST]') AND parent_object_id = OBJECT_ID(N'[dbo].[SMOVEST]'))
ALTER TABLE [dbo].[SMOVEST] DROP CONSTRAINT [FK_CVENDA_MOVEST]
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CVENDA_FLAN]') AND parent_object_id = OBJECT_ID(N'[dbo].[FLAN]'))
ALTER TABLE [dbo].[FLAN] DROP CONSTRAINT [FK_CVENDA_FLAN]
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CVENDA_FFATURA]') AND parent_object_id = OBJECT_ID(N'[dbo].[FFATURA]'))
ALTER TABLE [dbo].[FFATURA] DROP CONSTRAINT [FK_CVENDA_FFATURA]
GO

IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CITENSVENDA_CVENDA]') AND parent_object_id = OBJECT_ID(N'[dbo].[CITENSVENDA]'))
ALTER TABLE [dbo].[CITENSVENDA] DROP CONSTRAINT [FK_CITENSVENDA_CVENDA]
GO
IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CPGTOSVENDA_CVENDA]') AND parent_object_id = OBJECT_ID(N'[dbo].[CPGTOSVENDA]'))
ALTER TABLE [dbo].[CPGTOSVENDA] DROP CONSTRAINT [FK_CPGTOSVENDA_CVENDA]
GO

ALTER TABLE [dbo].[CITENSVENDA]  WITH CHECK 
  ADD  CONSTRAINT [FK_CITENSVENDA_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
   REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[CITENSVENDA] CHECK CONSTRAINT [FK_CITENSVENDA_CVENDA]
GO
ALTER TABLE [dbo].[CPGTOSVENDA]  WITH CHECK 
   ADD  CONSTRAINT [FK_CPGTOSVENDA_CVENDA] FOREIGN KEY([IDLOJA], [IDVENDA])
   REFERENCES [dbo].[CVENDA] ([IDLOJA], [IDVENDA])
   ON UPDATE CASCADE
GO
ALTER TABLE [dbo].[CPGTOSVENDA] CHECK CONSTRAINT [FK_CPGTOSVENDA_CVENDA]
GO

