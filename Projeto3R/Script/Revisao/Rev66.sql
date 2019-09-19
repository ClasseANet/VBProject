--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='66', ARQATU='Rev66.sql';
/****************************************************************************
****************************************************************************/

IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CFORMAPGTO' AND C.NAME='TXPARC')  ALTER TABLE CFORMAPGTO ADD TXPARC decimal(4,2) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBATIDA' AND C.NAME='IDMOVHH')  ALTER TABLE RBATIDA ADD IDMOVHH int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBATIDA' AND C.NAME='FLGMANUAL')  ALTER TABLE RBATIDA ADD FLGMANUAL int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBATIDA' AND C.NAME='OFOTO')  ALTER TABLE RBATIDA ADD OFOTO	binary NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBANCOHH' AND C.NAME='ACUMULADO')  ALTER TABLE RBANCOHH ADD ACUMULADO decimal(9,2) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBANCOHH' AND C.NAME='FLGFALTA')  ALTER TABLE RBANCOHH ADD FLGFALTA int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBANCOHH' AND C.NAME='OBS')  ALTER TABLE RBANCOHH ADD OBS varchar(100) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBANCOHH' AND C.NAME='FLGZERASALDO')  ALTER TABLE RBANCOHH ADD FLGZERASALDO int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBANCOHH' AND C.NAME='ACUMULADO')  ALTER TABLE RBANCOHH ADD ACUMULADO decimal(9,2)
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBANCOHH' AND C.NAME='APROVADO')  ALTER TABLE RBANCOHH ADD APROVADO int
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RFUNCIONARIO' AND C.NAME='DIAFOLGA')  ALTER TABLE RFUNCIONARIO ADD DIAFOLGA int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTPTRATAMENTO' AND C.NAME='FLGAREA')  ALTER TABLE OTPTRATAMENTO ADD FLGAREA int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTPTRATAMENTO' AND C.NAME='FLGAVALIACAO')  ALTER TABLE OTPTRATAMENTO ADD FLGAVALIACAO int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTPTRATAMENTO' AND C.NAME='FLGSESSAO')  ALTER TABLE OTPTRATAMENTO ADD FLGSESSAO int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OEVENTOAGENDA' AND C.NAME='IDSALA')  ALTER TABLE OEVENTOAGENDA ADD IDSALA int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OEVENTOAGENDA' AND C.NAME='FLGAVALIACAO')  ALTER TABLE OEVENTOAGENDA ADD FLGAVALIACAO int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RFUNCIONARIO' AND C.NAME='FLGVENDA')  ALTER TABLE RFUNCIONARIO ADD FLGVENDA int NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CVENDA' AND C.NAME='OBS')  ALTER TABLE CVENDA ADD OBS ntext NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OCLIENTE' AND C.NAME='OPER1')  ALTER TABLE OCLIENTE ADD OPER1 varchar(10) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OCLIENTE' AND C.NAME='OPER2')  ALTER TABLE OCLIENTE ADD OPER2 varchar(10) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OCLIENTE' AND C.NAME='OPERF')  ALTER TABLE OCLIENTE ADD OPERF varchar(10) NULL
go

exec sp_bindefault DF_0, 'OEVENTOAGENDA.FLGAVALIACAO'
go
exec sp_bindefault DF_0, 'OTPTRATAMENTO.FLGAREA'
go
exec sp_bindefault DF_0, 'OTPTRATAMENTO.FLGAVALIACAO'
go
exec sp_bindefault DF_1, 'OTPTRATAMENTO.FLGSESSAO'
go
exec sp_bindefault DF_0, 'RBANCOHH.FLGFALTA'
go
exec sp_bindefault DF_0, 'RBANCOHH.APROVADO'
go
exec sp_bindefault DF_0, 'RBANCOHH.FLGZERASALDO'
go
exec sp_bindefault DF_0, 'RBATIDA.FLGMANUAL'
go
exec sp_bindefault DF_0, 'RFUNCIONARIO.DIAFOLGA'
go
exec sp_bindefault DF_0, 'RFUNCIONARIO.FLGCERTIFICADO'
go
exec sp_bindefault DF_0, 'RFUNCIONARIO.FLGVENDA'
go

Update RFUNCIONARIO SET FLGVENDA=1 where FLGVENDA is Null
Go

IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_159]') AND parent_object_id = OBJECT_ID(N'[dbo].[RBATIDA]'))	ALTER TABLE [dbo].[RBATIDA] DROP CONSTRAINT [R_159]
GO

ALTER TABLE RBATIDA
	ADD CONSTRAINT  R_159 FOREIGN KEY (IDLOJA,IDFUNCIONARIO) REFERENCES RFUNCIONARIO(IDLOJA,IDFUNCIONARIO)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go
Update RBATIDA Set FLGMANUAL=0 Where FLGMANUAL is Null
go
Update RFUNCIONARIO Set DIAFOLGA=0 Where DIAFOLGA is Null
go
Update RBANCOHH Set FLGFALTA=0 Where FLGFALTA is Null
go
Update RBANCOHH Set FLGZERASALDO=0 Where FLGZERASALDO is Null
go
Update RBANCOHH Set APROVADO=0 Where APROVADO is Null
go
Update OTPTRATAMENTO Set FLGAREA=0 Where FLGAREA is Null
go
Update OTPTRATAMENTO Set FLGAREA=1 Where IDTPTRATAMENTO<=4 AND FLGAREA<>1
go
Update OTPTRATAMENTO Set FLGAVALIACAO=0 Where FLGAVALIACAO is Null
go
Update OTPTRATAMENTO Set FLGSESSAO=1 Where FLGSESSAO is Null
go
Update OTPTRATAMENTO Set FLGAVALIACAO=1 Where IDTPTRATAMENTO<=4 and FLGAVALIACAO<>1 
go
ALTER TABLE OSALA ALTER COLUMN CODSALA VARCHAR(10)
GO
Update OCLIENTE Set OPER1='' Where OPER1 is Null
go
Update OCLIENTE Set OPER2='' Where OPER2 is Null
go
Update OCLIENTE Set OPERF='' Where OPERF is Null
go
CREATE UNIQUE NONCLUSTERED INDEX [IX_RBANCOHH] ON [dbo].[RBANCOHH] ([IDLOJA] ASC,[IDFUNCIONARIO] ASC,[DTPONTO] ASC) WITH (PAD_INDEX  = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]
go