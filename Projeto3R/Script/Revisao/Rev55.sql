--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='55', ARQATU='Rev55.sql';
/****************************************************************************
****************************************************************************/

UPDATE OTPTRATAMENTO SET FLGDISPARO=1 WHERE IDTPTRATAMENTO<=4
go

IF EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='DTEMISSAO') EXEC sp_rename 'CPROMOCAO.DTEMISSAO', 'DTINI', 'COLUMN'
go
IF EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='DTVENC') EXEC sp_rename 'CPROMOCAO.DTVENC', 'DTFIM', 'COLUMN'
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='FLGSERV') ALTER TABLE CPROMOCAO ADD FLGSERV [int] NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='FLGTRAT') ALTER TABLE CPROMOCAO ADD FLGTRAT [int] NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='FLGAREA') ALTER TABLE CPROMOCAO ADD FLGAREA [int] NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='SERVIN') ALTER TABLE CPROMOCAO ADD SERVIN varchar(50) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='TRATIN') ALTER TABLE CPROMOCAO ADD TRATIN varchar(50) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CPROMOCAO' AND C.NAME='AREAIN') ALTER TABLE CPROMOCAO ADD AREAIN varchar(50) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDPACOTE') ALTER TABLE CCUPOM ADD IDPACOTE int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDATENDIMENTO') ALTER TABLE CCUPOM ADD IDATENDIMENTO int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDSESSAO') ALTER TABLE CCUPOM ADD IDSESSAO int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDVENDA') ALTER TABLE CCUPOM ADD IDVENDA int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDCLIENTE') ALTER TABLE CCUPOM ADD IDCLIENTE int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDTPSERVICO') ALTER TABLE CCUPOM ADD IDTPSERVICO int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDTPTRATAMENTO') ALTER TABLE CCUPOM ADD IDTPTRATAMENTO int Null
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='CCUPOM' AND C.NAME='IDAREA') ALTER TABLE CCUPOM ADD IDAREA int Null
go


exec sp_bindefault DF_0, 'CPROMOCAO.FLGSERV'
go
exec sp_bindefault DF_0, 'CPROMOCAO.FLGTRAT'
go
exec sp_bindefault DF_0, 'CPROMOCAO.FLGAREA'
GO

CREATE TABLE CPACOTE
(	IDLOJA  int  NOT NULL ,
	IDPROMO  int  NOT NULL ,
	IDPACOTE  integer  NOT NULL,
	IDCLIENTE  int  NULL ,
	DTEMISSAO  datetime  NULL ,
	VALOR  decimal(9,2)  NULL ,
	VLDESC  decimal(9,2)  NULL ,	 
	ALTERSTAMP  integer  NULL ,
	TIMESTAMP  datetime  NULL
)
go
ALTER TABLE CPACOTE	ADD CONSTRAINT  PK_CPACOTE PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,IDPROMO  ASC,IDPACOTE  ASC)
go
exec sp_bindefault DF_1, 'CPACOTE.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'CPACOTE.TIMESTAMP'
go

ALTER TABLE CPACOTE	ADD CONSTRAINT  R_206 FOREIGN KEY (IDLOJA,IDPROMO) REFERENCES CPROMOCAO(IDLOJA,IDPROMO)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

ALTER TABLE CPACOTE	ADD CONSTRAINT  R_205 FOREIGN KEY (IDLOJA,IDCLIENTE) REFERENCES OCLIENTE(IDLOJA,IDCLIENTE)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OEVENTOAGENDA' AND C.NAME='IDSALA') ALTER TABLE OEVENTOAGENDA ADD IDSALA [int] NULL
go
UPDATE OEVENTOAGENDA SET IDSALA= Right(SCHEDULEID ,1), IDAGENDA=1
GO
UPDATE OEVENTOAGENDA SET IDAGENDA=1
GO
UPDATE OEVENTOAGENDA SET SCHEDULEID= IDLOJA*1000+Right(SCHEDULEID ,1)
GO