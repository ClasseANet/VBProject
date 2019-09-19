--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='64', ARQATU='Rev64.sql';
/****************************************************************************
****************************************************************************/
/*
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='RBATIDA' AND C.NAME='IDMOVHH')  ALTER TABLE RBATIDA ADD IDMOVHH int NULL
go

exec sp_bindefault DF_0, 'OEVENTOAGENDA.FLGAVALIACAO'
go

Update RFUNCIONARIO SET FLGVENDA=1 where FLGVENDA is Null
Go

ALTER TABLE RBATIDA
	ADD CONSTRAINT  R_159 FOREIGN KEY (IDLOJA,IDFUNCIONARIO) REFERENCES RFUNCIONARIO(IDLOJA,IDFUNCIONARIO)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go
Update RBATIDA Set FLGMANUAL=0 Where FLGMANUAL is Null
go
*/
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTPTRATAMENTO' AND C.NAME='PESOCOMISSAO')  ALTER TABLE OTPTRATAMENTO ADD PESOCOMISSAO decimal(4,2) NULL
go
exec sp_bindefault DF_1, 'OTPTRATAMENTO.PESOCOMISSAO'
go
Update OTPTRATAMENTO SET PESOCOMISSAO=1 where PESOCOMISSAO is Null
Go
IF EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='FRECIBO' AND C.NAME='DSCVERV')  ALTER TABLE FRECIBO drop column DSCVERV 
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='FRECIBO' AND C.NAME='DSCSERV')  ALTER TABLE FRECIBO ADD DSCSERV varchar(50) NULL
go
Update FRECIBO SET DSCSERV='SERVI�O DE EST�TICA' where DSCSERV is Null
Go
