
/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='34', ARQATU='Rev34.sql';
/****************************************************************************
****************************************************************************/

UPDATE OSESSAO SET IDMANIPULO =1 WHERE IDMANIPULO =0
GO

CREATE TABLE FITEMFATURA
(   IDLOJA  int  NOT NULL ,
	IDFATURA  int  NOT NULL ,
	IDITEM  integer  NOT NULL ,
	IDPROD  int  NULL ,
	UNIDCONTROLE  varchar(5)  NULL ,
	QTDITEM  decimal(9,2)  NULL ,
	VLUNIT decimal(9, 2) NULL,
	ALTERSTAMP  integer  NULL ,
	TIMESTAMP  datetime  NULL)
go
ALTER TABLE FITEMFATURA ADD CONSTRAINT PK_FITEMFATURA PRIMARY KEY NONCLUSTERED (IDLOJA  ASC,IDFATURA  ASC,IDITEM  ASC)
go
exec sp_bindefault DF_1, 'FITEMFATURA.ALTERSTAMP'
go

exec sp_bindefault DF_Now, 'FITEMFATURA.TIMESTAMP'
go

ALTER TABLE FITEMFATURA
	ADD CONSTRAINT  R_156 FOREIGN KEY (IDLOJA,IDFATURA) REFERENCES FFATURA(IDLOJA,IDFATURA)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

ALTER TABLE FITEMFATURA
	ADD CONSTRAINT  R_157 FOREIGN KEY (IDPROD) REFERENCES SPRODUTO(IDPROD)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

CREATE TABLE RBATIDA
(	IDLOJA  int  NOT NULL ,
	IDFUNCIONARIO  int  NOT NULL ,
	IDBATIDA  int  NOT NULL ,
	DTBATIDA  datetime  NOT NULL ,
	SENTIDO  int  NOT NULL ,
	ALTERSTAMP  integer  NULL ,
	TIMESTAMP  datetime  NULL )
go

ALTER TABLE RBATIDA	ADD CONSTRAINT  PK_RBATIDA PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,IDFUNCIONARIO  ASC,IDBATIDA  ASC)
go
exec sp_bindefault DF_1, 'RBATIDA.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'RBATIDA.TIMESTAMP'
go
exec sp_bindefault DF_1, 'RBATIDA.SENTIDO'
go
exec sp_bindefault DF_Now, 'RBATIDA.DTBATIDA'
go
ALTER TABLE RBATIDA
	ADD CONSTRAINT  R_159 FOREIGN KEY (IDLOJA,IDFUNCIONARIO) REFERENCES RFUNCIONARIO(IDLOJA,IDFUNCIONARIO)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

ALTER TABLE OTRATAMENTOCLI ADD OBS ntext NULL
GO
ALTER TABLE FFATURA ADD VLDESC decimal(9,2)  NULL 
go

ALTER TABLE FITEMFATURA ADD VLUNIT decimal(9,2)  NULL 
go

