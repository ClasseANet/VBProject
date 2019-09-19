/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='33', ARQATU='Rev33.sql';
/****************************************************************************
****************************************************************************/

DROP TABLE FCONTAREC
GO

CREATE TABLE FFATURA
(	IDLOJA int  NOT NULL ,
	IDFATURA int  NOT NULL ,
	IDCLIENTE int  NULL ,
	DTEMISSAO datetime  NULL ,
	DTPREV datetime  NULL ,
	VALOR decimal(9,2)  NULL ,
	IDSUBDESP int  NULL ,
	HISTORICO varchar(80)  NULL ,
	IDATENDIMENTO int  NULL ,
	IDVENDA int  NULL ,
	SITFATURA int  NULL ,
	IDDESP int  NULL,
	ALTERSTAMP integer  NULL ,
	TIMESTAMP datetime  NULL 
)
go
ALTER TABLE FFATURA ADD CONSTRAINT  PK_FFATURA PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,IDFATURA ASC)
go

ALTER TABLE FFATURA	ADD CONSTRAINT  FK_OLOJA_FFATURA FOREIGN KEY (IDLOJA) REFERENCES OLOJA(IDLOJA)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go

exec sp_bindefault DF_1, 'FFATURA.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'FFATURA.TIMESTAMP'
go
exec sp_bindefault DF_0, 'FFATURA.SITFATURA'
go
