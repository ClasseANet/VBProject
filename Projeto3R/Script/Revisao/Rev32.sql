/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='32', ARQATU='Rev32.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE OTPMAQ ADD TPMANIPULO Int NULL
go
ALTER TABLE CFORMAPGTO ADD TXSERV DECIMAL(4,2)  NULL
go
UPDATE OTPMAQ SET TPMANIPULO=1
GO
UPDATE OTPMAQ SET TPMANIPULO=0 WHERE IDTPMAQ=1
GO

CREATE TABLE FCONTAREC
(	IDLOJA  varchar(50)  NOT NULL ,
	IDCONTA  int  NOT NULL ,
	IDCLIENTE  int  NULL ,
	DTPREV  datetime  NULL ,
	VALOR  decimal(9,2)  NULL ,
	IDSUBDESP  int  NULL ,
	HISTORICO  varchar(80)  NULL ,
	IDATENDIMENTO  int  NULL ,
	IDVENDA  int  NULL ,
	SITCONTA  int  NULL ,
	IDDESP  int  NULL,
	ALTERSTAMP  integer  NULL ,
	TIMESTAMP  datetime  NULL 
)
go
ALTER TABLE FCONTAREC
	ADD CONSTRAINT  PK_FCONTAREC PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,IDCONTA  ASC)
go
exec sp_bindefault DF_1, 'FCONTAREC.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'FCONTAREC.TIMESTAMP'
go
exec sp_bindefault DF_0, 'FCONTAREC.SITCONTA'
go

ALTER TABLE FLAN ADD SALDO decimal(9,2) NULL
go
ALTER TABLE OFORMAPGTO ADD TXSERV decimal(5,2) NULL
go