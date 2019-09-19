/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='28', ARQATU='Rev28.sql';
/****************************************************************************
****************************************************************************/

ALTER TABLE OMAQDISPAROS ADD IDTPMANIPULO int
go


ALTER TABLE OMAQDISPAROS
	ADD CONSTRAINT  R_OTPMANIPULO_DISPARO FOREIGN KEY (IDTPMANIPULO) REFERENCES OTPMANIPULO(IDTPMANIPULO)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go
exec sp_bindefault DF_1, 'OMAQDISPAROS.IDTPMANIPULO'
go
UPDATE OMAQDISPAROS SET IDTPMANIPULO=1
go
CREATE TABLE ODIARIO
(	IDLOJA  int  NOT NULL ,
	DTDIARIO  datetime  NOT NULL ,
	DSCDIARIO  text  NULL ,
	ALTERSTAMP  integer  NULL ,
	TIMESTAMP  datetime  NULL)
go
ALTER TABLE ODIARIO	ADD CONSTRAINT  PK_ODIARIO PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,DTDIARIO  ASC)
go
exec sp_bindefault DF_1, 'ODIARIO.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'ODIARIO.TIMESTAMP'
go

ALTER TABLE ODIARIO	ADD CONSTRAINT  R_OLOJA_DIARIO FOREIGN KEY (IDLOJA) REFERENCES OLOJA(IDLOJA)
		ON DELETE NO ACTION
		ON UPDATE NO ACTION
go
exec sp_bindefault DF_1, 'OMAQUINA.SITMAQUINA'
go
update OMAQUINA set SITMAQUINA=1
go
