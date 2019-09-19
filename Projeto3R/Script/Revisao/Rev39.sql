
/****************************************************************************
****************************************************************************/
--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='39', ARQATU='Rev39.sql';
/****************************************************************************
****************************************************************************/

alter table OTPTRATAMENTO ALTER COLUMN DSCTRATAMENTO VARCHAR(30)
go
alter table OTPSERVICO ALTER COLUMN DSCSERVICO VARCHAR(30)
go
alter table OSALA ADD ATIVO int
go
exec sp_bindefault DF_1, 'OSALA.ATIVO'
go
Update OSALA set ATIVO=1
go
Update OEVENTOAGENDA set ScheduleID=1 Where ScheduleID=0
go
ALTER TABLE CPROMOCAO ADD IDPROD INT NULL
go
ALTER TABLE CPROMOCAO ADD QTDPROD DECIMAL(9,2) NULL
go




