/****************************************************************************
****************************************************************************/
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='25', ARQATU='Rev25.sql';
/****************************************************************************
****************************************************************************/

CREATE TABLE PMETA
(  IDLOJA  int  NOT NULL, 
   IDMETA  int  NOT NULL ,
   FAIXA1  numeric(9,2)  NULL ,
   FAIXA2  numeric(9,2)  NULL ,
   FAIXA3  numeric(9,2)  NULL ,
   ALTERSTAMP  integer  NULL ,
   TIMESTAMP  datetime  NULL
	)
go

ALTER TABLE PMETA ADD CONSTRAINT  PK_PMETA PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,IDMETA  ASC)
go
exec sp_bindefault DF_1, 'PMETA.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'PMETA.TIMESTAMP'
go

CREATE TABLE PMETAITEM
(  IDLOJA  int  NOT NULL ,
   IDMETA  int  NOT NULL ,
   DTITEM  datetime  NOT NULL ,		
   VLPREV  decimal(9,2)  NULL ,
   ALTERSTAMP  integer  NULL ,
   TIMESTAMP  datetime  NULL)
go
ALTER TABLE PMETAITEM ADD CONSTRAINT PK_PMETAITEM PRIMARY KEY   NONCLUSTERED (IDLOJA  ASC,IDMETA  ASC,DTITEM  ASC)
go
exec sp_bindefault DF_1, 'PMETAITEM.ALTERSTAMP'
go
exec sp_bindefault DF_Now, 'PMETAITEM.TIMESTAMP'
go

ALTER TABLE PMETAITEM ADD CONSTRAINT  R_PMETA FOREIGN KEY (IDLOJA,IDMETA) REFERENCES PMETA(IDLOJA,IDMETA) 
   ON DELETE NO ACTION
   ON UPDATE NO ACTION
go
 