--USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='74', ARQATU='Rev74.sql';
/****************************************************************************
****************************************************************************/
ALTER TABLE CPROMO_PROD ALTER COLUMN NMPROD varchar(50) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTIPOTAREFA' AND C.NAME='EMAILKEY')  ALTER TABLE OTIPOTAREFA ADD EMAILKEY varchar(20) NULL
go
IF NOT EXISTS (SELECT * FROM sys.COLUMNS C JOIN  sys.OBJECTS O ON C.OBJECT_ID=O.OBJECT_ID WHERE O.NAME='OTIPOTAREFA' AND C.NAME='EMAILTIT')  ALTER TABLE OTIPOTAREFA ADD EMAILTIT varchar(50) NULL
go
Update OTIPOTAREFA Set EMAILKEY='Bemvindo'		, EMAILTIT='Boas Vindas'	Where IDTPTAREFA=1 
go
Update OTIPOTAREFA Set EMAILKEY='Confirmacao'	, EMAILTIT='Confirmação'	Where IDTPTAREFA=2 
go
Update OTIPOTAREFA Set EMAILKEY='Recomendacao'	, EMAILTIT='Recomendações' Where IDTPTAREFA=3 
go
Update OTIPOTAREFA Set EMAILKEY='Bemvindo'		, EMAILTIT='Boas Vindas'	Where IDTPTAREFA=4 
go
Update OTIPOTAREFA Set EMAILKEY='Remarcacao'	, EMAILTIT='Remarcação'	Where IDTPTAREFA=5 
go
Update OTIPOTAREFA Set EMAILKEY='Lembrete'		, EMAILTIT='Lembrete'		Where IDTPTAREFA=6 
go
Update OTIPOTAREFA Set EMAILKEY='Bemvindo'		, EMAILTIT='La Korpo-Boas Vindas'	Where IDTPTAREFA=1 And IDLOJA<=3
go
Update OTIPOTAREFA Set EMAILKEY='Confirmacao'	, EMAILTIT='La Korpo-Confirmação'	Where IDTPTAREFA=2 And IDLOJA<=3
go
Update OTIPOTAREFA Set EMAILKEY='Recomendacao'	, EMAILTIT='La Korpo-Recomendações' Where IDTPTAREFA=3 And IDLOJA<=3
go
Update OTIPOTAREFA Set EMAILKEY='Bemvindo'		, EMAILTIT='La Korpo-Boas Vindas'	Where IDTPTAREFA=4 And IDLOJA<=3
go
Update OTIPOTAREFA Set EMAILKEY='Remarcacao'	, EMAILTIT='La Korpo-Remarcação'	Where IDTPTAREFA=5 And IDLOJA<=3
go
Update OTIPOTAREFA Set EMAILKEY='Lembrete'		, EMAILTIT='La Korpo-Lembrete'		Where IDTPTAREFA=6 And IDLOJA<=3
go