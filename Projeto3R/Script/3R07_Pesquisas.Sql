
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--USE G3R;
--SELECT * FROM GPESQUISA WHERE CODSIS='P3R';
--SELECT * FROM MODULO WHERE Substring(IDMODU,3,1)='.'
--SELECT * FROM GPESQFERR;
--SELECT * FROM GPESQPARAM;
--SELECT * FROM GPESQUISA_PESQFERR;
--SELECT * FROM GPESQUISA_PESQPARAM;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
/**************************************************************************************************
*******************************************  PESQUISAS *******************************************
**************************************************************************************************/
--select * from modulo order by id
DELETE FROM GPESQUISA_PESQFERR  --WHERE IDPESQUISA IN (SELECT P.IDPESQUISA FROM GPESQUISA P JOIN MODULO M ON P.IDMODU=M.IDMODU WHERE Substring(M.IDMODU,3,1)='.'); -- M.ID>1000 
DELETE FROM GPESQUISA_PESQPARAM --WHERE IDPESQUISA IN (SELECT P.IDPESQUISA FROM GPESQUISA P JOIN MODULO M ON P.IDMODU=M.IDMODU WHERE Substring(M.IDMODU,3,1)='.'); -- M.ID>1000 

DELETE FROM GPESQFERR  --WHERE IDFERR  IN (SELECT J.IDFERR  FROM GPESQUISA_PESQFERR  J WHERE J.IDPESQUISA IN (SELECT P.IDPESQUISA FROM GPESQUISA P JOIN MODULO M ON P.IDMODU=M.IDMODU WHERE Substring(M.IDMODU,3,1)='.')); -- M.ID>1000 ;
DELETE FROM GPESQPARAM --WHERE IDPARAM IN (SELECT J.IDPARAM FROM GPESQUISA_PESQPARAM J WHERE J.IDPESQUISA IN (SELECT P.IDPESQUISA FROM GPESQUISA P JOIN MODULO M ON P.IDMODU=M.IDMODU WHERE Substring(M.IDMODU,3,1)='.')); -- M.ID>1000 ;
DELETE FROM GPESQUISA WHERE CODSIS='P3R' And IDMODU IN (SELECT M.IDMODU FROM MODULO M WHERE Substring(M.IDMODU,3,1)='.'); -- M.ID>1000
DELETE FROM MODULO_SISTEMA WHERE Substring(IDMODU,3,1)='.'; --And ID>1000;
DELETE FROM MODULO WHERE Substring(IDMODU,3,1)='.'; --And ID>1000;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT MODULO on;
INSERT MODULO(ID,IDMODU,MODUPAI,SITMODU,MENUDEFAULT,DSCMODU) VALUES (1001,'01.001','CNSGERAL','S','N','Lista de Clientes '    );
INSERT MODULO(ID,IDMODU,MODUPAI,SITMODU,MENUDEFAULT,DSCMODU) VALUES (1002,'01.002','CNSGERAL','S','N','Clientes Faltantes'     );
INSERT MODULO(ID,IDMODU,MODUPAI,SITMODU,MENUDEFAULT,DSCMODU) VALUES (1003,'01.003','CNSGERAL','S','N','Clientes Agendados'    );
INSERT MODULO(ID,IDMODU,MODUPAI,SITMODU,MENUDEFAULT,DSCMODU) VALUES (1004,'01.004','CNSGERAL','S','N','Lista de Aniversariantes');
INSERT MODULO(ID,IDMODU,MODUPAI,SITMODU,MENUDEFAULT,DSCMODU) VALUES (2001,'02.001','CNSGERAL','S','N','Produção por Funcionário');
SET IDENTITY_INSERT MODULO off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE MODULO SET IDPAI=(SELECT	M2.ID FROM MODULO M2 WHERE M2.IDMODU=MODULO.MODUPAI) WHERE MODULO.MODUPAI='CNSGERAL'
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT MODULO_SISTEMA(IDMODU,CODSIS,INDICE,MENU,GRUPOMENU) VALUES('01.001','P3R',0,'N',0);
INSERT MODULO_SISTEMA(IDMODU,CODSIS,INDICE,MENU,GRUPOMENU) VALUES('01.002','P3R',0,'N',0);
INSERT MODULO_SISTEMA(IDMODU,CODSIS,INDICE,MENU,GRUPOMENU) VALUES('01.003','P3R',0,'N',0);
INSERT MODULO_SISTEMA(IDMODU,CODSIS,INDICE,MENU,GRUPOMENU) VALUES('01.004','P3R',0,'N',0);
INSERT MODULO_SISTEMA(IDMODU,CODSIS,INDICE,MENU,GRUPOMENU) VALUES('02.001','P3R',0,'N',0);
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE MODULO_SISTEMA SET ID   =(SELECT M2.ID FROM MODULO M2 WHERE M2.IDMODU=MODULO_SISTEMA.IDMODU);
UPDATE MODULO_SISTEMA SET MODUPAI=(SELECT M2.MODUPAI FROM MODULO M2 WHERE M2.ID=MODULO_SISTEMA.ID);
UPDATE MODULO_SISTEMA SET IDPAI=(SELECT M2.ID FROM MODULO M2 WHERE M2.IDMODU=MODULO_SISTEMA.MODUPAI);
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT GPESQUISA ON;
INSERT INTO GPESQUISA(IDPESQUISA,CODSIS,IDMODU,IDUSU) VALUES (1, 'P3R', '01.001', 'DIO');
INSERT INTO GPESQUISA(IDPESQUISA,CODSIS,IDMODU,IDUSU) VALUES (2, 'P3R', '01.002', 'DIO');
INSERT INTO GPESQUISA(IDPESQUISA,CODSIS,IDMODU,IDUSU) VALUES (3, 'P3R', '01.003', 'DIO');
INSERT INTO GPESQUISA(IDPESQUISA,CODSIS,IDMODU,IDUSU) VALUES (4, 'P3R', '01.004', 'DIO');
INSERT INTO GPESQUISA(IDPESQUISA,CODSIS,IDMODU,IDUSU) VALUES (5, 'P3R', '02.001', 'DIO');
SET IDENTITY_INSERT GPESQUISA OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE GPESQUISA SET 
TIPOPESQUISA='SISTEMA'
, PESQDEFAULT=0
, IDCONEXAO= Null
, ESCOPO=2
, NOMEPESQUISA = 'Lista de Clientes'
, DSCPESQUISA='Lista Geral de Clientes'
, PESQSQL= 'Select C.IDCLIENTE, C.NOME, C.TEL1, C.TEL2, C.EMAIL
, SITUAÇÃO = CASE C.ATIVO WHEN 0 THEN '+ Char(39)+'Inativo'+ Char(39)+' WHEN 1 THEN '+ Char(39)+'Ativo'+ Char(39)+' WHEN 2 THEN '+ Char(39)+'Em Espera'+ Char(39)+' END
From OCLIENTE C 
WHERE C.IDLOJA=@@IDLOJA
And C.ISENTO=0
And CASE C.ATIVO WHEN 0 THEN '+ Char(39)+'Inativo'+ Char(39)+' WHEN 1 THEN '+ Char(39)+'Ativo'+ Char(39)+' WHEN 2 THEN '+ Char(39)+'Em Espera'+ Char(39)+' END in (@@CLIATIVO) 
Order By C.IDCLIENTE, C.NOME '
, TAGCAMPOS = '|CAMPO=IDCLIENTE|GRUPO=0|VISIBLE=S:|CAMPO=NOME|GRUPO=0|VISIBLE=S:'
WHERE IDPESQUISA=1;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE GPESQUISA SET 
TIPOPESQUISA='SISTEMA'
, PESQDEFAULT=0
, IDCONEXAO= Null
, ESCOPO=2
, NOMEPESQUISA = 'Clientes Faltantes'
, DSCPESQUISA='Lista de clientes que não possuem próxima sessão'
, PESQSQL= 'Select Distinct E.IDCLIENTE [Id.], C.NOME [Cliente], C.TEL1, C.TEL2, C.EMAIL, Cast(Convert(char, E.STARTDATETIME, 103) as Datetime) [Data]
, '+ Char(39) + 'Avaliação sem Sessão' + Char(39)+' [Obs]
, Convert(char, E2.STARTDATETIME, 103) [Prox. Sessão]
, TR.DSCTRATAMENTO [Tratamento]
From OEVENTOAGENDA  E
JOIN OSERVICOEVT S On E.IDLOJA=S.IDLOJA AND E.IDEVENTO=S.IDEVENTO
JOIN OCLIENTE C On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE
JOIN OTPSERVICO SE On S.IDLOJA=SE.IDLOJA AND S.IDTPSERVICO=SE.IDTPSERVICO 
JOIN OTPTRATAMENTO TR On S.IDLOJA=TR.IDLOJA AND S.IDTPTRATAMENTO=TR.IDTPTRATAMENTO
LEFT JOIN OEVENTOAGENDA  E2 On E.IDLOJA=E2.IDLOJA AND E.IDCLIENTE=E2.IDCLIENTE AND E.STARTDATETIME<E2.STARTDATETIME
WHERE E.IDLOJA=@@IDLOJA
And C.ATIVO=1
AND SE.IDTPSERVICO=1
AND E.FLGCANCELADO=0
AND Month(E.STARTDATETIME) = @@Mes
AND Year(E.STARTDATETIME) = @@Ano
AND E.STARTDATETIME < GETDATE()
AND E2.STARTDATETIME IS NULL 

Union all 

Select Distinct  E.IDCLIENTE [Id.], C.NOME [Cliente], C.TEL1, C.TEL2, C.EMAIL, Cast(Convert(char, E.STARTDATETIME, 103) as Datetime) [Data]
, SE.DSCSERVICO+'+ Char(39)+' Cancelad'+ Char(39)+'+case S.IDTPSERVICO when 1 then '+ Char(39)+'o'+ Char(39)+' else '+ Char(39)+'a'+ Char(39)+' end [Obs]
, Convert(char, E2.STARTDATETIME, 103) [Prox. Sessão]
, TR.DSCTRATAMENTO [Tratamento]
From OEVENTOAGENDA  E
JOIN OSERVICOEVT S On E.IDLOJA=S.IDLOJA AND E.IDEVENTO=S.IDEVENTO
JOIN OCLIENTE C On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE
JOIN OTPSERVICO SE On S.IDLOJA=SE.IDLOJA AND S.IDTPSERVICO=SE.IDTPSERVICO 
JOIN OTPTRATAMENTO TR On S.IDLOJA=TR.IDLOJA AND S.IDTPTRATAMENTO=TR.IDTPTRATAMENTO
LEFT JOIN OEVENTOAGENDA  E2 On E.IDLOJA=E2.IDLOJA AND E.IDCLIENTE=E2.IDCLIENTE AND E2.STARTDATETIME>E.STARTDATETIME
WHERE E.IDLOJA=@@IDLOJA	
AND C.ATIVO=1
AND E.FLGCANCELADO=1
AND Month(E.STARTDATETIME) = @@Mes
AND Year(E.STARTDATETIME) = @@Ano
AND E.STARTDATETIME < GETDATE() 
AND E2.STARTDATETIME IS NULL

Union all 

Select DISTINCT A.IDCLIENTE [Id.], C.NOME [Cliente], C.Tel1, C.Tel2, C.EMAIL, Cast(Convert(char, A.DTATEND, 103) as Datetime) [Data]
, '+ Char(39)+'Atendimento Sem Remarcação'+ Char(39)+' [Obs]
, Convert(char, E.STARTDATETIME, 103) [Prox. Sessão]
, TR.DSCTRATAMENTO [Tratamento]
From OATENDIMENTO A 
JOIN OCLIENTE C ON A.IDLOJA=C.IDLOJA AND A.IDCLIENTE=C.IDCLIENTE
JOIN OSESSAO S ON A.IDLOJA=S.IDLOJA AND A.IDATENDIMENTO=S.IDATENDIMENTO AND S.IDTPSERVICO<>1
JOIN OTPTRATAMENTO TR On S.IDLOJA=TR.IDLOJA AND S.IDTPTRATAMENTO=TR.IDTPTRATAMENTO
LEFT JOIN OEVENTOAGENDA  E On E.IDLOJA=A.IDLOJA AND E.IDCLIENTE=A.IDCLIENTE
WHERE A.IDLOJA=@@IDLOJA
AND C.ATIVO=1 
AND C.ISENTO=0
AND Year(A.DTATEND) = @@Ano
AND Month(A.DTATEND) >= @@Mes
AND TR.FREQUENCIA > 0
AND E.STARTDATETIME is null
Order By Obs, Cliente, Data, Tratamento' 
, TAGCAMPOS = '|CAMPO=ID.|GRUPO=0|VISIBLE=N:|CAMPO=CLIENTE|GRUPO=0|VISIBLE=S:|CAMPO=TEL1|GRUPO=0|VISIBLE=S:|CAMPO=TEL2|GRUPO=0|VISIBLE=S:|CAMPO=EMAIL|GRUPO=0|VISIBLE=S:|CAMPO=DATA|GRUPO=0|VISIBLE=S:|CAMPO=OBS|GRUPO=1|VISIBLE=N:|CAMPO=PROX. SESSÃO|GRUPO=0|VISIBLE=N:'
WHERE IDPESQUISA=2;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE GPESQUISA SET 
TIPOPESQUISA='SISTEMA'
, PESQDEFAULT=0
, IDCONEXAO= Null
, ESCOPO=2
, NOMEPESQUISA = 'Clientes Agendados'
, DSCPESQUISA='Lista de Clientes Agendados'
, PESQSQL= 'Select A.IDCLIENTE, C.NOME, Max(A.DTATEND) [DTATEND], Max(A.IDATENDIMENTO) [IDATENDIMENTO]
, (Select Cast(Convert(char(10), Max(E.StartDateTime), 103) as Datetime)
    From OEVENTOAGENDA E
    Where E.IDLOJA=A.IDLOJA
    And E.IDCLIENTE=A.IDCLIENTE
    Group By e.idcliente) as [AGENDA]
, (Select Right('+ Char(39)+'00'+ Char(39)+'+RTrim(LTrim(Cast(DatePart(hh,Max(E.StartDateTime)) as char))),2)+'+ Char(39)+':'+ Char(39)+'+Right('+ Char(39)+'00'+ Char(39)+'+RTrim(LTrim(Cast(DatePart(mi,Max(E.StartDateTime)) as char))),2)
    From OEVENTOAGENDA E
    Where E.IDLOJA=A.IDLOJA
    And E.IDCLIENTE=A.IDCLIENTE
    Group By e.idcliente) as [HORA]
From OSESSAO S 
Join OATENDIMENTO A On A.IDATENDIMENTO=S.IDATENDIMENTO 
Join OCLIENTE C On A.IDCLIENTE=C.IDCLIENTE
WHERE S.IDLOJA=@@IDLOJA
And S.IDTPSERVICO<>1  
And C.ISENTO=0
Group By A.IDCLIENTE, C.NOME 
Having   (select max(e.StartDateTime) from oeventoagenda e where e.idloja=a.idloja and e.idcliente=a.idcliente group by e.idcliente)  >=Getdate()-1  
And max(a.dtatend)<(Select Max(E.StartDateTime) From OEVENTOAGENDA E Where E.IDLOJA=A.IDLOJa And E.IDCLIENTE=A.IDCLIENTE Group By E.IDCLIENTE)    
Order By [AGENDA], [HORA]'
, TAGCAMPOS = '|CAMPO=IDCLIENTE|GRUPO=0|VISIBLE=N:|CAMPO=NOME|GRUPO=0|VISIBLE=S:|CAMPO=DTATEND|GRUPO=0|VISIBLE=N:|CAMPO=IDATENDIMENTO|GRUPO=0|VISIBLE=N:|CAMPO=AGENDA|GRUPO=0|VISIBLE=S:'
WHERE IDPESQUISA=3;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE GPESQUISA SET 
TIPOPESQUISA='SISTEMA'
, PESQDEFAULT=0
, IDCONEXAO= Null
, ESCOPO=2
, NOMEPESQUISA = 'Lista de Aniversariantes'
, DSCPESQUISA='Lista de Aniversariantes'
, PESQSQL= 'Select Distinct C.IDCLIENTE [Id.], C.NOME [Cliente]
, Right('+ Char(39)+'00'+ Char(39)+'+RTrim(Month(C.DTNASC)),2)+'+Char(39)+'-'+Char(39)+'+DateName(Month,C.DTNASC) [Mês]
, Right('+Char(39)+'00'+Char(39)+'+RTrim(Day(C.DTNASC)),2)+'+Char(39)+'/'+Char(39)+'+Right('+ Char(39)+'00'+ Char(39)+'+RTrim(Month(C.DTNASC)),2) [Data Nasc.]
, TEL1 [Telefone]
, EMAIL [e-Mail] From OCLIENTE C 
WHERE C.IDLOJA=@@IDLOJA
And C.ATIVO=1 
Order By 3,5'
, TAGCAMPOS = '|CAMPO=ID.|GRUPO=0|VISIBLE=N:|CAMPO=CLIENTE|GRUPO=0|VISIBLE=S:|CAMPO=MÊS|GRUPO=1|VISIBLE=N:|CAMPO=DATA NASC.|GRUPO=0|VISIBLE=S:|CAMPO=TELEFONE|GRUPO=0|VISIBLE=S:|CAMPO=E-MAIL|GRUPO=0|VISIBLE=S:|'
WHERE IDPESQUISA=4;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM GPESQUISA
UPDATE GPESQUISA SET 
TIPOPESQUISA='SISTEMA'
, PESQDEFAULT=0
, IDCONEXAO= Null
, ESCOPO=2
, NOMEPESQUISA = 'Produção por Funcionário'
, DSCPESQUISA='Quantidade de Áreas por Funcionário'
, PESQSQL= '
Select F.NOME [Funcionário] 
, COUNT(distinct cast(S.IDSESSAO as varchar(10))+cast(s.IDATENDIMENTO as varchar(10))) [Produção]
From OATENDIMENTO A
Join OSESSAO S On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO
Join RFUNCIONARIO F On F.IDLOJA=A.IDLOJA AND F.IDFUNCIONARIO=A.IDFUNCIONARIO
Join OCLIENTE C On F.IDLOJA=C.IDLOJA AND C.IDCLIENTE=A.IDCLIENTE 
Join OTPSERVICO SE On SE.IDLOJA=S.IDLOJA AND SE.IDTPSERVICO=S.IDTPSERVICO
Join OLOJA L On A.IDLOJA=L.IDLOJA
WHERE A.IDLOJA=@@IDLOJA
AND YEAR(A.DTATEND)=@@Ano
AND MONTH(A.DTATEND)=@@Mes
AND SE.IDTPSERVICO>1
AND C.ISENTO=0
GROUP BY  F.NOME
ORDER BY 1'
, TAGCAMPOS = '|CAMPO=Funcionário|GRUPO=0|VISIBLE=S:|CAMPO=Produção|GRUPO=0|VISIBLE=S:|'
WHERE IDPESQUISA=5;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--select * from GPESQPARAM
--																																																	IDPARAM, CODPARAM, TITULO     , TIPO, LISTA												, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, IDCONEXAO
SET IDENTITY_INSERT GPESQPARAM ON;
IF Not Exists(Select IDPARAM from GPESQPARAM Where IDPARAM=1) INSERT INTO GPESQPARAM (IDPARAM, CODPARAM, TITULO, TIPO, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, IDCONEXAO) VALUES (1, 'IDLOJA'  , 'Unidade'      , 'N', 'SELECT IDLOJA [ID], NOME [Unidade] FROM OLOJA WHERE ATIVO=1', 'N', 'N', 'S', 1, 1, 0);
IF Not Exists(Select IDPARAM from GPESQPARAM Where IDPARAM=2) INSERT INTO GPESQPARAM (IDPARAM, CODPARAM, TITULO, TIPO, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, IDCONEXAO) VALUES (2, 'Ano'     , 'Ano'          , 'N', 'SELECT DISTINCT YEAR(DTATEND) FROM OATENDIMENTO ORDER BY 1', 'N', 'N', 'S', 1, 1, 0)
IF Not Exists(Select IDPARAM from GPESQPARAM Where IDPARAM=3) INSERT INTO GPESQPARAM (IDPARAM, CODPARAM, TITULO, TIPO, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, IDCONEXAO) VALUES (3, 'Mes'     , 'Mês'          , 'N', 'SELECT DISTINCT Right('+Char(39)+'00'+Char(39)+'+rtrim(cast(month(DTATEND) as char(2))),2), DateName(MONTH, DTATEND) FROM OATENDIMENTO WHERE YEAR(DTATEND)=@@Ano ORDER BY 1', 'N', 'N', 'S', 1, 1, 0);
IF Not Exists(Select IDPARAM from GPESQPARAM Where IDPARAM=4) INSERT INTO GPESQPARAM (IDPARAM, CODPARAM, TITULO, TIPO, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, IDCONEXAO) VALUES (4, 'CLIATIVO', 'Situação'     , 'S', 'Ativo|Inativo|Em Espera|', 'S', 'N', 'N', 1, 3, 0);
--IF Not Exists(Select IDPARAM from GPESQPARAM Where IDPARAM=5) INSERT INTO GPESQPARAM (IDPARAM, CODPARAM, TITULO, TIPO, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, IDCONEXAO) VALUES (5, 'IDTPTRATAMENTO', 'Tratamento' , 'N', 'SELECT DISTINCT IDTPTRATAMENTO [ID], DSCTRATAMENTO [Tratamento] FROM OTPTRATAMENTO WHERE IDLOJA=(Select Min(IDLOJA) From OLOJA) Or IDLOJA in ( @@IDLOJA )', 'S', 'N', 'S', 2, 3, 0);
SET IDENTITY_INSERT GPESQPARAM OFF;
--select * from GPESQUISA_PESQPARAM
INSERT INTO GPESQUISA_PESQPARAM  SELECT 1, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=1;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 1, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=4;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 2, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=1;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 2, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=2;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 2, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 2 FROM GPESQPARAM WHERE IDPARAM=3;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 3, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=1;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 4, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=1;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 5, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 1 FROM GPESQPARAM WHERE IDPARAM=1;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 5, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 2 FROM GPESQPARAM WHERE IDPARAM=2;
INSERT INTO GPESQUISA_PESQPARAM  SELECT 5, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 3 FROM GPESQPARAM WHERE IDPARAM=3;
--INSERT INTO GPESQUISA_PESQPARAM  SELECT 5, IDPARAM, LISTA, SELMULT, LOV, OBRIGATORIO, SCALEWIDTH, SCALEHEIGHT, 2 FROM GPESQPARAM WHERE IDPARAM=5;

---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT GPESQFERR ON;
IF Not Exists(Select IDFERR from GPESQFERR Where IDFERR=1) INSERT INTO GPESQFERR (IDFERR, TITULO, DESCRICAO, SCRIPT) VALUES (1, 'Clientes', 'Cadastro de Clientes', '');
SET IDENTITY_INSERT GPESQFERR OFF;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO GPESQUISA_PESQFERR (IDPESQUISA, IDFERR) VALUES (1, 1)
INSERT INTO GPESQUISA_PESQFERR (IDPESQUISA, IDFERR) VALUES (2, 1)
INSERT INTO GPESQUISA_PESQFERR (IDPESQUISA, IDFERR) VALUES (3, 1)
INSERT INTO GPESQUISA_PESQFERR (IDPESQUISA, IDFERR) VALUES (4, 1)
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Update  GPESQFERR Set SCRIPT='Dim MyCli
 Set MyCli=CreateObject("Contato3R.TL_CadCliente")  
 Set Mycli.Sys=mvarSys
 Mycli.IDCLIENTE=mvarme.GrdDados.SelectedRows(0).Record(0).Value 
 MyCLi.Show'
 where IDFERR = 1;
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
