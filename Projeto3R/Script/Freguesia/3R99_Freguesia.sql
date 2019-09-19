USE [G3R];
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE COLIGADA SET NMCOLIGADA='GRUPO 3R'
, RAZAO='3 RIOS DEPILAÇÃO COMERCIO E SERVIÇOS LTDA'
, TAG='0D44497A61757A7C7A7B0802487C60747D76760907081A06061706090D01'
WHERE IDCOLIGADA = 1;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE OLOJA SET CNPJ='11647828000152'
,NOME='FREGUESIA'
,ENDERECO='Estrada dos Três Rios, 400 Lj.D'
,BAIRRO='Freguesia-jpa'
,CIDADE='Rio de Janeiro'
,ESTADO='RJ'
,INSCEST='04536894'
,EMAIL='Freguesia_RJ@DpilBrasil.com.br'
,TELEFONE1='2427-0821'
,TELEFONE2='3268-1203'
,CEP='22745-005'
,FAX=''
,NMCONTATO='Adriane Ramos'
,CARGOCONTATO='Operadora'
,DTOPERACAO='2010-03-10'
,DIMENSAO='3,80x6,80' 
WHERE IDLOJA = 1;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OSALA on;
INSERT OSALA(IDSALA, IDLOJA,CODSALA,DIMENSAO,DTOPERACAO) VALUES (1, 1,'01','2,20x3,00','2010-06-24');
SET IDENTITY_INSERT OSALA OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OMAQUINA on;
INSERT OMAQUINA(IDMAQUINA,CODMAQUINA,NREGISTRO,NANVISA,DTOPERACAO,DISPAROS,SITMAQUINA) VALUES (1,'01','','','2010-06-24',70000, 1);
SET IDENTITY_INSERT OMAQUINA OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT OSALA_MAQUINA(IDMAQUINA,DTINICIO,DTFIM,IDSALA,IDLOJA) VALUES (1,'2010-06-24',Null,1,1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OAGENDA on;
INSERT OAGENDA(IDAGENDA,IDLOJA,IDSALA,CODAGENDA) VALUES (1,1,1,'Sala 01');
SET IDENTITY_INSERT OAGENDA OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO SESTPROD (IDEST, IDPROD, UNIDCONTROLE, SLDATUAL, SLDDISPONIVEL, SLDFECHAMENTO) VALUES (1, 2, 'pç', 100, 100, 0)
INSERT INTO SMOVEST (IDEST, IDPROD, IDLOJA, DTMOV, QTDITEM, UNIDCONTROLE, TPDOC, NUMDOC, IDFOR, ITEMDOC) VALUES (1, 2, 1, '2010-05-24', 100 , 'pç', 'NF', '000001', 1, 1)
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
UPDATE FCCORRENTE SET DSCCONTA='Itau -3Rios', NUMBANCO='341', NUMCONTA='03751', DVCONTA='7', NUMAGENCIA='7789' WHERE TPCONTA = 'B' AND EVENDA = 1;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT RFUNCIONARIO on;
INSERT RFUNCIONARIO(IDFUNCIONARIO,IDLOJA,CHAPA,NOME,DTADMISSAO,DTDEMISSAO,FLGCERTIFICADO,SITFUNC) VALUES (1,1,'000001','Adriane',convert(datetime,'2010-04-10 00:00:00.000',121),convert(datetime,NULL,121),'1','A');
INSERT RFUNCIONARIO(IDFUNCIONARIO,IDLOJA,CHAPA,NOME,DTADMISSAO,DTDEMISSAO,FLGCERTIFICADO,SITFUNC) VALUES (2,1,'000002','Fabiana',convert(datetime,'2010-04-10 00:00:00.000',121),convert(datetime,NULL,121),'1','A');
INSERT RFUNCIONARIO(IDFUNCIONARIO,IDLOJA,CHAPA,NOME,DTADMISSAO,DTDEMISSAO,FLGCERTIFICADO,SITFUNC) VALUES (3,1,'000003','Priscila',convert(datetime,'2010-05-18 00:00:00.000',121),convert(datetime,NULL,121),'1','A');
SET IDENTITY_INSERT RFUNCIONARIO OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

SET IDENTITY_INSERT OCONTATO on;
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  3, GetDate(), 0, 'Caixa');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  4, GetDate(), 0, 'CCorrente');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  5, GetDate(), 0, 'Agua Mineral');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  6, GetDate(), 0, 'Xerox');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  7, GetDate(), 0, 'Adesivos');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  8, GetDate(), 0, 'Chaveiro');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1,  9, GetDate(), 0, 'Formedics');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME) VALUES (1, 10, GetDate(), 0, 'Marcos - Pedreiro');
SET IDENTITY_INSERT OCONTATO off;
----------------------------------------------------------------------------------
--Update PARAM Set VLPARAM = 'freguesia10'					Where CODPARAM= 'MailPWD' And CODSIS= 'P3R';
--Update PARAM Set VLPARAM = 'tresrios10'					Where CODPARAM= 'FtpPWD' And CODSIS= 'P3R';
--Update PARAM Set VLPARAM = 'http://www.comtele.com.br/sms/api/api_fuse_connection.php?fuse=send_msg' Where CODPARAM= 'SMSURLBASE' And CODSIS= 'P3R';
--Update PARAM Set VLPARAM = 'MjE0'							Where CODPARAM= 'SMSUID' And CODSIS= 'P3R';
--Update PARAM Set VLPARAM = ''								Where CODPARAM= 'SMSPWD' And CODSIS= 'P3R';
--Update PARAM Set VLPARAM = 'ftp.classeanet.com.br'			Where CODPARAM= 'FTP' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '045E5346515F43504751540204'		Where CODPARAM= 'MailPWD' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '044C4146514B445C5B4B0403'		Where CODPARAM= 'FtpPWD' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '53505D4040480C1A1B4F42441A5B5B55405D59511A5B5A5A1A5A46164755451A55485C1B55485C6B524D46506B5B59595A5D56405D575B184450450B524D4550094B505D5067594B53' Where CODPARAM= 'SMSURLBASE' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '0475785E7108'					Where CODPARAM= 'SMSUID' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '6471'							Where CODPARAM= 'SMSPWD' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '465E534044165559554B46565556514C1A5B5A591A5A47'	Where CODPARAM= 'FTP' And CODSIS= 'P3R';

Update PARAM Set VLPARAM = '0'								Where CODPARAM= 'FRANQCOM' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '1'								Where CODPARAM= 'SOCIOCOM' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'pop.dpilbrasil.com.br'			Where CODPARAM= 'POP3HOST' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'smtp.dpilbrasil.com.br'		Where CODPARAM= 'SMTPHost' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '587'							Where CODPARAM= 'SMTPPort' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'Dpil Freguesia/RJ'				Where CODPARAM= 'FromDisplayName' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'freguesia_rj@dpilbrasil.com.br' Where CODPARAM= 'MailUID' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '1'								Where CODPARAM= 'UseAuthentication' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '1'								Where CODPARAM= 'UsePopAuthentication' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '/Banco/'						Where CODPARAM= 'FTPBakPath' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'freguesia'						Where CODPARAM= 'FtpUID' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'diogenes72@bol.com.br'			Where CODPARAM= 'LstEMailSocio' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '2178344618'						Where CODPARAM= 'LstCelSocio' And CODSIS= 'P3R';
