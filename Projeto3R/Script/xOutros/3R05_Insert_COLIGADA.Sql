USE [G3R];
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT USUARIO(IDUSU,IDGRUPO,SENHAUSU,NMUSU,DTCADASTRO,DTEXPIRAUSU,DTEXPIRASENHA,TPUSU,SITUSU,USUCADASTRO,FEXPIRASENHA,MATRUSU,ICORREIO,ECORREIO,ASSINATURA,TIMEOUT,IDSETOR,IDEMPR,LOCALTRAB,NIVELOS,NMAD,IDPESSOA,IDCOLIGADA,DTVALIDAUSU,TROCARSENHA) 
VALUES('DIO','SISTEMA','7A7C717B78687E7C7A','Diogenes S. Ramos','2010-03-08', NULL, NULL,'','A', 'DIO', 0, '', '', 'suporte@classeanet.com.br', NULL,	0, 0, NULL,	NULL, 0, '', 0,	NULL, NULL, 0);

INSERT USUARIO(IDUSU,IDGRUPO,SENHAUSU,NMUSU,DTCADASTRO,DTEXPIRAUSU,DTEXPIRASENHA,TPUSU,SITUSU,USUCADASTRO,FEXPIRASENHA,MATRUSU,ICORREIO,ECORREIO,ASSINATURA,TIMEOUT,IDSETOR,IDEMPR,LOCALTRAB,NIVELOS,NMAD,IDPESSOA,IDCOLIGADA,DTVALIDAUSU,TROCARSENHA) 
VALUES('DPIL','LOJA','','Usuário Loja','2010-03-08', NULL, NULL,'','A', 'DIO', 0, '', '', '', NULL,	0, 0, NULL,	NULL, 0, '', 0,	NULL, NULL, 0)
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into OTPSERVICO (IDTPSERVICO, DSCSERVICO) Values (1,'Avaliação');
Insert into OTPSERVICO (IDTPSERVICO, DSCSERVICO) Values (2,'Sessão');
Insert into OTPSERVICO (IDTPSERVICO, DSCSERVICO) Values (3,'Manutenção');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into OTPTRATAMENTO (IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (1,'Depilação',30);
Insert into OTPTRATAMENTO (IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (2,'Rejuvenescimento',21);
Insert into OTPTRATAMENTO (IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (3,'Mancha',15);
Insert into OTPTRATAMENTO (IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (4,'Acne',3);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into OAREA (IDAREA, DSCAREA) Values (1,'Abdômen');
Insert into OAREA (IDAREA, DSCAREA) Values (2,'Aréola Mamaria');
Insert into OAREA (IDAREA, DSCAREA) Values (3,'Axilas');
Insert into OAREA (IDAREA, DSCAREA) Values (4,'Braço Direito');
Insert into OAREA (IDAREA, DSCAREA) Values (5,'Braço Esquerdo');
Insert into OAREA (IDAREA, DSCAREA) Values (6,'Dorso');
Insert into OAREA (IDAREA, DSCAREA) Values (7,'Face');
Insert into OAREA (IDAREA, DSCAREA) Values (8,'Glúteos');
Insert into OAREA (IDAREA, DSCAREA) Values (9,'Lombar');
Insert into OAREA (IDAREA, DSCAREA) Values (10,'Mãos');
Insert into OAREA (IDAREA, DSCAREA) Values (11,'Nuca');
Insert into OAREA (IDAREA, DSCAREA) Values (12,'Ombros');
Insert into OAREA (IDAREA, DSCAREA) Values (13,'Peito');
Insert into OAREA (IDAREA, DSCAREA) Values (14,'Pés');
Insert into OAREA (IDAREA, DSCAREA) Values (15,'Perianal');
Insert into OAREA (IDAREA, DSCAREA) Values (16,'Perna Direita');
Insert into OAREA (IDAREA, DSCAREA) Values (17,'Perna Esquerda');
Insert into OAREA (IDAREA, DSCAREA) Values (18,'Pescoço(Frontal)');
Insert into OAREA (IDAREA, DSCAREA) Values (19,'Virilhas');
Insert into OAREA (IDAREA, DSCAREA) Values (20,'Meia-Perna');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO ODIRECAO (IDDIRECAO, DSCDIRECAO, NORDEM) VALUES (1, 'Horizontal', 1);
INSERT INTO ODIRECAO (IDDIRECAO, DSCDIRECAO, NORDEM) VALUES (2, 'Vertical', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (1, 'Branca +Clara', 'Muitas', 'Loiro', 'Azuis', 'Céltico', 'Muito severas e dolorosas', 'Nenhum, vermelhidão');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (2, 'Branca -Clara', 'Raras Vezes', 'Loiro Escuro/Ruivo', 'Azuis/Cinzentados', 'Europeu(12%)', 'Sempre severas e muito dolorosas', 'Raramente');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (3, 'Avermelhada', 'Muitas', 'Castanho', 'Castanhos', 'Europeu(78%)', 'Menos frequentes, ligeiras ', 'Médio');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (4, 'Castanho Claro', 'Muitas', 'Castanho Escuro', 'Escuros', 'Mediterrâneo', 'Não', 'Rápido e profundo');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (5, 'Castanho Escuro', 'Muitas', 'Preto', 'Escuros', 'Raças Magreb/Índias', 'Não', 'Rápido e profundo');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (6, 'Negra', 'Muitas', 'Preto/Grisalho', 'Negros', 'Negras', 'Não', 'Rápido e profundo');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT OREACAOADV (IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES ('1','Edema','0',convert(datetime,'2010-07-02 22:29:04.890',121));
INSERT OREACAOADV (IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES ('2','Hiper/Hipopigmentação','0',convert(datetime,'2010-07-02 22:29:54.263',121));
INSERT OREACAOADV (IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES ('3','Casca','0',convert(datetime,'2010-07-02 22:30:04.200',121));
INSERT OREACAOADV (IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES ('4','Bolhas','0',convert(datetime,'2010-07-02 22:30:09.700',121));
INSERT OREACAOADV (IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES ('5','Desepetelização','0',convert(datetime,'2010-07-02 22:30:19.983',121));
INSERT OREACAOADV (IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES ('6','Pelo Encravado','0',convert(datetime,'2010-07-02 22:35:00.687',121));
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into CCONDPGTO (IDCONDPGTO, CONDPGTO, ATIVO) Values (1,'À VISTA', 1);
Insert into CCONDPGTO (IDCONDPGTO, CONDPGTO, ATIVO) Values (2,'30 DDL', 1);
Insert into CCONDPGTO (IDCONDPGTO, CONDPGTO, ATIVO) Values (3,'30/60 DDL', 1);
Insert into CCONDPGTO (IDCONDPGTO, CONDPGTO, ATIVO) Values (4,'30/60/90 DDL', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into CFORMAPGTO (IDFORMAPGTO, FORMAPGTO, ATIVO) Values (1,'Dinheiro', 1);
Insert into CFORMAPGTO (IDFORMAPGTO, FORMAPGTO, ATIVO) Values (2,'Débito', 1);
Insert into CFORMAPGTO (IDFORMAPGTO, FORMAPGTO, ATIVO) Values (3,'Crédito', 1);
Insert into CFORMAPGTO (IDFORMAPGTO, FORMAPGTO, ATIVO) Values (4,'Cheque', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert INTO SPRODUTO (IDPROD, IDLOJA, CODPROD, NMPROD, DSCPROD, UNIDCONTROLE, UNIDCOMPRA, UNIDVENDA, VLULTCOMPRA, VLMEDIO, VLVENDA, QTDMIN, QTDMAX, QTDCOMPRA, EMATPRIMA, EPRODACABADO, EVENDA, ESERVICO, SITPROD, IDUSU) Values (1, 0, '1', 'Serviço Único','Serviço Único', 'un', 'un', 'un', 0, 55, 55,0,0,0,0,1,1,1,'A','DIO' );
Insert INTO SPRODUTO (IDPROD, IDLOJA, CODPROD, NMPROD, DSCPROD, UNIDCONTROLE, UNIDCOMPRA, UNIDVENDA, VLULTCOMPRA, VLMEDIO, VLVENDA, QTDMIN, QTDMAX, QTDCOMPRA, EMATPRIMA, EPRODACABADO, EVENDA, ESERVICO, SITPROD, IDUSU) Values (2, 0, '2', 'Creme Dpil','Creme Dpil', 'pç', 'pç', 'pç', 10, 10, 25,20,200,100,0,1,1,0,'A','DIO' );
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OTRATAMENTO_PROD on;
Insert INTO OTRATAMENTO_PROD (ID, IDTPSERVICO, IDPROD) Values (1,2,1);
Insert INTO OTRATAMENTO_PROD (ID, IDTPSERVICO, IDPROD) Values (2,3,1);
SET IDENTITY_INSERT OTRATAMENTO_PROD off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SET IDENTITY_INSERT COLIGADA on;
--INSERT  INTO COLIGADA (IDCOLIGADA,NMCOLIGADA, CNPJ, RAZAO) VALUES (1,'GRUPO', '', 'RAZÃO SOCIAL DO GRUPO');
--UPDATE COLIGADA SET TAG='0544497A61757A7C7A7B0802487C60747D76760907081A06061706090509' WHERE IDCOLIGADA=1
--SET IDENTITY_INSERT COLIGADA OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SET IDENTITY_INSERT OLOJA on;
--INSERT  INTO OLOJA(IDLOJA, IDCOLIGADA,CNPJ,NOME,ENDERECO,BAIRRO,CIDADE,ESTADO,INSCEST,EMAIL,TELEFONE1,TELEFONE2,CEP,FAX,NMCONTATO,CARGOCONTATO,DTOPERACAO,DIMENSAO) VALUES (1,1,'','UNIDADE X','Estrada X, 0 Lj.X','Bairro X', 'Cidade X', 'XX', '','','','','','','Fulano','Operadora','1900-01-01','');
--SET IDENTITY_INSERT OLOJA OFF;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT SESTOQUE on;
INSERT INTO SESTOQUE (IDEST, DSCEST, ATIVO) VALUES (1, 'ESTOQUE 01',  1);
SET IDENTITY_INSERT SESTOQUE OFF;
INSERT INTO SEST_LOJA (IDLOJA, IDEST) VALUES (1,  1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- D=Dinheiro, B=Banco
SET IDENTITY_INSERT FCCORRENTE on;
INSERT  INTO FCCORRENTE (IDLOJA, IDCONTA, DSCCONTA, TPCONTA, EVENDA) VALUES(1, 1, 'Caixa', 'D',1 );
INSERT  INTO FCCORRENTE (IDLOJA, IDCONTA, DSCCONTA, TPCONTA, EVENDA) VALUES(1, 2, 'Corrente', 'B', 1);
SET IDENTITY_INSERT FCCORRENTE off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO GPAIS (IDPAIS, DSCPAIS, SIGLAPAIS) VALUES (1, 'Brasil', 'BR');
INSERT INTO FMOEDA (IDMOEDA, IDPAIS, NMMOEDA, SIGLAMOEDA) VALUES (1, 1, 'Real', 'R$' );
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (0,'Não Informado');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (1,'Panfletagem');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (2,'Passante');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (3,'Indicação');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (4,'Carro de Som');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (5,'Televisão');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (6,'Jornal');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (7,'Revista');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (8,'Outdoor');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (9,'Internet');
INSERT INTO OTPCONHEC (IDTPCONHEC, NMCONHEC) VALUES (10,'Outros');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OTPMAQ on;
INSERT INTO OTPMAQ (IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (1, 'FORMED'  , 2, '|GR=1|GR1=1,2,3,4');
INSERT INTO OTPMAQ (IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (2, 'M80'     , 1, '|GR=2|GR1=1|GR2=2,3,4');
INSERT INTO OTPMAQ (IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (3, 'A4'      , 1, '|GR=1|GR1=1,2,3,4');
INSERT INTO OTPMAQ (IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (4, 'BIOFLASH', 1, '|GR=1|GR1=1,2,3,4');
SET IDENTITY_INSERT OTPMAQ off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OTPMANUT on;
INSERT INTO OTPMANUT (IDTPMANUT, DSCMANUT, OBJMANUT) VALUES (1, 'Troca de Água', 1);
INSERT INTO OTPMANUT (IDTPMANUT, DSCMANUT, OBJMANUT) VALUES (2, 'Troca de Lampada', 2);
SET IDENTITY_INSERT OTPMANUT off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OTPMANIPULO on;
INSERT INTO OTPMANIPULO (IDTPMANIPULO, DSCMANIPULO) VALUES (1, 'Pequeno');
INSERT INTO OTPMANIPULO (IDTPMANIPULO, DSCMANIPULO) VALUES (2, 'Grande');
SET IDENTITY_INSERT OTPMANIPULO off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OMANIPULO on;
INSERT INTO OMANIPULO (IDMANIPULO, CODMANIPULO, IDTPMANIPULO) VALUES (1, 'MP01',1);
SET IDENTITY_INSERT OMANIPULO off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OLAMPADA(IDMANIPULO, IDLAMP) VALUES (1, 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 1, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 2, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 3, 1, 1, 0, 10, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 4, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 5, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 6, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 7, 1, 1, 0, 15, 0, 2);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 8, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES ( 9, 1, 1, 0, 15, 0, 2);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (10, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (11, 1, 1, 0, 10, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (12, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (13, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (14, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (15, 1, 1, 0, 10, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (16, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (17, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (18, 1, 1, 0, 20, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (19, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (20, 1, 1, 0, 40, 0, 0);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SET IDENTITY_INSERT OCONTATO on;
--INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME, EMPRESA) VALUES (1, 1, GetDate(), 0, 'Cliente', '');
--INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME, EMPRESA) VALUES (1, 2, GetDate(), 1, 'Dpil', 'Dpil Brasil');
--SET IDENTITY_INSERT OCONTATO off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT FGRUPODESP on;
INSERT INTO FGRUPODESP (IDGRPDESP, DSCGRUPO, IDPAI) VALUES (0, 'Sem Classificação', 0);
SET IDENTITY_INSERT FGRUPODESP off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--Delete From PARAM
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('GLOBAL', 'PATHSETUP', 'Caminho do arquivo de administração.', '%programfiles%\ClasseA\Admin\Dll\');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'DSVM', 'Modo de Operaçcao', 'False');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'MAXLENTEL', 'Quantidade de números do campo Telefone', '15');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'LENIDVENDA', 'Quantidade de dígitos da venda.', '6');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SENHAMESTRE', 'Senha Mestre.', '0709040607');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FRANQCOM', 'Serviço de Comunicação com a Franqueadora', '0');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SOCIOCOM', 'Serviço de Comunicação com os Sócios', '0');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'POP3HOST', 'Servidor POP', 'pop.dpilbrasil.com.br');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SMTPHost', 'Servidor SMTP', 'smtp.dpilbrasil.com.br');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SMTPPort', 'Porta SMTP', '587');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'UseAuthentication', 'Parametro que indica a necessidade de autenticação do do servidor.', '1');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'UsePopAuthentication', 'Parametro que indica a necessidade de autenticação do do servidor POP.', '1');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'DPILSuporte', 'eMail do Suporte Dpil', 'suporte@dpilbrasil.com.br');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FTP', 'Servidor FTP', '465E534044165559554B46565556514C1A5B5A591A5A47');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FTPBakPath', 'Caminho do do Backup no Servidor de FTP', '/Banco/');
--10
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FromDisplayName', 'Nome do usuário de e-Mail', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'MailUID', 'Usuário de e-Mail', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'MailPWD', 'Senha do e-Mail', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FtpUID', 'Usuário do FTP', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FtpPWD', 'Senha do FTP', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SMSURLBASE', 'URL base para oo serviço de SMS', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SMSUID', 'Usuário para o serviço de SMS', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'SMSPWD', 'Senha para o serviço de SMS.', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'LstEMailSocio', 'Lista de e-mails do Sócios', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'LstCelSocio', 'Lista de celulares dos Sócios', '');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'OCLIENTEUNICO', 'Cadastro de Clientes Unificado', '0');
--20
Update PARAM Set VLPARAM = '1'						    Where CODPARAM= 'SOCIOCOM' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'freguesia'					Where CODPARAM= 'FtpUID'   And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '044C4146514B445C5B4B0403'   Where CODPARAM= 'FtpPWD'   And CODSIS= 'P3R';
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OTIPOTAREFA on;
INSERT INTO OTIPOTAREFA (IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (1, 1,  -24, 'Confirmar Agenda');
INSERT INTO OTIPOTAREFA (IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (2, 1, -120, 'Enviar e-Mail de Orientação');
INSERT INTO OTIPOTAREFA (IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (3, 1,   24, 'Verificar Avaliação');
INSERT INTO OTIPOTAREFA (IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (4, 1,   24, 'Verificar Cancelamento/Falta de Sessão');
INSERT INTO OTIPOTAREFA (IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (5, 1,  600, 'Verificar Sessão não Marcada');
SET IDENTITY_INSERT OTIPOTAREFA off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Delete From GFERIADO;
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-01-01', 'Paz', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-01-20', 'Aniversario Cidade', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-04-21', 'Tiradentes', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-04-23', 'São Jorge', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-09-07', 'Independência', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-10-12', 'Padroeira', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-11-15', 'República', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-12-24', 'Natal', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-12-25', 'Natal', 1);
INSERT INTO GFERIADO (DATA, DSCFERIADO, ESCOPO) VALUES ('2011-12-31', 'Paz', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SET IDENTITY_INSERT OMAQUINA on;
DELETE FROM OMAQUINA
INSERT INTO OMAQUINA(IDMAQUINA, IDTPMAQ, CODMAQUINA, DTOPERACAO) VALUES (1, 2, 'M80_01', GETDATE());
INSERT INTO OSALA_MAQUINA (IDLOJA, IDSALA, IDMAQUINA, DTINICIO) VALUES (1, 1, 1, GETDATE());
SET IDENTITY_INSERT OMAQUINA off;