USE [G3R];
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM COLIGADA
SET IDENTITY_INSERT COLIGADA on;
Insert into COLIGADA (IDCOLIGADA,NMCOLIGADA) Values (1,'Empresa');
Insert into COLIGADA (IDCOLIGADA,NMCOLIGADA) Values (2,'Grupo R');
SET IDENTITY_INSERT COLIGADA off;
--SELECT * FROM OLOJA
Insert into OLOJA (IDCOLIGADA,IDLOJA,NOME) Values (1,1,'Loja 01');
Insert into OLOJA (IDCOLIGADA,IDLOJA,NOME) Values (2,2,'Freguesia');
Insert into OLOJA (IDCOLIGADA,IDLOJA,NOME) Values (2,3,'Copacabana1');
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM OTPSERVICO
Insert into OTPSERVICO (IDLOJA,IDTPSERVICO, DSCSERVICO) Values (1,1,'Avaliação');
Insert into OTPSERVICO (IDLOJA,IDTPSERVICO, DSCSERVICO) Values (1,2,'Sessão');
Insert into OTPSERVICO (IDLOJA,IDTPSERVICO, DSCSERVICO) Values (1,3,'Manutenção');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into OTPTRATAMENTO (IDLOJA,IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (1,1,'Depilação',30);
Insert into OTPTRATAMENTO (IDLOJA,IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (1,2,'Rejuvenescimento',21);
Insert into OTPTRATAMENTO (IDLOJA,IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (1,3,'Mancha',15);
Insert into OTPTRATAMENTO (IDLOJA,IDTPTRATAMENTO, DSCTRATAMENTO, FREQUENCIA) Values (1,4,'Acne',3);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,1,'Abdômen');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,2,'Aréola Mamaria');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,3,'Axilas');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,4,'Braço Direito');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,5,'Braço Esquerdo');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,6,'Dorso');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,7,'Face');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,8,'Glúteos');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,9,'Lombar');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,10,'Mãos');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,11,'Nuca');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,12,'Ombros');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,13,'Peito');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,14,'Pés');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,15,'Perianal');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,16,'Meia-Perna (Canelas)');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,17,'Meia-Perna (Coxas)' );
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,18,'Pescoço(Frontal)');
Insert into OAREA (IDLOJA,IDAREA, DSCAREA) Values (1,19,'Virilhas');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO ODIRECAO (IDLOJA,IDDIRECAO, DSCDIRECAO, NORDEM) VALUES (1,1, 'Horizontal', 1);
INSERT INTO ODIRECAO (IDLOJA,IDDIRECAO, DSCDIRECAO, NORDEM) VALUES (1,2, 'Vertical', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM OFOTOTIPO 
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (1, 'Branca +Clara', 'Muitas', 'Loiro', 'Azuis', 'Céltico', 'Muito severas e dolorosas', 'Nenhum, vermelhidão');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (2, 'Branca -Clara', 'Raras Vezes', 'Loiro Escuro/Ruivo', 'Azuis/Cinzentados', 'Europeu(12%)', 'Sempre severas e muito dolorosas', 'Raramente');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (3, 'Avermelhada', 'Muitas', 'Castanho', 'Castanhos', 'Europeu(78%)', 'Menos frequentes, ligeiras ', 'Médio');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (4, 'Castanho Claro', 'Muitas', 'Castanho Escuro', 'Escuros', 'Mediterrâneo', 'Não', 'Rápido e profundo');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (5, 'Castanho Escuro', 'Muitas', 'Preto', 'Escuros', 'Raças Magreb/Índias', 'Não', 'Rápido e profundo');
INSERT INTO OFOTOTIPO (FOTOTIPO, PELE, SARDAS, PELO, OLHOS, DEFINICAO, QUEIMADURA, BRONZEADO) VALUES (6, 'Negra', 'Muitas', 'Preto/Grisalho', 'Negros', 'Negras', 'Não', 'Rápido e profundo');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT OREACAOADV (IDLOJA,IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES (1,'1','Edema','0',convert(datetime,'2010-07-02 22:29:04.890',121));
INSERT OREACAOADV (IDLOJA,IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES (1,'2','Hiper/Hipopigmentação','0',convert(datetime,'2010-07-02 22:29:54.263',121));
INSERT OREACAOADV (IDLOJA,IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES (1,'3','Casca','0',convert(datetime,'2010-07-02 22:30:04.200',121));
INSERT OREACAOADV (IDLOJA,IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES (1,'4','Bolhas','0',convert(datetime,'2010-07-02 22:30:09.700',121));
INSERT OREACAOADV (IDLOJA,IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES (1,'5','Desepetelização','0',convert(datetime,'2010-07-02 22:30:19.983',121));
INSERT OREACAOADV (IDLOJA,IDREACAOADV, DSCREACAOADV, ALTERSTAMP, TIMESTAMP) VALUES (1,'6','Pelo Encravado','0',convert(datetime,'2010-07-02 22:35:00.687',121));
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into CCONDPGTO (IDLOJA,IDCONDPGTO, CONDPGTO, ATIVO) Values (1,1,'À VISTA', 1);
Insert into CCONDPGTO (IDLOJA,IDCONDPGTO, CONDPGTO, ATIVO) Values (1,2,'30 DDL', 1);
Insert into CCONDPGTO (IDLOJA,IDCONDPGTO, CONDPGTO, ATIVO) Values (1,3,'30/60 DDL', 1);
Insert into CCONDPGTO (IDLOJA,IDCONDPGTO, CONDPGTO, ATIVO) Values (1,4,'30/60/90 DDL', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert into CFORMAPGTO (IDLOJA,IDFORMAPGTO, FORMAPGTO, ATIVO) Values (1,1,'Dinheiro', 1);
Insert into CFORMAPGTO (IDLOJA,IDFORMAPGTO, FORMAPGTO, ATIVO) Values (1,2,'Débito', 1);
Insert into CFORMAPGTO (IDLOJA,IDFORMAPGTO, FORMAPGTO, ATIVO) Values (1,3,'Crédito', 1);
Insert into CFORMAPGTO (IDLOJA,IDFORMAPGTO, FORMAPGTO, ATIVO) Values (1,4,'Cheque', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert INTO SPRODUTO (IDPROD, IDLOJA, CODPROD, NMPROD, DSCPROD, UNIDCONTROLE, UNIDCOMPRA, UNIDVENDA, VLULTCOMPRA, VLMEDIO, VLVENDA, QTDMIN, QTDMAX, QTDCOMPRA, EMATPRIMA, EPRODACABADO, EVENDA, ESERVICO, SITPROD, IDUSU) Values (1, 1, '1', 'Serviço Único','Serviço Único', 'un', 'un', 'un', 0, 55, 55,0,0,0,0,1,1,1,'A','DIO' );
Insert INTO SPRODUTO (IDPROD, IDLOJA, CODPROD, NMPROD, DSCPROD, UNIDCONTROLE, UNIDCOMPRA, UNIDVENDA, VLULTCOMPRA, VLMEDIO, VLVENDA, QTDMIN, QTDMAX, QTDCOMPRA, EMATPRIMA, EPRODACABADO, EVENDA, ESERVICO, SITPROD, IDUSU) Values (2, 1, '2', 'Creme Dpil','Creme Dpil', 'pç', 'pç', 'pç', 12, 12, 29,20,200,100,0,1,1,0,'A','DIO' );
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert INTO OTRATAMENTO_PROD (IDLOJA,ID, IDTPSERVICO, IDPROD) Values (1,1,2,1);
Insert INTO OTRATAMENTO_PROD (IDLOJA,ID, IDTPSERVICO, IDPROD) Values (1,2,3,1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM SESTOQUE 
--SELECT * FROM SEST_LOJA 
INSERT INTO SESTOQUE (IDLOJA,IDEST, DSCEST, ATIVO) VALUES (1, 1,'ESTOQUE 01',  1);

IF EXISTS(SELECT IDLOJA FROM OLOJA WHERE IDLOJA=1) INSERT INTO SEST_LOJA (IDLOJA, IDEST) VALUES (1,  1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
-- D=Dinheiro, B=Banco
INSERT  INTO FCCORRENTE (IDLOJA, IDCONTA, DSCCONTA, TPCONTA, EVENDA) VALUES(1, 1, 'Caixa', 'D',1 );
INSERT  INTO FCCORRENTE (IDLOJA, IDCONTA, DSCCONTA, TPCONTA, EVENDA) VALUES(1, 2, 'Corrente', 'B', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SELECT * FROM GPAIS
--SELECT * FROM FMOEDA
INSERT INTO GPAIS (IDPAIS, DSCPAIS, SIGLAPAIS) VALUES (1, 'Brasil', 'BR');
INSERT INTO FMOEDA (IDMOEDA, IDPAIS, NMMOEDA, SIGLAMOEDA) VALUES (1, 1, 'Real', 'R$' );
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,0,'Não Informado');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,1,'Panfletagem');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,2,'Passante');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,3,'Indicação');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,4,'Carro de Som');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,5,'Televisão');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,6,'Jornal');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,7,'Revista');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,8,'Outdoor');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,9,'Internet');
INSERT INTO OTPCONHEC (IDLOJA,IDTPCONHEC, NMCONHEC) VALUES (1,10,'Outros');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OTPMAQ (IDLOJA,IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (1,1, 'FORMED'  , 2, '|GR=1|GR1=1,2,3,4');
INSERT INTO OTPMAQ (IDLOJA,IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (1,2, 'M80'     , 1, '|GR=2|GR1=1|GR2=2,3,4');
INSERT INTO OTPMAQ (IDLOJA,IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (1,3, 'A4'      , 1, '|GR=1|GR1=1,2,3,4');
INSERT INTO OTPMAQ (IDLOJA,IDTPMAQ, DSCMAQ, TPNUM, NUMTRATAMENTO) VALUES (1,4, 'BIOFLASH', 1, '|GR=1|GR1=1,2,3,4');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OTPMANUT (IDLOJA,IDTPMANUT, DSCMANUT, OBJMANUT) VALUES (1,1, 'Troca de Água', 1);
INSERT INTO OTPMANUT (IDLOJA,IDTPMANUT, DSCMANUT, OBJMANUT) VALUES (1,2, 'Troca de Lampada', 2);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OTPMANIPULO (IDLOJA,IDTPMANIPULO, DSCMANIPULO) VALUES (1,1, 'Pequeno');
INSERT INTO OTPMANIPULO (IDLOJA,IDTPMANIPULO, DSCMANIPULO) VALUES (1,2, 'Grande');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OMANIPULO (IDLOJA,IDMANIPULO, CODMANIPULO, IDTPMANIPULO) VALUES (1,1, 'MP01',1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OLAMPADA(IDLOJA,IDMANIPULO, IDLAMP) VALUES (1,1, 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 1, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 2, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 3, 1, 1, 0, 10, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 4, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 5, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 6, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 7, 1, 1, 0, 15, 0, 2);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 8, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1, 9, 1, 1, 0, 15, 0, 2);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,10, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,11, 1, 1, 0, 10, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,12, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,13, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,14, 1, 1, 0, 15, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,15, 1, 1, 0, 10, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,16, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,17, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,18, 1, 1, 0, 20, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,19, 1, 1, 0, 30, 0, 0);
INSERT INTO OAREA_TRATAMENTO (IDLOJA,IDAREA, IDTPTRATAMENTO, IDTPMANIPULO, TEMPOMIN, TEMPOMED, TEMPOMAX, INDICEH) VALUES (1,20, 1, 1, 0, 40, 0, 0);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME, EMPRESA) VALUES (1, 1, GetDate(), 0, 'Cliente', '');
INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME, EMPRESA) VALUES (1, 2, GetDate(), 1, 'Dpil', 'Dpil Brasil');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO FGRUPODESP (IDLOJA,IDGRPDESP, DSCGRUPO, IDPAI) VALUES (1,0, 'Sem Classificação', 0);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--Delete From PARAM
--SELECT * FROM PARAM 
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'GLOBAL', 'PATHSETUP', 'Caminho do arquivo de administração.', '%programfiles%\ClasseA\Admin\Dll\');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'DSVM', 'Modo de Operaçcao', 'False');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'MAXLENTEL', 'Quantidade de números do campo Telefone', '15');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'LENIDVENDA', 'Quantidade de dígitos da venda.', '6');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SENHAMESTRE', 'Senha Mestre.', '0709040607');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'FRANQCOM', 'Serviço de Comunicação com a Franqueadora', '0');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SOCIOCOM', 'Serviço de Comunicação com os Sócios', '0');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'POP3HOST', 'Servidor POP', 'pop.dpilbrasil.com.br');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SMTPHost', 'Servidor SMTP', 'smtp.dpilbrasil.com.br');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SMTPPort', 'Porta SMTP', '587');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'UseAuthentication', 'Parametro que indica a necessidade de autenticação do do servidor.', '1');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'UsePopAuthentication', 'Parametro que indica a necessidade de autenticação do do servidor POP.', '1');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'DPILSuporte', 'eMail do Suporte Dpil', 'suporte@dpilbrasil.com.br');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'FTP', 'Servidor FTP', '465E534044165559554B46565556514C1A5B5A591A5A47');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'FTPBakPath', 'Caminho do do Backup no Servidor de FTP', '');
--10
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'FromDisplayName', 'Nome do usuário de e-Mail', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'MailUID', 'Usuário de e-Mail', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'MailPWD', 'Senha do e-Mail', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'FtpUID', 'Usuário do FTP', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'FtpPWD', 'Senha do FTP', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SMSURLBASE', 'URL base para o serviço de SMS', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SMSUID', 'Usuário para o serviço de SMS', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'SMSPWD', 'Senha para o serviço de SMS.', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'LstEMailSocio', 'Lista de e-mails do Sócios', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'LstCelSocio', 'Lista de celulares dos Sócios', '');
Insert Into PARAM (IDLOJA,CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values (1,'P3R', 'OCLIENTEUNICO', 'Cadastro de Clientes Unificado', '0');
--20
Update PARAM Set VLPARAM = '1'				Where CODPARAM= 'SOCIOCOM' And CODSIS= 'P3R';
Update PARAM Set VLPARAM = 'clientedpil'		Where CODPARAM= 'FtpUID'   And CODSIS= 'P3R';
Update PARAM Set VLPARAM = '045C51445D5407050508'   	Where CODPARAM= 'FtpPWD'   And CODSIS= 'P3R';
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO OTIPOTAREFA (IDLOJA,IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (1,1, 1,  -24, 'Confirmar Agenda');
INSERT INTO OTIPOTAREFA (IDLOJA,IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (1,2, 1, -120, 'Enviar e-Mail de Orientação');
INSERT INTO OTIPOTAREFA (IDLOJA,IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (1,3, 1,   24, 'Verificar Avaliação');
INSERT INTO OTIPOTAREFA (IDLOJA,IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (1,4, 1,   24, 'Verificar Cancelamento/Falta de Sessão');
INSERT INTO OTIPOTAREFA (IDLOJA,IDTPTAREFA, ATIVO, DIFHORA, DSCTPTAREFA) VALUES (1,5, 1,  600, 'Verificar Sessão não Marcada');
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Delete From GFERIADO;
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-01-01', 'Paz', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-01-20', 'Aniversario Cidade', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-04-21', 'Tiradentes', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-04-23', 'São Jorge', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-09-07', 'Independência', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-10-12', 'Padroeira', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-11-15', 'República', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-12-24', 'Natal', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-12-25', 'Natal', 1);
INSERT INTO GFERIADO (IDLOJA,DATA, DSCFERIADO, ESCOPO) VALUES (1,'2011-12-31', 'Paz', 1);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO PMETA (IDLOJA, IDMETA, FAIXA1, FAIXA2, FAIXA3) VALUES (1, 1, 280, 320, 400);
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--SET IDENTITY_INSERT OMAQUINA on;
--DELETE FROM OMAQUINA
--INSERT INTO OMAQUINA(IDMAQUINA, IDTPMAQ, CODMAQUINA, DTOPERACAO) VALUES (1, 2, 'M80_01', '2010-01-01');
--INSERT INTO OSALA_MAQUINA (IDLOJA, IDSALA, IDMAQUINA, DTINICIO) VALUES (1, 1, 1, '2010-01-01');
--SET IDENTITY_INSERT OMAQUINA off;
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
--DELETE FROM RFUNCIONARIO
--SELECT * FROM RFUNCIONARIO
INSERT INTO RFUNCIONARIO(IDFUNCIONARIO, IDLOJA, CHAPA, NOME, DTADMISSAO, FLGCERTIFICADO, SITFUNC, DTCADASTRO) VALUES (1, 1, '000001', 'Funcionario01', '2010-01-01', 1, 'A', GETDATE());
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
INSERT INTO FDESPESA (IDLOJA,IDDESP,DSCDESP,CODDESP,IDPAI,IDGRPDESP,ATIVO,OCULTO,TPDESP,ALTERSTAMP,TIMESTAMP) VALUES (1,1,'Vendas'  ,'1.01'   ,0,0,1,0,1,1,getdate());
INSERT INTO FDESPESA (IDLOJA,IDDESP,DSCDESP,CODDESP,IDPAI,IDGRPDESP,ATIVO,OCULTO,TPDESP,ALTERSTAMP,TIMESTAMP) VALUES (1,2,'Material','2.01'   ,0,0,1,0,2,1,getdate());
INSERT INTO FDESPESA (IDLOJA,IDDESP,DSCDESP,CODDESP,IDPAI,IDGRPDESP,ATIVO,OCULTO,TPDESP,ALTERSTAMP,TIMESTAMP) VALUES (1,3,'Consumo' ,'2.01.01',0,0,1,0,2,1,getdate());
INSERT INTO FDESPESA (IDLOJA,IDDESP,DSCDESP,CODDESP,IDPAI,IDGRPDESP,ATIVO,OCULTO,TPDESP,ALTERSTAMP,TIMESTAMP) VALUES (1,4,'Serviço' ,'2.02'   ,0,0,1,0,2,1,getdate());
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
