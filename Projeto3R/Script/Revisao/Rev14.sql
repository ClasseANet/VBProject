/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='14', ARQATU='Rev14.sql';
/****************************************************************************
****************************************************************************/
UPDATE OEVENTOAGENDA SET FLGCANCELADO=0 WHERE FLGCANCELADO IS NULL;
SET IDENTITY_INSERT OTRATAMENTO_PROD on;
Insert INTO OTRATAMENTO_PROD (ID, IDTPSERVICO, IDPROD) Values (2,3,1);
SET IDENTITY_INSERT OTRATAMENTO_PROD off;

--Delete From PARAM
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('GLOBAL', 'PATHSETUP', 'Caminho do arquivo de administração.', '%programfiles\ClasseA\Admin\Dll\%');
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
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FTP', 'Servidor FTP', 'ftp.classeanet.com.br');
Insert Into PARAM (CODSIS, CODPARAM, DSCPARAM, VLPARAM) Values ('P3R', 'FTPBakPath', 'Caminho do do Backup no Servidor de FTP', '\Banco\');
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
--20
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

UPDATE SMOVEST SET FLGDELETE=0 WHERE FLGDELETE IS NULL;
UPDATE SMOVEST SET IDFOR = 2 WHERE QTDITEM>0;