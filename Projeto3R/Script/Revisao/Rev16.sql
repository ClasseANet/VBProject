/****************************************************************************
****************************************************************************/
USE G3R;
if NOt Exists(Select * From VERSAOBD Where IDBD=1) INSERT INTO VERSAOBD(IDBD, DSCBD, VSBD, ATUBD, DTATU, ARQATU) VALUES (1, 'Banco Dpil', '1.0', '0', GetDate(), '');
UPDATE VERSAOBD SET VSBD='1.0', DTATU=GetDate(), ATUBD='16', ARQATU='Rev16.sql';
/****************************************************************************
****************************************************************************/
Alter Table RFUNCIONARIO ADD [COMPROD] [int] NULL;
Alter Table RFUNCIONARIO ADD [VLCOMPROD] [decimal](9, 2) NULL;
Alter Table RFUNCIONARIO ADD [TPCOMPROD] [int] NULL;
Alter Table RFUNCIONARIO ADD [COMSERV] [int] NULL;
Alter Table RFUNCIONARIO ADD [VLCOMSERV] [decimal](9, 2) NULL;
Alter Table RFUNCIONARIO ADD [TPCOMSERV] [int] NULL;
Alter Table RFUNCIONARIO ADD [OBS] [VARCHAR](80) NULL;
Alter Table RFUNCIONARIO ADD [TELEFONE] [varchar](20) NULL;
Alter Table RFUNCIONARIO ADD [CELULAR] [varchar](20) NULL;
Alter Table RFUNCIONARIO ADD [EMAIL] [varchar](80) NULL;
Alter Table RFUNCIONARIO ADD [ENDERECO] [varchar](100) NULL;
Alter Table RFUNCIONARIO ADD [BAIRRO] [varchar](50) NULL;
Alter Table RFUNCIONARIO ADD [CIDADE] [varchar](50) NULL;
Alter Table RFUNCIONARIO ADD [ESTADO] [varchar](2) NULL;
Alter Table RFUNCIONARIO ADD [CEP] [varchar](15) NULL;
Alter Table RFUNCIONARIO ADD [PAIS] [varchar](30) NULL;
Alter Table RFUNCIONARIO ADD [DTNASC] [datetime] NULL;
Alter Table RFUNCIONARIO ADD [SALARIO] [decimal](9,2) NULL;
Alter Table RFUNCIONARIO ADD [DTCADASTRO] [datetime] NULL;


exec sp_bindefault DF_0, 'RFUNCIONARIO.COMPROD';
exec sp_bindefault DF_0, 'RFUNCIONARIO.TPCOMPROD';
exec sp_bindefault DF_0, 'RFUNCIONARIO.COMSERV';
exec sp_bindefault DF_0, 'RFUNCIONARIO.TPCOMSERV';
