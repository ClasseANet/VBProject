Attribute VB_Name = "PADRAO"
Option Explicit
'============================
'= Classes                  =
'============================
Global DB As New DS_BANCO
'Global LOV As New LOV
Global Sys As New SETTING

'============================
'= Arquivo de Inicialização =
'============================
Global dbDrive As String
Global dbName As String
Global dbVersao As String
Global dbVersion As String
'Global isODBC As Boolean
Global DrvRpt As String
Global DrvDrive As String
Global dbDrive_Orig As String
Global MICRO As String
Global FundoTela As String
Global MDIFilho As Form
Global SysMdi As MDIForm
'============================
'= Variáveis de Território  =
'============================
Global Sys_CodLocal$ 'Código do Local
Global Sys_DscLocal$ 'Descrição do Local
Global Sys_IdPais$   'Código do País
Global Sys_DscPais$  'Descrição do País
'**************************************************************
'======================= Constantes do Sistema ================
'**************************************************************
'===============================
'= Índice do Toolbar Principal =
'===============================
Global Const BT_BANCO = 1
Global Const BT_CLS = 2
Global Const BT_PRJ = 3

Global Const BT_TABELAS = 4
'Global Const BT_PRJ = 5
Global Const BT_RPT = 7
'Global Const BT_08 = 8
'Global Const BT_09 = 9
Global Const BT_REFRESH = 10
Global Const BT_FIND = 11
Global Const BT_SAVE = 12
Global Const BT_DEL = 13
Global Const BT_VOLTAR = 14
Global Const BT_SAIR = 15
Global Const BT_PAIS = 16
Global Const BT_USU = 17
Global Const BT_UTILIT = 18
Global Const BT_BACKUP = 19
Global Const BT_ABOUT = 20
Global Const BT_HELP = 21
Public Function F_LOV(Tabela$)
   Dim Sql As Variant 'Coluna de ordenação (Padrao = 2 )
                      'ou Query de acesso
   Dim Cab, IdCampo, Tit$
   Dim MyClass As New LOV
   Select Case UCase(Tabela)
      Case "PIECE"
         Cab = Array("Código", "IDPIECE", 10, vbLeftJustify, _
                      "Descrição", "DSCPIECE", 30, vbLeftJustify)
         IdCampo = Array("IDPIECE")
         Tit$ = "Lista de Material / Serviço"
      Case "SUPPLIER"
         Cab = Array("Código", "IDSUPP", 10, vbLeftJustify, _
                      "Descrição", "NMSUPP", 30, vbLeftJustify)
         IdCampo = Array("IDSUPP")
         Tit$ = "Lista de Fornecedor"
   End Select
   With MyClass
'      DSR100.Sistema = UCase(sys.AppName$ + " - " + AppTitle$)
      .FundoTela = FundoTela
      .Tipo = "LOV"
      
      .Versao = Sys.AppVer
      .Empresa = Sys.NomeEmpresa
      
      Set .dBase = DB
      .Table = UCase(Tabela)
'      .Query = Sql
      .Cab = Cab
      .IdField = IdCampo
           
      .Caption = Tit$
      .Show
      F_LOV = .Id
   End With
   Set MyClass = Nothing
End Function
Public Sub F_CAD(Tabela$, Optional Ac$ = "")
   Dim i%
   Dim Sql$, Cab, Id, Tit$, Frm As Form
   Dim MyClass As New CAD
   
   On Error Resume Next
   Select Case UCase(Tabela)
      Case "TB_PAIS"
         Tit$ = "País"
         Sql = "select IDPAIS,DSCPAIS,MOEDA,NMMOEDA "
         Sql = Sql + " from TB_PAIS "
         Sql = Sql + " order by 2" 'DSCPAIS;"
         Cab = Array("Cod.", "IDPAIS", 4, vbLeftJustify, _
                  "País", "DSCPAIS", 15, vbLeftJustify, _
                  "Moeda", "MOEDA", 6, vbCenter, _
                  "Descrição", "NMMOEDA", 15, vbLeftJustify)
         Id = Array("IDPAIS", 0, "C")
'         Set Frm = FrmCadPais
   End Select
   With MyClass
      .Acesso = Ac
'      .MeForm = New FrmCad
      .dBase = DB.dBase
'SE LOV   With LOV
'SE LOV      .Tipo = "CAD"
      .Table = UCase(Tabela)
      .Cab = Cab
      .Caption = Tit$
      .IdField = Id
      .Query = Sql$
      
      .FrmCad = Frm
      .Show
   End With
End Sub
Public Sub PreenchePais(Frm As Form)
   Dim Sql$, i%
   Sql$ = "select DSCPAIS from TB_PAIS order by DSCPAIS;"
   Call MontarDbCombo(DB, Frm.CmbPais, Sql$, "DSCPAIS")
End Sub
Public Sub ToolbarChild()
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Abilitar e desabilitar os botões, do Toolbar    **
'**            pricipal, necessários para o FormRM             **
'**                                                            **
'** Recebe: Oper$ - Título da operação que está sen realizada  **                   **
'**                                                            **
'** Retorna : Título Principal ajustado                        **
'**                                                            **
'****************************************************************
   Dim BOTAO As Buttons
   
   Set SysMdi.Toolbar.ImageList = SysMdi.ImageList1
   Set BOTAO = SysMdi.Toolbar.Buttons
     'Call AbilitarObj(BOTAO.Item(REFRESH), True, SysMdi.MnuEdit(0))
     'Call AbilitarObj(BOTAO.Item(FIND), True, SysMdi.MnuEdit(1))
     'Call AbilitarObj(BOTAO.Item(SAVE), True, SysMdi.MnuEdit(2))
     'Call AbilitarObj(BOTAO.Item(DEL), True, SysMdi.MnuEdit(3))

    
'     Call AbilitarBotao(BOTAO.Item(USU), False, 207)  "chave0.bmp")
'     Call AbilitarBotao(BOTAO.Item(SAIR), False, 206) "door0.bmp")
     SysMdi.Toolbar.Refresh
End Sub
Public Sub ToolbarPrincipal()
   Dim ObjTBL As Object
   Dim Img As Object
   Dim i%
   
   Set SysMdi.Toolbar.ImageList = SysMdi.ImageList1
   Set ObjTBL = SysMdi.Toolbar
   Set Img = SysMdi.ImageList1
   Call AbilitarToolBar(ObjTBL, Img, 0, "BANCO", "Abrir Banco de Dados")
   Call AbilitarToolBar(ObjTBL, Img, 0, "CLASSE", "Montar Classe")
   Call AbilitarToolBar(ObjTBL, Img, 0, "TABLES", "Descrever Projeto")
   Call AbilitarToolBar(ObjTBL, Img, 0, "TABLES", "Montar 'Setup...'")
   Call AbilitarToolBar(ObjTBL, Img, 0, , "Project")
   Call AbilitarToolBar(ObjTBL, Img, 0)
   Call AbilitarToolBar(ObjTBL, Img, 0)  '0, "PRINT")
   Call AbilitarToolBar(ObjTBL, Img, 0)
   Call AbilitarToolBar(ObjTBL, Img, 0)
   Call AbilitarToolBar(ObjTBL, Img, 0)  ', 0, "REFRESH", "Atualizar Tela")
   Call AbilitarToolBar(ObjTBL, Img, 0)  '0, "SEEK","Procurar")
   Call AbilitarToolBar(ObjTBL, Img, 0)  '0, "SAVE","Salvar")
   Call AbilitarToolBar(ObjTBL, Img, 0)  '0, "DELETE","Excluir")
   Call AbilitarToolBar(ObjTBL, Img, 0, "BACK", "Tela Anterior")
   Call AbilitarToolBar(ObjTBL, Img, 0, "EXIT", "Sair do Sistema")
   Call AbilitarToolBar(ObjTBL, Img, 0)
   Call AbilitarToolBar(ObjTBL, Img, 0, "CHAVE", "Usuário")
   Call AbilitarToolBar(ObjTBL, Img, 0, "TOOLS", "Atualizar Base de Dados")
   Call AbilitarToolBar(ObjTBL, Img, 0, , "Cópia de Segurança") '"FITA"
   Call AbilitarToolBar(ObjTBL, Img, 0, "OLHO", "Informação Sobre o Sistema...")
   Call AbilitarToolBar(ObjTBL, Img, 0, "HELP", "Ajuda")
    
   SysMdi.Toolbar.Refresh
End Sub
Public Sub GetConfig()
   '*** [ Database Format ] ***
   Sys.isODBC = GetSetting(Sys.AppName, "Database Format", "isODBC", False)
   dbVersao = GetSetting(Sys.AppName, "Database Format", "DBVERSAO", "ACCESS3.0")
   Select Case UCase(dbVersao)
      Case "ACCESS1.0": dbVersion = dbVersion10
      Case "ACCESS1.1": dbVersion = dbVersion11
      Case "ACCESS2.0": dbVersion = dbVersion20
      Case "ACCESS3.0": dbVersion = dbVersion30
      Case Else: dbVersion = 0
   End Select
   dbName = GetSetting(Sys.AppName, "Database Format", "DBNAME", UCase(Sys.Appexe) + ".MDB")
   
   '*** [ Database Drive ] ***
   dbDrive = GetSetting(Sys.AppName, "Database Drive", "DBDRIVE", "C:\DSR\" + UCase(Sys.AppName) + "\")
   DrvRpt = GetSetting(Sys.AppName, "Database Drive", "DRVRPT", "C:\DSR\" + UCase(Sys.AppName) + "\REPORT\")
   
   '*** [ Setup ] ***
   MICRO = GetSetting(Sys.AppName, "Setup", "MICRO", "RUN TIME")
   Select Case GetSetting(Sys.AppName, "Setup", "IDIOMA", "Portugues")
      Case "Portugues": Sys.Idioma = 5000
      Case "Ingles": Sys.Idioma = 6000
   End Select
   FundoTela = GetSetting(Sys.AppName, "Setup", "FUNDOTELA", "FUNDO")

   DrvDrive = SysMdi.Drv1.List(0) + "\"
   dbDrive_Orig = dbDrive

End Sub
