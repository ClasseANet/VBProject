Attribute VB_Name = "PADRAO"
Option Explicit
'============================
'= Classes                  =
'============================
Global Sys  As New SETTING
Global DB As New DS_BANCO
Global BANCO As New BANCO_TK
'Global ClsUser As New User
Global SysMdi As New MdiPrincipal
'**************************************************************
'======================= Constantes do Sistema ================
'**************************************************************
'===============================
'= Índice do Toolbar Principal =
'===============================
Enum BOT
   BT01_BANCO = 1
   BT02_CLS = 2
   BT03_PRJ = 3
   BT04_TABELAS = 4
   BT05_05 = 5
   BT06_RPT = 6
   BT07_RPT = 7
   BT08_08 = 8
   BT09_09 = 9
   BT10_REFRESH = 10
   BT11_FIND = 11
   BT12_SAVE = 12
   BT13_DEL = 13
   BT14_VOLTAR = 14
   BT15_SAIR = 15
   BT16_PAIS = 16
   BT17_USU = 17
   BT18_UTILIT = 18
   BT19_BACKUP = 19
   BT20_ABOUT = 20
   BT21_HELP = 21
End Enum
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
'      DSR100.Sistema = UCase(Sys.AppName$ + " - " + AppTitle$)
      .FundoTela = Sys.FundoTela
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
      .dBase = DB.dBase.Name
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
   Dim Img As Object
   Dim i%, lBot
   
   Set SysMdi.Toolbar.ImageList = SysMdi.ImageList1
   Set Img = SysMdi.ImageList1
   
   SysMdi.Toolbar.AllowCustomize = False

   For Each lBot In SysMdi.Toolbar.Buttons
      i = i + 1
      SysMdi.Toolbar.Buttons(i).Style = tbrSeparator
      SysMdi.Toolbar.Buttons(i).Visible = False
   Next
   With SysMdi.Toolbar.Buttons
      Call AbilitarToolBar(.Item(BOT.BT01_BANCO), Img, True, "BANCO")
      Call AbilitarToolBar(.Item(BOT.BT02_CLS), Img, True, "CLASSE")
      Call AbilitarToolBar(.Item(BOT.BT03_PRJ), Img, True, "SQL")
      'Call AbilitarToolBar(.Item(BOT.BT04_TABELAS), Img, True, "SQL")
      'Call AbilitarToolBar(.Item(BOT.BT05_05), Img, False) ', "PRINT")
      Call AbilitarToolBar(.Item(BOT.BT06_RPT), Img, False, "PRINT")
      Call AbilitarToolBar(.Item(BOT.BT07_RPT), Img, False, "PRINT")
      Call AbilitarToolBar(.Item(BOT.BT08_08), Img, False, "PRINT")
      Call AbilitarToolBar(.Item(BOT.BT09_09), Img, False, "PRINT")
      '* Edit
      Call AbilitarToolBar(.Item(BOT.BT10_REFRESH), Img, True, "REFRESH")
      Call AbilitarToolBar(.Item(BOT.BT11_FIND), Img, False) ', "SEEK")
      Call AbilitarToolBar(.Item(BOT.BT12_SAVE), Img, False) ', "SAVE")
      Call AbilitarToolBar(.Item(BOT.BT13_DEL), Img, False) ', "DELETE")
      Call AbilitarToolBar(.Item(BOT.BT14_VOLTAR), Img, True, "BACK")
      Call AbilitarToolBar(.Item(BOT.BT15_SAIR), Img, True, "EXIT")
      Call AbilitarToolBar(.Item(BOT.BT16_PAIS), Img, False)
      Call AbilitarToolBar(.Item(BOT.BT17_USU), Img, False) ', "CHAVE")
      Call AbilitarToolBar(.Item(BOT.BT18_UTILIT), Img, False, "TOOLS")
      Call AbilitarToolBar(.Item(BOT.BT19_BACKUP), Img, False, "FITA")
      Call AbilitarToolBar(.Item(BOT.BT20_ABOUT), Img, True, "OLHO")
      Call AbilitarToolBar(.Item(BOT.BT21_HELP), Img, True, "HELP")
      
      SysMdi.Toolbar.Refresh
   
      .Item(1).ToolTipText = "Abrir Banco de Dados"
      .Item(2).ToolTipText = "Localizar / Analizar Carga"
      .Item(3).ToolTipText = "Descrever Projeto"
      .Item(4).ToolTipText = "04-"
      .Item(5).ToolTipText = "05-"
      .Item(6).ToolTipText = "06-"
      .Item(7).ToolTipText = "Relatório"
      .Item(8).ToolTipText = "08-"
      .Item(9).ToolTipText = "09-"
      '==============================
      '==============================
      .Item(10).ToolTipText = "Atualizar Tela"
      .Item(11).ToolTipText = "Procurar"
      .Item(12).ToolTipText = "Salvar"
      .Item(13).ToolTipText = "Excluir"
      .Item(14).ToolTipText = "Tela Anterior"
      .Item(15).ToolTipText = "Sair do Sistema"
      .Item(16).ToolTipText = "16-"
      .Item(17).ToolTipText = "Usuário"
      .Item(18).ToolTipText = "Atualizar Base de Dados"
      .Item(19).ToolTipText = "Cópia de Segurança"
      .Item(20).ToolTipText = "Informação Sobre o Sistema..."
      .Item(21).ToolTipText = "Ajuda"
      For i = 1 To .Count
         .Item(i).Description = .Item(i).ToolTipText
      Next
   End With
   SysMdi.Toolbar.Refresh
End Sub
Public Sub GetConfig()
   '*** [ Database Format ] ***
   Sys.dbODBC = GetSetting(Sys.AppName, "Database Format", "DBODBC", "N")
   Sys.dbVersao = GetSetting(Sys.AppName, "Database Format", "DBVERSAO", "ACCESS3.0")
   Select Case UCase(Sys.dbVersao)
      Case "ACCESS1.0": Sys.dbVersion = dbVersion10
      Case "ACCESS1.1": Sys.dbVersion = dbVersion11
      Case "ACCESS2.0": Sys.dbVersion = dbVersion20
      Case "ACCESS3.0": Sys.dbVersion = dbVersion30
      Case Else: Sys.dbVersion = 0
   End Select
   Sys.dbName = GetSetting(Sys.AppName, "Database Format", "DBNAME", UCase(Sys.AppExe) + ".MDB")
   
   '*** [ Database Drive ] ***
   Sys.dbDrive = GetSetting(Sys.AppName, "Database Drive", "DBDRIVE", "C:\DSR\" + UCase(Sys.AppName) + "\")
   Sys.DrvRpt = GetSetting(Sys.AppName, "Database Drive", "DRVRPT", "C:\DSR\" + UCase(Sys.AppName) + "\REPORT\")
   
   '*** [ Setup ] ***
   Sys.MICRO = GetSetting(Sys.AppName, "Setup", "MICRO", "RUN TIME")
   Select Case GetSetting(Sys.AppName, "Setup", "IDIOMA", "Portugues")
      Case "Portugues": Sys.Idioma = 5000
      Case "Ingles": Sys.Idioma = 6000
   End Select
   Sys.FundoTela = GetSetting(Sys.AppName, "Setup", "FUNDOTELA", "FUNDO")

   Sys.DrvDrive = SysMdi.Drv1.List(0) + "\"
   Sys.dbDrive_Orig = Sys.dbDrive

End Sub
