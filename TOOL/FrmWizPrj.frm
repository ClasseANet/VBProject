VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmWizPrj 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Wizard..."
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   10830
   Icon            =   "FrmWizPrj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6735
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabPrj 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Classes"
      TabPicture(0)   =   "FrmWizPrj.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lbl(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lbl(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lbl(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbl(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtNmDbObj"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LstOp"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CmbOwner"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdCarregaVetor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "OptInPrj(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "OptInPrj(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtSuperClasse"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "LstCampoCls"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "CmdDrv(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtDrvDest"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LstTabCls"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CmdChk(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CmdChk(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CmdChk(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdScript"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CmdOperCls"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Formularios"
      TabPicture(1)   =   "FrmWizPrj.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmbNotNull"
      Tab(1).Control(1)=   "LstTabForm"
      Tab(1).Control(2)=   "CmbCtrl"
      Tab(1).Control(3)=   "TxtCampo"
      Tab(1).Control(4)=   "ChkCampo"
      Tab(1).Control(5)=   "GrdCampos"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmWizPrj.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton CmdOperCls 
         Caption         =   "GERAR CLASSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5520
         TabIndex        =   0
         Top             =   5640
         Width           =   3495
      End
      Begin VB.CommandButton cmdScript 
         Caption         =   "Gerar Script de Inicialização"
         Height          =   435
         Left            =   5520
         TabIndex        =   7
         Top             =   5220
         Width           =   3495
      End
      Begin VB.CommandButton CmdChk 
         Caption         =   "Inverter"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   20
         Top             =   6120
         Width           =   840
      End
      Begin VB.CommandButton CmdChk 
         Caption         =   "Nenhuma"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   29
         Top             =   6120
         Width           =   840
      End
      Begin VB.CommandButton CmdChk 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   6120
         Width           =   840
      End
      Begin VB.ListBox LstTabCls 
         Height          =   5235
         ItemData        =   "FrmWizPrj.frx":035E
         Left            =   120
         List            =   "FrmWizPrj.frx":0365
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   900
         Width           =   2535
      End
      Begin VB.TextBox TxtDrvDest 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Left            =   5520
         LinkTimeout     =   30
         TabIndex        =   18
         Top             =   1260
         Width           =   3555
      End
      Begin VB.CommandButton CmdDrv 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procurar Projeto"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7320
         TabIndex        =   17
         Top             =   780
         Width           =   1755
      End
      Begin VB.ListBox LstCampoCls 
         Height          =   5235
         ItemData        =   "FrmWizPrj.frx":0372
         Left            =   2760
         List            =   "FrmWizPrj.frx":0379
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   900
         Width           =   2535
      End
      Begin VB.TextBox TxtSuperClasse 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Left            =   5520
         LinkTimeout     =   30
         TabIndex        =   15
         Top             =   1860
         Width           =   3555
      End
      Begin VB.OptionButton OptInPrj 
         Caption         =   "Sim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   900
         Width           =   735
      End
      Begin VB.OptionButton OptInPrj 
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   900
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CmdCarregaVetor 
         Caption         =   "Carregar Vetor"
         Height          =   375
         Left            =   7800
         TabIndex        =   12
         Top             =   2460
         Width           =   1215
      End
      Begin VB.ComboBox CmbOwner 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.ListBox LstOp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         ItemData        =   "FrmWizPrj.frx":0385
         Left            =   5520
         List            =   "FrmWizPrj.frx":038C
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   3180
         Width           =   3495
      End
      Begin VB.TextBox TxtNmDbObj 
         Height          =   285
         Left            =   5520
         TabIndex        =   9
         Text            =   "XDb"
         Top             =   2580
         Width           =   1935
      End
      Begin VB.ComboBox CmbNotNull 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmWizPrj.frx":03B0
         Left            =   -68520
         List            =   "FrmWizPrj.frx":03BA
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox LstTabForm 
         Height          =   5460
         ItemData        =   "FrmWizPrj.frx":03CA
         Left            =   -74880
         List            =   "FrmWizPrj.frx":03D1
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   480
         Width           =   2500
      End
      Begin VB.ComboBox CmbCtrl 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmWizPrj.frx":03DE
         Left            =   -69960
         List            =   "FrmWizPrj.frx":03F4
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1060
         Width           =   1460
      End
      Begin VB.TextBox TxtCampo 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -71640
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1080
         Width           =   1680
      End
      Begin VB.CheckBox ChkCampo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         Height          =   200
         Left            =   -72240
         TabIndex        =   2
         Top             =   1080
         Value           =   1  'Checked
         Width           =   200
      End
      Begin MSFlexGridLib.MSFlexGrid GrdCampos 
         Bindings        =   "FrmWizPrj.frx":0436
         Height          =   5595
         Left            =   -72360
         TabIndex        =   6
         Top             =   405
         WhatsThisHelpID =   10506
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   9869
         _Version        =   393216
         Rows            =   20
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   315
         BackColor       =   12648447
         BackColorSel    =   12648447
         AllowBigSelection=   0   'False
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmWizPrj.frx":0447
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campo de &Descrição"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   2760
         TabIndex        =   27
         Top             =   660
         Width           =   1650
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Tabelas do Banco de Dados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   660
         Width           =   2175
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Classe Banco de Dados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   5520
         TabIndex        =   25
         Top             =   1620
         Width           =   1875
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incluir no Projeto ? "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   5520
         TabIndex        =   24
         Top             =   660
         Width           =   1650
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Owner :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Propiedade de Banco :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   5520
         TabIndex        =   22
         Top             =   2340
         Width           =   1785
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selecionar Todos os Itens"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   7
         Left            =   5640
         TabIndex        =   21
         Top             =   6180
         Width           =   2085
      End
   End
   Begin MSComDlg.CommonDialog CmDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmWizPrj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
Option Explicit
'**** Form ******
Private ArrId()
Public LinAnt        As Integer
Public Tabela        As String
Public TabName       As String
Public MyPrj         As New PROJETO
Public UserDB        As String
Public MyGrd         As MSGrid
Public DrvLocal      As String
Public CHEstrangeira As String

'**** Classe ******
Public bSelecionado  As Boolean

Dim bComDLL          As Boolean
Dim bTipoQuery       As Boolean
Dim bItensExcluidos  As Boolean
Dim bExiste          As Boolean
Dim bVB6             As Boolean
Dim bGlobalClass     As Boolean
Dim bisDirt          As Boolean

Dim NmDbObj          As String
Dim isODBC           As Boolean
Public Function MontarClasse(Tabela As String, pDscExclusao As String) As Boolean
   Dim ArrChave
   Dim Drv           As String
   Dim CLASSE        As String
   Dim Arq           As String
   Dim CamposSel     As String
   Dim CamposIns     As String
   Dim ChaveOptional As String
   Dim ChaveParamDef As String
   Dim ChaveParam    As String
   Dim Chaves        As String
   Dim AlterChParam  As String
   Dim sAutoIncrement As String
   Dim sChaveMax     As String
   Dim bTimeStamp    As Boolean
   Dim nVarTamanho   As Integer
      
   On Error GoTo Fim
   
   Tabela = UCase(Tabela)
   
   If Trim(Me.TxtDrvDest.Text) = "" Or Right(Me.TxtDrvDest.Text, 3) = "..." Then
      Me.TxtDrvDest.Text = DrvLocal & "..."
      Me.TxtDrvDest.Tag = DrvLocal
   End If
   
   Drv$ = Me.TxtDrvDest.Tag
   CLASSE$ = IIf(Mid(Tabela, 1, 3) = "TB_", Tabela, "TB_" & Tabela)
   Arq = CLASSE & ".cls"
   Call SetHourglass(hWnd)
   Call MakePath(Drv$)
   Call Del(Drv$ & Arq$)
   '   AbrirTxt% = FreeFile()
   Close #1
   Open Drv & Arq For Output As #1
   
   Call MontarClasse_Begin(CLASSE)
   Call DefineIdentity(Tabela, sAutoIncrement, sChaveMax)
   Call DefineChaveMax(Tabela, sChaveMax)
   Call DefineChaveEstrangeira(Tabela, CHEstrangeira)
   Call DefineVarTamanho(Tabela, nVarTamanho, bTimeStamp)
   Call MontarClasse_DefaultProperty1(nVarTamanho)
      
   With XDb.Tables(Tabela)
      Call DefineParametros(Tabela, sAutoIncrement, ArrChave, Chaves, ChaveParam, ChaveOptional, ChaveParamDef, AlterChParam)
      Call DefineStrCampos(Tabela, sAutoIncrement, nVarTamanho, CamposSel, CamposIns)
      Call MontarClasse_Property(Tabela, sAutoIncrement, nVarTamanho, ArrChave)
      Call MontarClasse_DefaultProperty2
      
      Call MontarClasse_QryInsert(Tabela, sAutoIncrement, sChaveMax, CamposIns, ArrChave)
      Call MontarClasse_QryDelete(Tabela, ChaveOptional, ArrChave, bTimeStamp)
      Call MontarClasse_QryUpDate(Tabela, nVarTamanho, sAutoIncrement, ArrChave)
      Call MontarClasse_QrySave(Tabela, sChaveMax)
      Call MontarClasse_QrySelect(Tabela, ChaveOptional, CamposSel, ArrChave)
      Call MontarClasse_GRAVAR
      Call MontarClasse_PESQUISAR(ChaveOptional, ChaveParam)
      Call MontarClasse_POPULA(Tabela, ArrChave)
      Call MontarClasse_LIMPAR(Tabela, ArrChave)
      Call MontarClasse_SALVAR(Tabela, sChaveMax)
      Call MontarClasse_INCLUIR
      Call MontarClasse_EXCLUIR(pDscExclusao, Chaves)
      Call MontarClasse_ALTERAR
      Call MontarClasse_ALTERARCHAVE(Tabela, ArrChave, sAutoIncrement, AlterChParam)
      Call MontarClasse_Initialize
      Call MontarClasse_Terminate
   End With
   Close #1
   
   Call MontarSuperClasse(CLASSE$)
   
   Call SetDefault(hWnd)
   Exit Function
Fim:
   'If Err = 55 Then
   Close #1
   Call ShowError
   MontarClasse = False
   '   MsgBox CStr(Err) & " - " & CStr(Error)
End Function
Private Sub MontarClasse_Begin(CLASSE As String)
   Print #1, "VERSION 1.0 CLASS"
   Print #1, "BEGIN"
   Print #1, "  MultiUse = -1  'True"
   If bVB6 Then
      Print #1, "  Persistable = 0  'NotPersistable"
      Print #1, "  DataBindingBehavior = 0  'vbNone"
      Print #1, "  DataSourceBehavior = 0   'vbNone"
      Print #1, "  MTSTransactionMode = 0   'NotAnMTSObject"
   End If
   Print #1, "END"
   Print #1, "Attribute VB_Name = """ & CLASSE & """"
   Print #1, "Attribute VB_GlobalNameSpace = " & IIf(bGlobalClass, "True", "False")
   Print #1, "Attribute VB_Creatable = True"
   Print #1, "Attribute VB_PredeclaredId = False"
   Print #1, "Attribute VB_Exposed = " & IIf(bGlobalClass, "True", "False")
   Print #1, "Attribute VB_Ext_KEY = """ & "SavedWithClassBuilder"" ,""Yes"""
   If Me.OptInPrj(0) Then
      Print #1, "Attribute VB_Ext_KEY = """ & "Top_Level"" ,""No"""
   Else
      Print #1, "Attribute VB_Ext_KEY = """ & "Top_Level"" ,""Yes"""
   End If
End Sub
Private Sub MontarClasse_DefaultProperty1(nVarTamanho As Integer)
   Dim nSpcVar  As Integer
   
   'Dim MyRs As ADODB.Recordset
   'Set MyRs = XDb.ADOConect.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
   'Set MyRs = XDb.ADOConect.OpenSchema(adSchemaColumns, Array(XDb.dbName, XDb.Tables(Tabela).Owner, Tabela, Empty))
   'For i = 0 To 32
   '   Debug.Print MyRs.Fields(i).Name & " = " & MyRs.Fields(i).Value
   'Next
   
   nSpcVar = nVarTamanho
   nSpcVar = IIf(nSpcVar >= Len(NmDbObj), nSpcVar, Len(NmDbObj))
   nSpcVar = IIf(nSpcVar >= Len("QryInsert"), nSpcVar, Len("QryInsert"))
   If bItensExcluidos Then
      nSpcVar = IIf(nSpcVar >= Len("ItensExcluidos"), nSpcVar, Len("ItensExcluidos"))
   End If
   nSpcVar = nSpcVar + (3 - (nSpcVar Mod 3)) - 1
            
   Print #1, "Option Explicit "
   'Print #1, "Private mvar" & NmDbObj & Space(nSpcVar - Len(NmDbObj)) & IIf(bComDLL, " As DS_BANCO", " As Object ")
   'Print #1, "Private mvarRS" & Space(nSpcVar - Len("RS")) & IIf(bComDLL, " As Recordset", " As Object ")
   Print #1, "Private mvar" & NmDbObj & Space(nSpcVar - Len(NmDbObj)) & " As Object "
   Print #1, "Private mvarRS" & Space(nSpcVar - Len("RS")) & " As Object "
   If bExiste Then
      Print #1, "Private mvarEXISTE" & Space(nSpcVar - Len("EXISTE")) & " As Integer "
   End If
   Print #1,
   Print #1, "Private mvarQryInsert" & Space(nSpcVar - Len("QryInsert")) & " As String"
   Print #1, "Private mvarQryUpDate" & Space(nSpcVar - Len("QryUpDate")) & " As String"
   Print #1, "Private mvarQryDelete" & Space(nSpcVar - Len("QryDelete")) & " As String"
   Print #1, "Private mvarQrySelect" & Space(nSpcVar - Len("QrySelect")) & " As String"
   Print #1, "Private mvarQrySave" & Space(nSpcVar - Len("QrySave")) & " As String"
   Print #1,
   If bItensExcluidos Then
      Print #1, "Private mvarItensExcluidos" & Space(nSpcVar - Len("ItensExcluidos")) & " As Collection"
   End If
   If bTipoQuery Then
      Print #1, "Private mvarTipoQuery" & Space(nSpcVar - Len("TipoQuery")) & " As String"
   End If
   If bisDirt Then
      Print #1, "Private mvarisDirt" & Space(nSpcVar - Len("isDirt")) & " As Boolean"
   End If
   If bItensExcluidos Or bTipoQuery Or bisDirt Then
      Print #1,
   End If
End Sub
Private Sub MontarClasse_DefaultProperty2()
   If bTipoQuery Then
      Print #1, "Public Property Let TipoQuery(ByVal vData As String)"
      Print #1, "    mvarTipoQuery = vData"
      Print #1, "End Property"
      Print #1, "Public Property Get TipoQuery() As String"
      Print #1, "    TipoQuery = mvarTipoQuery"
      Print #1, "End Property"
   End If
   If bItensExcluidos Then
      Print #1, "Public Property Set ItensExcluidos(ByVal vData As Object)"
      Print #1, "    Set mvarItensExcluidos = vData"
      Print #1, "End Property"
      Print #1, "Public Property Get ItensExcluidos() As Collection"
      Print #1, "   If mvarItensExcluidos Is Nothing Then"
      Print #1, "      Set mvarItensExcluidos = New Collection"
      Print #1, "   End If"
      Print #1, "   Set ItensExcluidos = mvarItensExcluidos"
      Print #1, "End Property"
   End If
   If bExiste Then
      Print #1, "Public Property Get EXISTE() As Integer"
      Print #1, "   EXISTE = mvarEXISTE"
      Print #1, "End Property"
   End If
   If bisDirt Then
      Print #1, "Public Property Get isDirt() As Boolean"
      Print #1, "   isDirt = mvarisDirt"
      Print #1, "End Property"
   End If
   Print #1, "Public Property Set " & NmDbObj & "(ByVal vData As Object)"
   Print #1, "   Set mvar" & NmDbObj & " = vData"
   Print #1, "End Property"
   Print #1, "Public Property Let " & NmDbObj & "(ByVal vData As Object)"
   Print #1, "   Set mvar" & NmDbObj & " = vData"
   Print #1, "End Property"
   Print #1, "Public Property Get " & NmDbObj & "() As Object"
   Print #1, "   Set " & NmDbObj & " = mvar" & NmDbObj
   Print #1, "End Property"
   'Print #1, "Public Property Get RS()" & IIf(bComDLL, " As RecordSet", " As Object")
   Print #1, "Public Property Get RS() As Object"
   Print #1, "   Set RS = mvarRS"
   Print #1, "End Property"
End Sub
Private Sub MontarClasse_Property(pTabela As String, pAutoIncrement As String, pVarTamanho As Integer, pArrChave)
   Dim i As Integer
   Dim j As Integer
   Dim Txt As String
   Dim TpVar As String
   Dim nSpcVar As Integer
   
   With XDb.Tables(pTabela)
      For i = 1 To .Fields.Count
         If Not .Fields(i).isSys Then
            Select Case GrpTipoCampo(.Fields(i).Tipo)
               Case 1: TpVar = " As Double" '* Numérico
               Case 2: TpVar = " As String" '* Data
               Case 3: TpVar = " As String" '* Caracter
               Case 4: TpVar = " As Object" '* Caracter
            End Select
         
            'If .Fields(i).NOME <> "TIMESTAMP" And .Fields(i).NOME <> pAutoIncrement Then
            If .Fields(i).NOME <> pAutoIncrement Then
               If GrpTipoCampo(.Fields(i).Tipo) = 4 Then
                  Print #1, "Public Property Set " & .Fields(i).NOME & "(ByVal vData" & TpVar & ")"
               Else
                  Print #1, "Public Property Let " & .Fields(i).NOME & "(ByVal vData" & TpVar & ")"
               End If
               If .Fields(i).NOME <> "ALTERSTAMP" Then
                  Print #1, "   If Not mvarisDirt Then mvarisDirt = (mvar" & .Fields(i).NOME & " <> vData)"
                  If GrpTipoCampo(.Fields(i).Tipo) = 4 Then
                     Print #1, "   Set mvar" & .Fields(i).NOME & " = vData"
                  Else
                     Print #1, "   mvar" & .Fields(i).NOME & " = vData"
                  End If
               Else
                  Print #1, "   Dim Sql As String"
                  Print #1, " "
                  Print #1, "   Sql = """ & "Update " & pTabela & " Set ALTERSTAMP=" & """" & " & vData & VbNewLine "
                  Print #1, "   Sql = Sql & """ & " Where """ & " & VbNewLine "
                  If Not IsEmpty(ArrId(0)) Then
                     For j = LBound(pArrChave) To UBound(pArrChave) - 1
                        nSpcVar = pVarTamanho - Len(pArrChave(j)) - IIf(j = 0, -2, 2)
                        nSpcVar = IIf(nSpcVar < 0, 0, nSpcVar)
                        Txt = "   Sql = Sql & """ & IIf(j = 0, Space(1), " And ") & pArrChave(j) & Space(nSpcVar) & " = """
                        Select Case GrpTipoCampo(.Fields(pArrChave(j)).Tipo)
                           Case 1: Txt = Txt & " & SqlNum(mvar" & pArrChave(j) & ")"
                           Case 2: Txt = Txt & " & SqlDate(mvar" & pArrChave(j) & ", eSysDate.Data_Hora)"
                           Case 3: Txt = Txt & " & SqlStr(mvar" & pArrChave(j) & ")"
                        End Select
                        Print #1, Txt & " & VbNewLine "
                     Next
                  End If
                  Print #1, "   If Not mvar" & NmDbObj & " Is Nothing Then"
                  Print #1, "      If mvar" & NmDbObj & ".Conectado Then"
                  Print #1, "         If mvar" & NmDbObj & ".Executa(Sql, True) Then"
                  Print #1, "            mvar" & .Fields(i).NOME & " = vData"
                  Print #1, "         End If"
                  Print #1, "      End If"
                  Print #1, "   End If"
               End If
               Print #1, "End Property"
            End If
                        
            Print #1, "Public Property Get " & .Fields(i).NOME & "()" & TpVar
            If GrpTipoCampo(.Fields(i).Tipo) = 4 Then
               Print #1, "   Set " & .Fields(i).NOME & " = mvar" & .Fields(i).NOME
            Else
               Print #1, "   " & .Fields(i).NOME & " = mvar" & .Fields(i).NOME
            End If
            Print #1, "End Property"
         End If
      Next
   End With
End Sub
Private Sub MontarClasse_QryInsert(pTabela As String, pAutoIncrement As String, pChaveMax As String, pCamposIns As String, pArrChave)
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim Txt As String
   Dim bAnd As Boolean
   Dim bAutoId As Boolean
   
   If pChaveMax <> "" Then bAutoId = (GrpTipoCampo(XDb.Tables(pTabela).Fields(pChaveMax).Tipo) = 1)
   With XDb.Tables(pTabela)
      '* QryInsert
      Print #1, "Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String"
      Print #1, "   Dim Sql As String"
      Print #1, " "
      Print #1, "   Sql = """ & "Insert Into " & pTabela & " (" & pCamposIns & ") """ & " & VbNewLine "
      If pAutoIncrement <> "" Or pChaveMax <> "" Then
         Print #1, "   Sql = Sql & """ & " Output Inserted.*""" & " & VbNewLine "
         Print #1, "   Sql = Sql & """ & " Select "
         j = 0
         For i = 1 To .Fields.Count
            If Not .Fields(i).isSys And pAutoIncrement <> .Fields(i).NOME Then 'And sChaveMax <> .Fields(i).NOME Then
               If .Fields(i).NOME = "TIMESTAMP" Or .Fields(i).NOME = "ALTERSTAMP" Or .Fields(i).NOME = pChaveMax Then
                  If .Fields(i).NOME = pChaveMax And GrpTipoCampo(XDb.Tables(pTabela).Fields(i).Tipo) = 1 Then
                     Txt = ""
                     Txt = Txt & "   If pAutoId Then" & vbNewLine
                     Txt = Txt & "      Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ")
                     If UBound(pArrChave) = 1 Then
                        Txt = Txt & """(Select isNull(Max(" & pChaveMax & "),0)+1 From " & pTabela & ")""" & vbNewLine
                     Else
                        If Not IsEmpty(ArrId(0)) Then
                           Txt = Txt & """(Select isNull(Max(" & pChaveMax & "),0)+1 From " & pTabela & " Where "
                           bAnd = False
                           For k = LBound(pArrChave) To UBound(pArrChave) - 1
                              If pChaveMax <> pArrChave(k) Then
                                 Txt = Txt & IIf(bAnd, " & "" And ", "") & pArrChave(k) & " = """
                                 bAnd = True
                                 Select Case GrpTipoCampo(.Fields(pArrChave(k)).Tipo)
                                    Case 1: Txt = Txt & " & SqlNum(mvar" & pArrChave(k) & ")"
                                    Case 2: Txt = Txt & " & SqlDate(mvar" & pArrChave(k) & ", eSysDate.Data_Hora)"
                                    Case 3: Txt = Txt & " & SqlStr(mvar" & pArrChave(k) & ")"
                                 End Select
                              End If
                           Next
                        End If
                        Txt = Txt & " & "")"" & vbNewLine" & vbNewLine
                     End If
                     Txt = Txt & "   Else" & vbNewLine
                     Txt = Txt & "      Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ")
                     Select Case GrpTipoCampo(.Fields(i).Tipo)
                        Case 1: Txt = Txt & "SqlNum(mvar" & .Fields(i).NOME & ")"
                        Case 2: Txt = Txt & "SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora)"
                        Case 3: Txt = Txt & "SqlStr(mvar" & .Fields(i).NOME & ")"
                     End Select
                     Txt = Txt & " & VbNewLine " & vbNewLine
                     Txt = Txt & "   End If"
                     Print #1, Txt
                  Else
                     If .Fields(i).NOME = "TIMESTAMP" Or .Fields(i).NOME = "ALTERSTAMP" Then
                        Txt = ""
                        Txt = Txt & "   If pSinc Then " & vbNewLine
                        Txt = Txt & "      Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ")
                        If .Fields(i).NOME = "TIMESTAMP" Then Txt = Txt & " SqlDate(mvarTIMESTAMP) & vbNewLine" & vbNewLine
                        If .Fields(i).NOME = "ALTERSTAMP" Then Txt = Txt & " SqlNum(mvarALTERSTAMP) & vbNewLine" & vbNewLine
                        Txt = Txt & "   Else" & vbNewLine
                        Txt = Txt & "      Sql = Sql & " & IIf(j = 0, Space(1), """, ")
                        If .Fields(i).NOME = "TIMESTAMP" Then Txt = Txt & "GetDate()"" & vbNewLine" & vbNewLine
                        If .Fields(i).NOME = "ALTERSTAMP" Then Txt = Txt & "1"" & vbNewLine" & vbNewLine
                        Txt = Txt & "   End If"
                        Print #1, Txt
                     
                     Else
                        Txt = "   Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ")
                        Select Case GrpTipoCampo(.Fields(i).Tipo)
                           Case 1: Txt = Txt & "SqlNum(mvar" & .Fields(i).NOME & ")"
                           Case 2: Txt = Txt & "SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora)"
                           Case 3: Txt = Txt & "SqlStr(mvar" & .Fields(i).NOME & ")"
                        End Select
                        Print #1, Txt & " & VbNewLine "
                     End If
                  End If
                  
               Else
                  If GrpTipoCampo(.Fields(i).Tipo) > 0 And GrpTipoCampo(.Fields(i).Tipo) <= 3 Then
                     Txt = "   Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ")
                     If InStr(CHEstrangeira, .Fields(i).NOME) <> 0 And XDb.Tables(pTabela).Fields(i).IsNull Then
                        Select Case GrpTipoCampo(.Fields(i).Tipo)
                           Case 1: Txt = Txt & "IIf(mvar" & .Fields(i).NOME & " = 0, ""Null"", SqlNum(mvar" & .Fields(i).NOME & "))"
                           Case 2: Txt = Txt & "IIf(mvar" & .Fields(i).NOME & " = '', ""Null"", SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora))"
                           Case 3: Txt = Txt & "IIf(mvar" & .Fields(i).NOME & " = '', ""Null"", SqlStr(mvar" & .Fields(i).NOME & "))"
                        End Select
                     Else
                        Select Case GrpTipoCampo(.Fields(i).Tipo)
                           Case 1: Txt = Txt & "SqlNum(mvar" & .Fields(i).NOME & ")"
                           Case 2: Txt = Txt & "SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora)"
                           Case 3: Txt = Txt & "SqlStr(mvar" & .Fields(i).NOME & ")"
                        End Select
                     End If
                     Print #1, Txt & " & VbNewLine "
                  End If
               End If
               
               j = 1
            End If
         Next
      Else
         Print #1, "   Sql = Sql & """ & " Values " & """" & " & VbNewLine "
         Print #1, "   Sql = Sql & """ & "(" & """" & " & VbNewLine "
         j = 0
         For i = 1 To .Fields.Count
            If Not .Fields(i).isSys And pAutoIncrement <> .Fields(i).NOME Then
               Txt = "   Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ")
               If .Fields(i).NOME = "TIMESTAMP" Or .Fields(i).NOME = "ALTERSTAMP" Then
                  If .Fields(i).NOME = "TIMESTAMP" Then Txt = Txt & """" & "GetDate()"""
                  If .Fields(i).NOME = "ALTERSTAMP" Then Txt = Txt & """" & "1"""
                  'If .Fields(i).NOME = "TIMESTAMP" Then Txt = Txt & "GetDate()" '"SqlStr(mvar" & NmDbObj & ".Sysdate(eSysDate.Data_Hora))"
                  'If .Fields(i).NOME = "ALTERSTAMP" Then Txt = Txt & "SqlNum(1)"
               Else
                  Select Case GrpTipoCampo(.Fields(i).Tipo)
                     Case 1: Txt = Txt & "SqlNum(mvar" & .Fields(i).NOME & ")"
                     Case 2: Txt = Txt & "SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora)"
                     Case 3: Txt = Txt & "SqlStr(mvar" & .Fields(i).NOME & ")"
                  End Select
               End If
               Print #1, Txt & " & VbNewLine "
               j = 1
            End If
         Next
         Print #1, "   Sql = Sql & """ & ")" & """" & " & VbNewLine "
      End If
      Print #1, ""
      Print #1, "   mvarQryInsert = Sql"
      Print #1, "   QryInsert = mvarQryInsert"
      Print #1, "End Property"
   End With
End Sub
Private Sub MontarClasse_QryDelete(pTabela As String, pChaveOptional As String, pArrChave, pTimeStamp As Boolean)
   Dim i       As Integer
   Dim Txt     As String
   
   With XDb.Tables(pTabela)
      '* QryDelete
      Print #1, "Public Property Get QryDelete(" & pChaveOptional & IIf(pChaveOptional = "", "", ", ") & "Optional Ch_WHERE) As String"
      Print #1, "   Dim Sql As String"
      Print #1, " "
      '         For i = LBound(pArrChave) To UBound(pArrChave) - 1
      '            If i = LBound(pArrChave) Then
      '               Txt = "   If isMissing(Ch_" & pArrChave(i) & ")"
      '            Else
      '               Txt = Txt & " And isMissing(Ch_" & pArrChave(i) & ")"
      '            End If
      '         Next
      '         Txt = Txt & " Then"
      '         Print #1, Txt
      '         For i = LBound(pArrChave) To UBound(pArrChave) - 1
      '            Print #1, "      If Trim(mvar" & pArrChave(i) & ") = """" Then Exit Property"
      '         Next
      '         For i = LBound(pArrChave) To UBound(pArrChave) - 1
      '            Print #1, "      Ch_" & pArrChave(i) & " = mvar" & pArrChave(i)
      '         Next
      '         Print #1, "   End If"
      Print #1, "   Sql = """ & "Delete From " & pTabela & """" & " & VbNewLine "
      Print #1, "   Sql = Sql & """ & " Where" & """" & " & VbNewLine "
      If Not IsEmpty(ArrId(0)) Then
         Txt = "   If IsMissing(Ch_WHERE) "
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            Txt = Txt & " And IsMissing(Ch_" & pArrChave(i) & ") "
         Next
         Txt = Txt & " Then "
         Print #1, Txt
         
         Txt = ""
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            Txt = "      Sql = Sql & """ & " " & pArrChave(i) & " = """
            Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
               Case 1: Txt = Txt & " & SqlNum(mvar" & pArrChave(i) & ")"
               Case 2: Txt = Txt & " & SqlDate(mvar" & pArrChave(i) & ", eSysDate.Data_Hora)"
               Case 3: Txt = Txt & " & SqlStr(mvar" & pArrChave(i) & ")"
            End Select
            Txt = Txt & " & "" AND """
            Print #1, Txt & " & VbNewLine "
         Next
         
         Print #1, "   Else "
         
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            Txt = "      If Not isMissing(Ch_" & pArrChave(i) & ") Then "
            Txt = Txt & "Sql = Sql & """ & " " & pArrChave(i) & " = """
            Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
               Case 1: Txt = Txt & " & SqlNum(Cstr(Ch_" & pArrChave(i) & "))"
               Case 2: Txt = Txt & " & SqlDate(Cstr(Ch_" & pArrChave(i) & "), eSysDate.Data_Hora)"
               Case 3: Txt = Txt & " & SqlStr(Cstr(Ch_" & pArrChave(i) & "))"
            End Select
            Txt = Txt & " & "" AND """
            Print #1, Txt & " & VbNewLine "
         Next
         Print #1, "      If Not IsMissing(Ch_WHERE) Then"
         Print #1, "         If Trim(Ch_WHERE) = " & """""" & " And Right(Trim(Replace(Sql, vbNewLine, " & """""" & ")), Len(" & """" & "Where" & """" & ")) = " & """" & "Where" & """" & " Then"
         Print #1, "            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len("" Where  "")))"
         Print #1, "         Else"
         Print #1, "            Sql = Sql & Ch_WHERE"
         Print #1, "         End If"
         Print #1, "         Sql = Sql & " & """ And """ & " & VbNewLine"
         Print #1, "      End If"
         Print #1, "   End If "
      End If
      Print #1, "   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len("" AND  "")))"
      Print #1, "   mvarQryDelete = Sql"
      
      If pTimeStamp And pTabela <> "DELETEDROWS" Then
         Print #1, ""
         Print #1, "   Dim MyDelRow As New TB_DELETEDROWS"
         Print #1, "   Dim sTag     As String"
         Print #1, ""
         Print #1, "   sTag = """
         For i = 1 To .Fields.Count
            If Not .Fields(i).isSys And .Fields(i).NOME <> "ALTERSTAMP" And .Fields(i).NOME <> "TIMESTAMP" Then
               Print #1, "   sTag = sTag & ""|" & .Fields(i).NOME & " = "" & mvar" & .Fields(i).NOME
            End If
         Next
         Print #1, "   sTag = sTag & ""|Where = "" & IIf(IsMissing(Ch_WHERE), """", Ch_WHERE)"
         Print #1, "   sTag = sTag & ""|"""
         Print #1, ""
         Print #1, "   MyDelRow.Query = Sql & """ & ";"""
         Print #1, "   MyDelRow.Tag = sTag"
         Print #1, "   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & """ & ";"""
         Print #1, "   Set MyDelRow = Nothing"
         Print #1, ""
         'Print #1, "   mvarQryDelete = mvarQryDelete & vbNewLine"
         'Print #1, "   mvarQryDelete = mvarQryDelete & "; Insert; Into; DELETEDROWS; ""
         'Print #1, "   mvarQryDelete = mvarQryDelete & "(QUERY, SITQUERY)"
         'Print #1, "   mvarQryDelete = mvarQryDelete & "; Value; ""
         'Print #1, "   mvarQryDelete = mvarQryDelete & "( '" & Sql & "', 0)"
      End If
      Print #1, "   QryDelete = mvarQryDelete"
      Print #1, "End Property"
   End With
End Sub
Private Sub MontarClasse_QryUpDate(pTabela As String, pVarTamanho As Integer, pAutoIncrement As String, pArrChave)
   Dim i       As Integer
   Dim j       As Integer
   Dim k       As Integer
   Dim nSpcVar As Integer
   Dim Txt     As String
   Dim isKey   As Boolean
   Dim UboundChave    As Integer
   Dim LboundChave    As Integer
   
   With XDb.Tables(pTabela)
      '* QryUpDate
      Print #1, "Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String"
      Print #1, "   Dim Sql As String"
      Print #1, " "
      Print #1, "   Sql = """ & "Update " & pTabela & " Set " & """" & " & VbNewLine "
      j = 0
      nSpcVar = pVarTamanho
      For i = 1 To .Fields.Count
         If Not .Fields(i).isSys And pAutoIncrement <> .Fields(i).NOME Then
            isKey = False
            UboundChave = 0
            LboundChave = 0
            If IsEmpty(pArrChave) Then
               If Not pArrChave Is Nothing Then
                  UboundChave = UBound(pArrChave)
                  LboundChave = LBound(pArrChave)
               End If
            End If
            If UboundChave <> .Fields.Count Then
               For k = LboundChave To UboundChave - 1
                  If pArrChave(k) = .Fields(i).NOME Then
                     isKey = True
                     Exit For
                  End If
               Next
            End If
            If Not isKey Then
               If GrpTipoCampo(.Fields(i).Tipo) <> 4 Then
                  If .Fields(i).NOME = "TIMESTAMP" Or .Fields(i).NOME = "ALTERSTAMP" Then
                     Txt = ""
                     Txt = Txt & "   If pSinc Then " & vbNewLine
                     Txt = Txt & "      Sql = Sql & """ & IIf(j = 0, Space(1), ", ") & .Fields(i).NOME & Space(nSpcVar - (Len(.Fields(i).NOME))) & " = """
                     If .Fields(i).NOME = "TIMESTAMP" Then Txt = Txt & " & SqlDate(mvarTIMESTAMP) & vbNewLine" & vbNewLine
                     If .Fields(i).NOME = "ALTERSTAMP" Then Txt = Txt & " & SqlNum(mvarALTERSTAMP) & vbNewLine" & vbNewLine
                     Txt = Txt & "   Else" & vbNewLine
                     Txt = Txt & "      Sql = Sql & """ & IIf(j = 0, Space(1), ", ") & .Fields(i).NOME & Space(nSpcVar - (Len(.Fields(i).NOME))) & " = "
                     If .Fields(i).NOME = "TIMESTAMP" Then Txt = Txt & "GetDate()"" & vbNewLine" & vbNewLine
                     If .Fields(i).NOME = "ALTERSTAMP" Then Txt = Txt & "1"" & vbNewLine" & vbNewLine
                     Txt = Txt & "   End If"
                     Print #1, Txt
                  Else
                     Txt = "   Sql = Sql & """ & IIf(j = 0, Space(1), " , ") & .Fields(i).NOME & Space(nSpcVar - (Len(.Fields(i).NOME))) & " = """
                     If InStr(CHEstrangeira, .Fields(i).NOME) <> 0 And XDb.Tables(pTabela).Fields(i).IsNull Then
                        Select Case GrpTipoCampo(.Fields(i).Tipo)
                           Case 1: Txt = Txt & " & IIf(mvar" & .Fields(i).NOME & " = 0, ""Null"", SqlNum(mvar" & .Fields(i).NOME & "))"
                           Case 2: Txt = Txt & " & IIf(mvar" & .Fields(i).NOME & " = '', ""Null"", SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora))"
                           Case 3: Txt = Txt & " & IIf(mvar" & .Fields(i).NOME & " = '', ""Null"", SqlStr(mvar" & .Fields(i).NOME & "))"
                        End Select
                     Else
                        Select Case GrpTipoCampo(.Fields(i).Tipo)
                           Case 1: Txt = Txt & " & SqlNum(mvar" & .Fields(i).NOME & ")"
                           Case 2: Txt = Txt & " & SqlDate(mvar" & .Fields(i).NOME & ", eSysDate.Data_Hora)"
                           Case 3: Txt = Txt & " & SqlStr(mvar" & .Fields(i).NOME & ")"
                        End Select
                     End If
                     Print #1, Txt & " & VbNewLine "
                  End If
                  
               End If
               j = 1
            End If
         End If
      Next
      
      Print #1, "   If Not mvarXDb.ExisteReg(" & """Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= '" & pTabela & "'"" ) Then"
      Print #1, "      Sql = Sql & """ & " Output Inserted.*""" & " & VbNewLine "
      Print #1, "   End If"
      Print #1, "   Sql = Sql & """ & " Where """ & " & VbNewLine "
      If Not IsEmpty(ArrId(0)) Then
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            nSpcVar = pVarTamanho - Len(pArrChave(i)) - IIf(i = 0, -2, 2)
            nSpcVar = IIf(nSpcVar < 0, 0, nSpcVar)
            Txt = "   Sql = Sql & """ & IIf(i = 0, Space(1), " And ") & pArrChave(i) & Space(nSpcVar) & " = """
            Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
               Case 1: Txt = Txt & " & SqlNum(mvar" & pArrChave(i) & ")"
               Case 2: Txt = Txt & " & SqlDate(mvar" & pArrChave(i) & ", eSysDate.Data_Hora)"
               Case 3: Txt = Txt & " & SqlStr(mvar" & pArrChave(i) & ")"
            End Select
            Print #1, Txt & " & VbNewLine "
         Next
      End If
      nSpcVar = pVarTamanho
      
      Print #1, ""
      Print #1, "   mvarQryUpDate = Sql"
      Print #1, "   QryUpDate = mvarQryUpDate"
      Print #1, "End Property"
      
   End With
End Sub
Private Sub MontarClasse_QrySave(pTabela As String, pChaveMax As String)
   Dim bAutoId As Boolean
   If pChaveMax <> "" Then bAutoId = (GrpTipoCampo(XDb.Tables(pTabela).Fields(pChaveMax).Tipo) = 1)
      
   With XDb.Tables(pTabela)
      '* QrySave
      Print #1, "Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String"
      Print #1, "   Dim Sql As String"
      Print #1, ""
      Print #1, "   Sql = """ & " If Exists(""" & " & Me.QrySelect() & """ & ") """ & " & VbNewLine "
      Print #1, "   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)"
      Print #1, "   Sql = Sql & """ & " Else """ & " & VbNewLine "
      Print #1, "   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)"
      Print #1, " "
      Print #1, "   mvarQrySave = Sql"
      Print #1, "   QrySave = mvarQrySave"
      Print #1, "End Property"
   End With
End Sub
Private Sub MontarClasse_QrySelect(pTabela As String, pChaveOptional As String, pCamposSel As String, pArrChave)
   Dim i As Integer
   Dim Txt As String
   
   With XDb.Tables(pTabela)
      '* QrySelect
      Print #1, "Public Property Get QrySelect(" & pChaveOptional & IIf(pChaveOptional = "", "", ", ") & "Optional Ch_WHERE, Optional Ch_ORDERBY) As String"
      Print #1, "   Dim Sql As String"
      Print #1, " "
      Print #1, "   Sql = """ & "Select " & pCamposSel
      Print #1, "   Sql = Sql &""" & " From " & pTabela & """" & " & VbNewLine "
      Print #1, "   Sql = Sql & """ & " Where """ & " & VbNewLine "
      If Not IsEmpty(ArrId(0)) Then
         Txt = "   If IsMissing(Ch_WHERE) "
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            Txt = Txt & " And IsMissing(Ch_" & pArrChave(i) & ") "
         Next
         Txt = Txt & " Then "
         Print #1, Txt
         
         Txt = ""
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            Txt = "      Sql = Sql & """ & " " & pArrChave(i) & " = """
            Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
               Case 1: Txt = Txt & " & SqlNum(mvar" & pArrChave(i) & ")"
               Case 2: Txt = Txt & " & SqlDate(mvar" & pArrChave(i) & ", eSysDate.Data_Hora)"
               Case 3: Txt = Txt & " & SqlStr(mvar" & pArrChave(i) & ")"
            End Select
            Txt = Txt & " & "" AND """
            Print #1, Txt & " & VbNewLine "
         Next
         
         Print #1, "   Else "
         
         For i = LBound(pArrChave) To UBound(pArrChave) - 1
            Txt = "      If Not isMissing(Ch_" & pArrChave(i) & ") Then "
            Txt = Txt & "Sql = Sql & """ & " " & pArrChave(i) & " = """
            Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
               Case 1: Txt = Txt & " & SqlNum(Cstr(Ch_" & pArrChave(i) & "))"
               Case 2: Txt = Txt & " & SqlDate(Cstr(Ch_" & pArrChave(i) & "), eSysDate.Data_Hora)"
               Case 3: Txt = Txt & " & SqlStr(Cstr(Ch_" & pArrChave(i) & "))"
            End Select
            Txt = Txt & " & "" AND """
            Print #1, Txt & " & VbNewLine "
         Next
         Print #1, "      If Not IsMissing(Ch_WHERE) Then"
         Print #1, "         If Trim(Ch_WHERE) = " & """""" & " And Right(Trim(Replace(Sql, vbNewLine, " & """""" & ")), Len(" & """" & "Where" & """" & ")) = " & """" & "Where" & """" & " Then"
         Print #1, "            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len("" Where  "")))"
         Print #1, "         Else"
         Print #1, "            Sql = Sql & Ch_WHERE"
         Print #1, "         End If"
         Print #1, "         Sql = Sql & " & """ And """ & " & VbNewLine"
         Print #1, "      End If"
         Print #1, "   End If "
      End If
      Print #1, "   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len("" AND  "")))"
      Print #1, "   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & "" Order By "" & Ch_ORDERBY "
      Print #1, ""
      Print #1, "   mvarQrySelect = Sql"
      Print #1, "   QrySelect = mvarQrySelect"
      Print #1, "End Property"
   End With
End Sub
Private Sub MontarClasse_GRAVAR()
   '* GRAVAR
   If bComDLL And bExiste Then
      Print #1, "Public Function Gravar(Optional ByVal ExibeResult = True) As Variant"
      Print #1, "   Dim Result"
      Print #1, "   Select Case mvarEXISTE"
      Print #1, "      Case ALTERACAO: Result = Alterar()"
      Print #1, "      Case INCLUSAO: Result = Incluir()"
      Print #1, "   End Select"
      Print #1, "   If Not ExibeResult Then Exit Function"
      Print #1, "   If Result = FOUND Then"
      Print #1, "      'Call ExibirAviso(LoadMsg(34), LoadMsg(57))"
      Print #1, "   Else"
      Print #1, "      Call ExibirAviso(LoadMsg(48), LoadMsg(57))"
      Print #1, "   End If"
      Print #1, "End Function"
   End If
End Sub
Private Sub MontarClasse_PESQUISAR(pChaveOptional As String, pChaveParam As String)
   '* PESQUISAR
   Print #1, "Public Function Pesquisar(" & pChaveOptional & IIf(pChaveOptional = "", "", ", ") & "Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean" '& IIf(bComDLL, "Integer", "Boolean")
   Print #1, "   Dim Sql     As String"
   Print #1, "   Dim bExiste As Boolean"
   
   Print #1,
   Print #1, "   Sql = QrySelect(" & pChaveParam & IIf(pChaveOptional = "", "", ", ") & "Ch_WHERE, Ch_ORDERBY)"
   If bComDLL Then
      Print #1, "   bExiste = mvar" & NmDbObj & ".AbreTabela(Sql, mvarRS)"
   Else
      Print #1, "   Set mvarRS = mvar" & NmDbObj & ".OpenRecordset(Sql, dbOpenSnapshot, dbExecDirect)"
      Print #1, "   bExiste = True"
   End If
   If bisDirt Then
      Print #1, "   mvarisDirt = False"
   End If
   Print #1, "   With mvarRS"
   Print #1, "      If bExiste Then bExiste = Not .EOF"
   Print #1, "      If bExiste Then"
   Print #1, "         Me.Popula"
   Print #1, "         Pesquisar = True"
   If bTipoQuery Then
      Print #1, "         mvarTipoQuery = ""A"""
   End If
   Print #1, "      Else"
   Print #1, "         Pesquisar = False"
   If bTipoQuery Then
      Print #1, "         mvarTipoQuery = ""I"""
   End If
   Print #1, "      End If"
   Print #1, "   End With"
   Print #1, "   Exit Function"
   Print #1, "PesquisarErr:"
   Print #1, "    call ShowError(Sql)"
   Print #1, "    Pesquisar = False"
   Print #1, "End Function"
End Sub
Private Sub MontarClasse_POPULA(pTabela As String, pArrChave)
   Dim i       As Integer
   Dim j       As Integer
   Dim Txt     As String
   Dim isKey   As Boolean
   
   With XDb.Tables(pTabela)
      '* POPULA
      Print #1, "Public Sub Popula(Optional pRcSet)"
      Print #1, "   If IsMissing(pRcSet) Then Set pRcSet = mvarRS"
      Print #1, "   With pRcSet"
      j = 0
      For i = 1 To .Fields.Count
         If Not .Fields(i).isSys Then
            If Not IsEmpty(ArrId(0)) Then
               isKey = InArray(.Fields(i).NOME, pArrChave) And (j = 0)
            Else
               isKey = False
            End If
            Select Case GrpTipoCampo(.Fields(i).Tipo)
               Case 1: Txt = "      mvar" & .Fields(i).NOME & " = XVal(!" & .Fields(i).NOME & " & """")"
               Case 2: Txt = "      mvar" & .Fields(i).NOME & " = xDate(!" & .Fields(i).NOME & " & """"" & ", True)"
               Case 3: Txt = "      mvar" & .Fields(i).NOME & " = !" & .Fields(i).NOME & " & """""
            End Select
            Print #1, Txt
         End If
         j = 1
      Next
      Print #1, "   End With"
      Print #1, "   mvarisDirt = False"
      Print #1, "End Sub"
   End With
End Sub
Private Sub MontarClasse_LIMPAR(pTabela As String, pArrChave)
   Dim i       As Integer
   Dim j       As Integer
   Dim Txt     As String
   Dim isKey   As Boolean
   
   With XDb.Tables(pTabela)
      '* POPULA
      Print #1, "Public Sub Limpar()"
      j = 0
      For i = 1 To .Fields.Count
         If Not .Fields(i).isSys Then
            If Not IsEmpty(ArrId(0)) Then
               isKey = InArray(.Fields(i).NOME, pArrChave) And (j = 0)
            Else
               isKey = False
            End If
            Select Case GrpTipoCampo(.Fields(i).Tipo)
               Case 1: Txt = "   mvar" & .Fields(i).NOME & " = 0"
               Case 2: Txt = "   mvar" & .Fields(i).NOME & " = " & """"""
               Case 3: Txt = "   mvar" & .Fields(i).NOME & " = " & """"""
            End Select
            Print #1, Txt
         End If
         j = 1
      Next
      Print #1, ""
      Print #1, "   On Error Resume Next"
      Print #1, "   Call Class_Initialize"
      Print #1, "End Sub"
   End With
End Sub
Private Sub MontarClasse_SALVAR(pTabela As String, pChaveMax As String)
   '* SALVAR
   If bComDLL Then
      Print #1, "Public Function Salvar(Optional ComCOMMIT As Boolean = True, Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As Boolean"
      Print #1, "   Salvar = mvar" & NmDbObj & ".Executa(Me.QrySave(pAutoId:=pAutoId, pSinc:=pSinc), ComCOMMIT)"
      Print #1, ""
      Print #1, "   On Error Resume Next"
      Print #1, "   Call Popula(mvar" & NmDbObj & ".ADORs)"
      Print #1, "End Function"
   End If
End Sub
Private Sub MontarClasse_INCLUIR()
   '* INCLUIR
   If bComDLL Then
      Print #1, "Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean"
      Print #1, "   Incluir = mvar" & NmDbObj & ".Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)"
      Print #1, ""
      Print #1, "   On Error Resume Next"
      Print #1, "   Call Popula(mvar" & NmDbObj & ".ADORs)"
      
      If bTipoQuery Then
         Print #1, "   mvarTipoQuery = IIf(Incluir, ""A"", mvarTipoQuery)"
      End If
   Else
      Print #1, "Public Function Incluir() As Boolean"
      Print #1, "   Dim Sql As String"
      Print #1, "   On Error GoTo Fim"
      Print #1, "   Sql = QryInsert"
      Print #1, "   mvar" & NmDbObj & ".Execute Sql, dbExecDirect"
      Print #1, "   Incluir = True"
      If bTipoQuery Then
         Print #1, "   mvarTipoQuery = ""A"""
      End If
      Print #1, "Exit Function"
      Print #1, "Fim:"
      Print #1, "   Incluir = False"
      Print #1, "   msgInformacao " & """" & "Problema na inclusão." & """" & " & vbNewLine & Errors(0).Description"
   End If
   Print #1, "End Function"
End Sub
Private Sub MontarClasse_EXCLUIR(pDscExclusao As String, pChaves As String)
   '* EXCLUIR
   If bComDLL Then
      Print #1, "Public Function Excluir(Optional ComCOMMIT = False) As Boolean"
      pDscExclusao = IIf(pDscExclusao = "", "", ", mvar" & pDscExclusao)
      '            If DscExclusao = "" Then
      Print #1, "   Excluir = mvar" & NmDbObj & ".Executa(Me.QryDelete(" & pChaves & "), ComCOMMIT)"
      '            Else
      '               Print #1, "      Incluir = mvar" & NmDbObj & ".Executa(Me.QryInsert)"
      '               Print #1, "      Incluir = mvar" & NmDbObj & ".Executa(Me.QryInsert)"
      '               Print #1, "   End if"
      '            End If
   Else
      Print #1, "Public Function Excluir() As Boolean"
      Print #1, "   Dim Sql As String"
      Print #1, "   On Error GoTo Fim"
      Print #1, "   Sql = QryDelete(" & pChaves & ")"
      Print #1, "   mvar" & NmDbObj & ".Execute Sql, dbExecDirect"
      Print #1, "   Excluir = True"
      Print #1, "Exit Function"
      Print #1, "Fim:"
      Print #1, "   Excluir = False"
      Print #1, "   msgInformacao " & """" & "Problema na exclusão." & """" & " & vbNewLine & Errors(0).Description"
   End If
   Print #1, "End Function"
End Sub
Private Sub MontarClasse_ALTERAR()
   '* ALTERAR
   If bComDLL Then
      Print #1, "Public Function Alterar(Optional ComCOMMIT = False) As Boolean"
      Print #1, "   Alterar =  mvar" & NmDbObj & ".Executa(Me.QryUpDate, ComCOMMIT)"
      Print #1, ""
      Print #1, "   On Error Resume Next"
      Print #1, "   Call Popula(mvar" & NmDbObj & ".ADORs)"
      
      If bTipoQuery Then
         Print #1, "   mvarTipoQuery = IIf(Alterar, ""A"", mvarTipoQuery)"
      End If
   Else
      Print #1, "Public Function Alterar() As Boolean"
      Print #1, "   Dim Sql As String"
      Print #1, "   On Error GoTo Fim"
      Print #1, "   Sql = QryUpDate"
      Print #1, "   mvar" & NmDbObj & ".Execute Sql, dbExecDirect"
      Print #1, "   Alterar = True"
      If bTipoQuery Then
         Print #1, "   mvarTipoQuery = ""A"""
      End If
      Print #1, "Exit Function"
      Print #1, "Fim:"
      Print #1, "   Alterar = False"
      Print #1, "   msgInformacao " & """" & "Problema na atualização." & """" & " & vbNewLine & Errors(0).Description"
   End If
   Print #1, "End Function"
End Sub
Private Sub MontarClasse_ALTERARCHAVE(pTabela As String, pArrChave, pAutoIncrement As String, pAlterChParam As String)
   Dim i    As Integer
   Dim bAux As Boolean
   Dim Txt As String
   
   With XDb.Tables(pTabela)
      '* ALTERARCHAVE
      bAux = False
      If Not IsEmpty(ArrId(0)) Then
         If UBound(pArrChave) <= 1 Then
            bAux = (pArrChave(LBound(pArrChave)) = pAutoIncrement)
         End If
      End If
      
      If Not bAux Then
         If bComDLL Then
            Print #1, "Public Function AlterarChave(" & pAlterChParam & ", Optional ComCOMMIT = False) As Integer"
         Else
            Print #1, "Public Function AlterarChave(" & pAlterChParam & ") As Integer"
         End If
         Print #1, "   Dim Sql As String"
         Print #1, " "
         If Not bComDLL Then
            Print #1, "   On Error GoTo Fim"
         End If
         Print #1, "   Sql = """ & "Update " & pTabela & " Set " & """"
         If Not IsEmpty(ArrId(0)) Then
            For i = LBound(pArrChave) To UBound(pArrChave) - 1
               If pAutoIncrement <> pArrChave(i) Then
                  Txt = "   Sql = Sql & """ & IIf(Trim(Txt) = "", Space(1), " , ") & .Fields(pArrChave(i)).NOME & " = """
                  Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
                     Case 1: Txt = Txt & " & SqlNum(Ch_" & .Fields(pArrChave(i)).NOME & ")"
                     Case 2: Txt = Txt & " & SqlDate(Ch_" & .Fields(pArrChave(i)).NOME & ", eSysDate.Data_Hora)"
                     Case 3: Txt = Txt & " & SqlStr(Ch_" & .Fields(pArrChave(i)).NOME & ")"
                  End Select
                  Print #1, Txt
               End If
            Next
         End If
         Print #1, "   Sql = Sql & """ & " Where "
         If Not IsEmpty(ArrId(0)) Then
            For i = LBound(pArrChave) To UBound(pArrChave) - 1
               Txt = "   Sql = Sql & """ & IIf(i = 0, Space(1), " and ") & pArrChave(i) & " = """
               Select Case GrpTipoCampo(.Fields(pArrChave(i)).Tipo)
                  Case 1: Txt = Txt & " & SqlNum(mvar" & pArrChave(i) & ")"
                  Case 2: Txt = Txt & " & SqlDate(mvar" & pArrChave(i) & ", eSysDate.Data_Hora)"
                  Case 3: Txt = Txt & " & SqlStr(mvar" & pArrChave(i) & ")"
               End Select
               Print #1, Txt
            Next
         End If
         If bComDLL Then
            Print #1, "   AlterarChave = mvar" & NmDbObj & ".Executa(Sql, ComCOMMIT)"
         Else
            Print #1, "   mvar" & NmDbObj & ".Execute Sql, dbExecDirect"
            Print #1, "   AlterarChave = True"
            Print #1, "Exit Function"
            Print #1, "Fim:"
            Print #1, "   AlterarChave = False"
            Print #1, "   msgInformacao " & """" & "Problema na atualização." & """" & " & vbNewLine & Errors(0).Description"
         End If
         Print #1, "End Function"
      End If
   End With
End Sub
Private Sub MontarClasse_Initialize()
   '* INITIALIZE
   Print #1, "Private Sub Class_Initialize()"
   Print #1, "   Set mvarRS = Nothing"
   Print #1, "   Set mvarXDb = Nothing"
   Print #1, "   mvarQryInsert = " & """"""
   Print #1, "   mvarQryUpDate = " & """"""
   Print #1, "   mvarQryDelete = " & """"""
   Print #1, "   mvarQrySelect = " & """"""
   Print #1, "   mvarQrySave = " & """"""
   If bisDirt Then Print #1, "   mvarisDirt = False"
   If bTipoQuery Then Print #1, "   mvarTipoQuery = ""I"""
   If bItensExcluidos Then Print #1, "mvarItensExcluidos = Nothing"
   Print #1, "End Sub"
End Sub
Private Sub MontarClasse_Terminate()
   '* TERMINATE
   Print #1, "Private Sub Class_Terminate()"
   Print #1, "   Set mvar" & NmDbObj & " = Nothing"
   Print #1, "   Set mvarRS = Nothing"
   If bItensExcluidos Then
      Print #1, "   Set mvarItensExcluidos = Nothing"
   End If
   Print #1, "End Sub"
End Sub
Private Sub DefineIdentity(pTabela As String, ByRef sAutoIncrement As String, ByRef sChaveMax As String)
   '*************
   '* Define Coluna 'Identity'
   Dim Sql As String
   
   sAutoIncrement = ""
   Sql = "SELECT Table_Name, Column_Name"
   Sql = Sql & " From information_schema.Columns"
   Sql = Sql & " Where Table_name='" & pTabela & "'"
   Sql = Sql & " and COLUMNPROPERTY( OBJECT_ID(table_name),Column_Name,'IsIdentity') = 1"
   If XDb.ExisteReg(Sql) Then
      sAutoIncrement = XDb.RSAux("Column_Name")
   End If
   sChaveMax = sAutoIncrement
End Sub
Private Sub DefineChaveMax(pTabela As String, ByRef sChaveMax As String)
   '*************
   '* Define Foreign Keys
   Dim Sql As String
   Dim RsForeign  As ADODB.Recordset
   
   Set RsForeign = XDb.ADOConect.OpenSchema(adSchemaForeignKeys, Array(Empty, Empty, pTabela))
   
   
   Sql = " SELECT DISTINCT C.COLUMN_NAME"
   Sql = Sql & " From information_schema.TABLE_CONSTRAINTS F"
   Sql = Sql & " , information_schema.KEY_COLUMN_USAGE C"
   Sql = Sql & " Where F.TABLE_NAME = C.TABLE_NAME"
   Sql = Sql & " And F.CONSTRAINT_NAME = C.CONSTRAINT_NAME"
   Sql = Sql & " And F.CONSTRAINT_TYPE = 'PRIMARY KEY'"
   Sql = Sql & " And F.TABLE_NAME='" & pTabela & "'"
   Sql = Sql & " And Substring(C.COLUMN_NAME,1,2) = 'ID'"
   Sql = Sql & " And C.COLUMN_NAME <> '" & sChaveMax & "'" ' sAutoIncrement & "'"
   Sql = Sql & " And C.COLUMN_NAME NOT IN (SELECT DISTINCT C2.COLUMN_NAME"
   Sql = Sql & " From information_schema.TABLE_CONSTRAINTS F2"
   Sql = Sql & " , information_schema.KEY_COLUMN_USAGE C2"
   Sql = Sql & " Where F2.TABLE_NAME = C2.TABLE_NAME"
   Sql = Sql & " And F2.CONSTRAINT_NAME = C2.CONSTRAINT_NAME"
   Sql = Sql & " And F2.CONSTRAINT_TYPE = 'FOREIGN KEY'"
   Sql = Sql & " And F2.TABLE_NAME='" & pTabela & "')"
   
   sChaveMax = ""
   If XDb.ExisteReg(Sql, RsForeign) Then
      If RsForeign.RecordCount = 1 Then
         sChaveMax = RsForeign("Column_Name")
      End If
   End If
End Sub
Private Sub DefineChaveEstrangeira(pTabela As String, ByRef pCHEstrangeira)
   '*************
   '* Define Foreign Keys
   Dim Sql As String
   Dim RsForeign  As ADODB.Recordset
   Dim sAux As String
   
   Set RsForeign = XDb.ADOConect.OpenSchema(adSchemaForeignKeys, Array(Empty, Empty, pTabela))
   
   
   Sql = " SELECT DISTINCT C.COLUMN_NAME"
   Sql = Sql & " From information_schema.TABLE_CONSTRAINTS F"
   Sql = Sql & " , information_schema.KEY_COLUMN_USAGE C"
   Sql = Sql & " Where F.TABLE_NAME = C.TABLE_NAME"
   Sql = Sql & " And F.CONSTRAINT_NAME = C.CONSTRAINT_NAME"
   Sql = Sql & " And F.CONSTRAINT_TYPE = 'FOREIGN KEY'"
   Sql = Sql & " And F.TABLE_NAME='" & pTabela & "'"
   Sql = Sql & " And C.COLUMN_NAME <> ''"
   Sql = Sql & " And C.COLUMN_NAME NOT IN (SELECT DISTINCT C2.COLUMN_NAME"
   Sql = Sql & " From information_schema.TABLE_CONSTRAINTS F2"
   Sql = Sql & " , information_schema.KEY_COLUMN_USAGE C2"
   Sql = Sql & " Where F2.TABLE_NAME = C2.TABLE_NAME"
   Sql = Sql & " And F2.CONSTRAINT_NAME = C2.CONSTRAINT_NAME"
   Sql = Sql & " And F2.CONSTRAINT_TYPE = 'PRIMARY KEY'"
   Sql = Sql & " And F2.TABLE_NAME='" & pTabela & "')"
   
   sAux = ""
   If XDb.ExisteReg(Sql, RsForeign) Then
      While Not RsForeign.EOF
         sAux = sAux & "|" & RsForeign("Column_Name")
         RsForeign.MoveNext
      Wend
      sAux = sAux & "|"
   End If
   pCHEstrangeira = sAux
End Sub
Private Sub DefineVarTamanho(pTabela As String, ByRef nVarTamanho As Integer, ByRef bTimeStamp As Boolean)
   '******
   '* Define Maior cadeia de Caracteres
   Dim i As Long
   
   bTimeStamp = False
   nVarTamanho = 0
   
   With XDb.Tables(pTabela)
      For i = 1 To .Fields.Count
         If Not bTimeStamp Then bTimeStamp = (.Fields(i).NOME = "TIMESTAMP")
         If Not .Fields(i).isSys Then
            nVarTamanho = IIf(nVarTamanho >= Len(.Fields(i).NOME), nVarTamanho, Len(.Fields(i).NOME))
         End If
      Next
   End With
End Sub
Private Sub DefineParametros(pTabela As String, pAutoIncrement As String, ByRef ArrChave, ByRef Chaves As String, ByRef ChaveParam As String, ByRef ChaveOptional As String, ByRef ChaveParamDef As String, ByRef AlterChParam As String)
   Dim i As Integer
   With XDb.Tables(pTabela)
      '* Define Chave e a seguencia de parametros opcionais
      On Error Resume Next
      Set ArrChave = Nothing
      If .PrimaryKey.Count > 0 Then
         ReDim ArrChave(.PrimaryKey.Count)
      End If
      For i = 0 To .PrimaryKey.Count - 1
         ArrChave(i) = .PrimaryKey(i + 1).NOME
      Next
      ChaveOptional = ""
      ChaveParamDef = ""
      ChaveParam = ""
      AlterChParam = ""
      Chaves = ""
      If Not IsEmpty(ArrId(0)) Then
         For i = LBound(ArrChave) To UBound(ArrChave) - 1
            ChaveOptional = ChaveOptional & IIf(i = LBound(ArrChave), Space(1), ", ") & "Optional Ch_" & ArrChave(i)
            ChaveParamDef = ChaveParamDef & IIf(i = 0, Space(1), ", ") & "Ch_" & ArrChave(i) & " As String"
            ChaveParam = ChaveParam & IIf(i = 0, Space(1), ", ") & "Ch_" & ArrChave(i)
            If pAutoIncrement <> ArrChave(i) Then
               AlterChParam = AlterChParam & IIf(Trim(AlterChParam) = "", Space(1), ", ") & "Ch_" & ArrChave(i) & " As String"
            End If
            Chaves = Chaves & IIf(i = 0, Space(1), ", ") & "mvar" & ArrChave(i)
         Next
      End If
   End With
End Sub
Private Sub DefineStrCampos(pTabela As String, pAutoIncrement As String, pVarTamanho As Integer, ByRef CamposSel As String, ByRef CamposIns As String)
   Dim i       As Integer
   Dim j       As Integer
   Dim nSpcVar As Integer
   Dim Txt     As String
   
   With XDb.Tables(pTabela)
      nSpcVar = pVarTamanho
      '* Define CamposSel
      j = 0
      For i = 1 To .Fields.Count
         If Not .Fields(i).isSys Then
            Select Case GrpTipoCampo(.Fields(i).Tipo)
               Case 1: Txt = " As Double" '* Numérico
               Case 2: Txt = " As String" '* Data
               Case 3: Txt = " As String" '* Caracter
               Case 4: Txt = " As Object" '* Object
            End Select
            Print #1, "Private mvar" & .Fields(i).NOME & Space(nSpcVar - Len(.Fields(i).NOME)) & Txt
            If GrpTipoCampo(.Fields(i).Tipo) > 0 And GrpTipoCampo(.Fields(i).Tipo) <= 3 Then
               If (i Mod 5) = 0 And i <> 0 Then
                  CamposSel = CamposSel & """" & " & VbNewLine " & vbNewLine
                  CamposSel = CamposSel & "   Sql = Sql & """ & IIf(j = 0, "", ", ") & .Fields(i).NOME
                  If pAutoIncrement <> .Fields(i).NOME Then
                     CamposIns = CamposIns & """" & " & VbNewLine " & vbNewLine
                     CamposIns = CamposIns & "   Sql = Sql & """ & IIf(j = 0, "", ", ") & .Fields(i).NOME
                  End If
               Else
                  CamposSel = CamposSel & IIf(j = 0, "", ", ") & .Fields(i).NOME
                  If pAutoIncrement <> .Fields(i).NOME Then
                     CamposIns = CamposIns & IIf(j = 0 Or Len(Trim(CamposIns)) = 0, "", ", ") & .Fields(i).NOME
                  End If
               End If
            End If
            j = 1
         End If
      Next
      If Right(CamposSel, Len(" & VbNewLine ")) <> " & VbNewLine " Then CamposSel = CamposSel & """" & " & VbNewLine "
   End With
End Sub
Public Sub F_REFRESH()
   '   Call Popula_
End Sub
Public Sub MontaLstTabForm()
   Dim i As Integer
   Dim Pos As Integer
   '* Montar Combo de Tabelas
   
   With XDb
      If .isADO Then
         Me.LstTabForm.Clear
         For i = 1 To .Tables.Count
            If Not .Tables(i).isSys Then
               Pos = InStr(.Tables(i).NOME, ".")
               If Pos > 0 Then
                  UserDB = Mid(.Tables(i).NOME, 1, Pos - 1)
                  Me.LstTabForm.AddItem Mid(.Tables(i).NOME, Pos + 1)
               Else
                  Me.LstTabForm.AddItem .Tables(i).NOME
               End If
            End If
         Next
         UserDB = ""
      Else
         With .dBase
            Me.LstTabForm.Clear
            For i = 0 To .TableDefs.Count - 1
               If (.TableDefs(i).Attributes And dbSystemObject) = 0 Then
                  Pos = InStr(XDb.dBase.TableDefs(i).Name, ".")
                  If Pos > 0 Then
                     UserDB = Mid(XDb.dBase.TableDefs(i).Name, 1, Pos - 1)
                     Me.LstTabForm.AddItem Mid(XDb.dBase.TableDefs(i).Name, Pos + 1)
                  Else
                     Me.LstTabForm.AddItem XDb.dBase.TableDefs(i).Name
                  End If
               End If
            Next
         End With
      End If
      If UserDB <> "" Then UserDB = UserDB & "."
   End With
   
   Call LstTabForm_Click
End Sub
Private Sub ChkCampo_Click()
   Dim Campo$
   
   If Tabela = "" Then Exit Sub
   If Campo = "" Then Exit Sub
   If Not Me.LstTabForm.Selected(Me.LstTabForm.ListIndex) Then Exit Sub
   Campo = GrdCampos.TextMatrix(GrdCampos.Row, 1)
   MyPrj.FORMS(Tabela).CONTROLES(Campo).Flag = (ChkCampo.Value = vbChecked)
End Sub
Private Sub CmbCtrl_Click()
   Dim Campo$
   If Tabela = "" Then Exit Sub
   If Not Me.LstTabForm.Selected(Me.LstTabForm.ListIndex) Then Exit Sub
   Campo = GrdCampos.TextMatrix(GrdCampos.Row, 1)
   On Error Resume Next
   MyPrj.FORMS(Tabela).CONTROLES(Campo).Tipo = Me.CmbCtrl.Text
End Sub
Private Sub CmbNotNull_Click()
   Dim Campo$
   On Error Resume Next
   If Tabela = "" Then Exit Sub
   If Not Me.LstTabForm.Selected(Me.LstTabForm.ListIndex) Then Exit Sub
   Campo = GrdCampos.TextMatrix(GrdCampos.Row, 1)
   If MyPrj.FORMS(Tabela).CONTROLES.Count > 0 Then
      MyPrj.FORMS(Tabela).CONTROLES(Campo).NotNull = (Me.CmbNotNull.Text = "Not Null")
   End If
   If Err <> 0 Then
      Call CmbNotNull_Click
      On Error GoTo 0
   End If
End Sub
Private Sub CmbOwner_Click()
   UserDB = Me.CmbOwner.Text
   Call MontarLstTabClss
End Sub
Private Sub CmdCarregaVetor_Click()
   Dim Arq$, CLASSE$
   Dim pTabela As String
   Dim Drv As String
   Dim TabName$, Campos$, Txt$
   Dim Chave, isKey%
   Dim i%, j%, File&
   On Error GoTo Fim
   
   Tabela = Me.LstTabCls
   If UserDB <> "" Then
      TabName$ = UCase(pTabela)
      Tabela = UserDB & "." & UCase(pTabela)
   Else
      TabName$ = UCase(pTabela)
      Tabela = UCase(pTabela)
   End If
   Drv$ = IIf(Me.TxtDrvDest = "", "C:\TMP\", Me.TxtDrvDest.Tag)
   CLASSE$ = TabName$
   Arq = CLASSE & ".txt"
   Call SetHourglass(hWnd)
   Call Del(Drv$ & Arq$)
   '   AbrirTxt% = FreeFile()
   Open Drv & Arq For Output As #1
   Print #1, "Function Carrega_Vetor(cTabela As String, Atributos() As String, Indices() As String) As Integer"
   Print #1, "'------------------------------------------------------------------------"
   Print #1, "' Funcao     : Carrega_Vetor"
   Print #1, "' Autor      : Diogenes"
   Print #1, "' Atualização:"
   Print #1, "' Data       : 06/12/1999"
   Print #1, "' Parametro  : cTabela - Nome da tabela a ser carregada"
   Print #1, "'              Atributos() - Vetor que sera preenchido com a estrutura da"
   Print #1, "'                            tabela"
   Print #1, "'              Indices()   - vetor que sera preenchido com os indices"
   Print #1, "' Retorno    : true/false - se carregada com sucesso"
   Print #1, "' Obj.       : preenche os vetores com a estrutura da tabela e seus indices"
   Print #1, "'------------------------------------------------------------------------"
   Print #1, ""
   Print #1, "    On Error GoTo Carga_Err"
   
   Print #1, "    Carrega_Vetor = False"
   Print #1, "    '" & pTabela
   With XDb.dBase.TableDefs(pTabela)
      Print #1, "    ReDim Atributos(" & CStr(.Fields.Count - 1) & ", 3)"
      j = 0
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Name <> "TKP_SEQ" Then
            Print #1, "    Atributos(" & CStr(j) & ", 1) = """ & .Fields(i).Name & """"
            Select Case GrpTipoCampo(.Fields(i).Type)
               Case 1: Print #1, "    Atributos(" & CStr(j) & ", 2) = dbDouble"
               Case 2: Print #1, "    Atributos(" & CStr(j) & ", 2) = dbDate"
               Case 3: Print #1, "    Atributos(" & CStr(j) & ", 2) = dbText"
            End Select
            Print #1, "    Atributos(" & CStr(j) & ", 3) = """ & .Fields(i).Size & """"
            j = j + 1
         End If
      Next
   End With
   Print #1, ""
   Print #1, "    ReDim Indices(0, 2)"
   Print #1, ""
   Print #1, "    Carrega_Vetor = True"
   Print #1, ""
   Print #1, "    GoTo Carga_Fim"
   Print #1, ""
   Print #1, "Carga_Err:"
   Print #1, "    SqlError"
   Print #1, "    Resume Carga_Fim"
   Print #1, ""
   Print #1, "Carga_Fim:"
   Print #1, ""
   Print #1, "End Function"
   Close #1
   Call SetDefault(hWnd)
Fim:
   ShowError
End Sub

Private Sub CmdChk_Click(Index As Integer)
   Dim i       As Integer
   Dim nAntes  As Integer
   Me.LstTabCls.Visible = False
   nAntes = Me.LstTabCls.ListIndex
   bSelecionado = True
   For i = Me.LstTabCls.ListCount - 1 To 0 Step -1
      Select Case Index
         Case 0: Me.LstTabCls.Selected(i) = True
         Case 1: Me.LstTabCls.Selected(i) = False
         Case 2: Me.LstTabCls.Selected(i) = Not Me.LstTabCls.Selected(i)
      End Select
   Next
   Me.LstTabCls.ListIndex = nAntes
   Me.LstTabCls.Visible = True
   Me.LstTabCls.SetFocus
   bSelecionado = False
End Sub

Private Sub CmdDrv_Click(Index As Integer)
   Dim PATH$
   Dim Tit$, Filtro$, Arq$, Ind%
   
   Tit$ = "Find Project"
   Filtro = "Project Files (*.vbp)|*.vbp"
   Ind% = 1
   Me.CmDialog.InitDir = "C:\SISTEMAS\"
   Arq$ = ProcurarArquivo(Me.CmDialog, Tit$, Arq$, Filtro$, Ind%)
   If Arq$ = "" Then
      Me.OptInPrj(1).Value = True
      Exit Sub
   End If
   Me.TxtDrvDest.Text = UCase(Me.CmDialog.Tag) & Arq
   Me.TxtDrvDest.Tag = UCase(Me.CmDialog.Tag)
   
   Tit$ = "Find DataBase Class"
   Filtro = "Project Files (*.cls)|*.cls"
   Ind% = 1
   Me.CmDialog.InitDir = Me.TxtDrvDest.Tag
   Arq$ = ProcurarArquivo(Me.CmDialog, Tit$, Arq$, Filtro$, Ind%)
   If Arq$ = "" Then
      Exit Sub
   End If
   Me.TxtSuperClasse.Text = UCase(Me.CmDialog.Tag) & Arq
   Me.TxtSuperClasse.Tag = Arq
End Sub
Private Sub CmdOperSair_Click()
   Unload Me
End Sub
Private Sub CmdOperCls_Click()
   Dim Sql As String
   Dim DscExclusao As String
   Dim i As Integer
   Dim j As Integer
   Call SetHourglass(hWnd)
   Call CarregaOPs
   If ValidaCampos Then
      For i = 0 To Me.LstTabCls.ListCount - 1
         If Me.LstTabCls.Selected(i) Then
            Me.LstTabCls.ListIndex = i
            If Me.LstTabCls.ItemData(i) > 0 Then
               DscExclusao = Me.LstCampoCls.List(Me.LstTabCls.ItemData(i))
            Else
               DscExclusao = ""
            End If
            'On Error Resume Next
            Call MontarClasse(Me.LstTabCls.List(i), DscExclusao)
            j = j + 1
            If j = Me.LstTabCls.SelCount Then Exit For
         End If
      Next
      Call ExibirAviso(LoadMsg(34), LoadMsg(1))
      Call Shell("explorer " & DrvLocal, vbMaximizedFocus)
   End If
   Call SetDefault(hWnd)
End Sub
Private Sub CmdOperForm_Click()
   Call MontaForm
End Sub


Private Sub cmdScript_Click()
   Dim mTabela As Object
   Dim TabName$
   Dim Arq           As String
   Dim Drv           As String
   Dim Sql As String, DscExclusao$, i%, j%
   Dim CLASSE$
   
   If Trim(Me.TxtDrvDest.Text) = "" Or Right(Me.TxtDrvDest.Text, 3) = "..." Then
      Me.TxtDrvDest.Text = DrvLocal & "..."
      Me.TxtDrvDest.Tag = DrvLocal
   End If
   
   Call SetHourglass(hWnd)
   Call CarregaOPs
   If ValidaCampos Then
      For i = 0 To Me.LstTabCls.ListCount - 1
         If Me.LstTabCls.Selected(i) Then
            Me.LstTabCls.ListIndex = i
            TabName$ = Me.LstTabCls.List(i)
            Set mTabela = CreateObject("BANCO.TB_" & TabName$)
            With mTabela
               Set .XDb = XDb
               If .PESQUISAR() Then
                  Do Until .Rs.EOF = True
                     .POPULA .Rs
                     Sql = Sql & .QryInsert & vbCrLf
                     .Rs.MoveNext
                  Loop
               End If
            End With
            
            Drv$ = Me.TxtDrvDest.Tag
            CLASSE$ = IIf(Mid(TabName$, 1, 3) = "TB_", TabName$, "TB_" & TabName$)
            Arq = CLASSE & ".sql"
            Call SetHourglass(hWnd)
            Call MakePath(Drv$)
            Call Del(Drv$ & Arq$)
            Close #1
            Open Drv & Arq For Output As #1
            Print #1, "-- Insert de inicialização da tabela " & TabName$
            Print #1, "-- Gerado em : " & Format(Now, "dd/mm/yyyy hh:mm")
            Print #1, "SET DATEFORMAT 'dmy' "
            Print #1, Sql
            Print #1, "-- "
            Close #1
            Sql = ""
            j = j + 1
            If j = Me.LstTabCls.SelCount Then Exit For
         End If
      Next
      
      
      Call ExibirAviso(LoadMsg(34), LoadMsg(1))
      Call Shell("explorer " & DrvLocal, vbMaximizedFocus)
   End If
   Call SetDefault(hWnd)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   DoEvents
   Me.Refresh
   If Not XDb.Conectado Then
      Unload Me
      Exit Sub
   End If
   Call Formata_TabClass
   Call ConfigForm(Me, Me.Icon)
   
   
   Me.GrdCampos.Height = Me.LstTabForm.Height + 100
   Me.Visible = True
   Me.Refresh
   Me.LstTabCls.SetFocus
   Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Initialize()
   Dim Arq As String
   Dim i As Integer

   mvarLocalReg = App.PATH & "\" & App.EXEName & ".reg"
   
   With XDb
      If Not .Conectado Then
         FrmOpBanco.Show vbModal
         If XDb.dbDrive = "" And XDb.Server = "" Then
            End
         End If
         If XDb.isADO Then
            XDb.SrvConecta
         ElseIf Not XDb.isODBC Then
            .isODBC = True
            frmODBCLog.Show vbModal
         Else
            If Dir("C:\DSR\", vbDirectory) <> "" Then
               Me.CmDialog.InitDir = "C:\DSR\"
            End If
            Arq$ = .Alias
            If .Alias = "" Then
               Arq$ = ProcurarArquivo(Me.CmDialog, "Abrir Banco de Dados Access", , "Microsoft Access MDBs (*.mdb)|*.mdb")
               .isODBC = False
               .dbDrive = Me.CmDialog.Tag
               .dbName = Arq$
            End If
            If Arq$ <> "" Then
               Call .SrvConecta(.dbDrive, .dbName, "", "", "", "")
            End If
         End If
         For i = 5 To 2 Step -1
            Arq$ = Trim(GetSetting(App.EXEName, "Outros", "BDRecente" & CStr(i - 1), ""))
            If XDb.Alias <> Arq$ And Arq$ <> "" Then
               Call SaveSetting(App.EXEName, "Outros", "BDRecente" & CStr(i), Arq$)
            End If
         Next
         If XDb.Alias <> "" Then
            Call SaveSetting(App.EXEName, "Outros", "BDRecente1", XDb.Alias)
         End If
         '         If Not .Conectado Then
         '            Unload Me
         '         End If
      End If
   End With
   If XDb.Conectado Then
      Me.Show
   Else
      MsgBox "Conexão Falhou!!"
      End
      'Call Form_Initialize
   End If
End Sub
Private Sub Form_Load()
   Me.Visible = False
   Me.TabPrj.TabVisible(1) = False
   Me.TabPrj.TabVisible(2) = False

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape:  Unload Me
      Case Else: KeyAscii = SendTab(Me, KeyAscii)
   End Select
End Sub
Private Sub Form_Resize()
   Call PintarFundo(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set MyPrj = Nothing
   Call SetDefault(hWnd)
End Sub
Private Sub GrdCampos_EnterCell()
   '  Call MyGrd.EnterCell
End Sub
Private Sub GrdCampos_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      Me.ChkCampo.Value = IIf(Me.ChkCampo.Value = vbChecked, vbUnchecked, vbChecked)
   End If
End Sub
Private Sub GrdCampos_RowColChange()
   Dim Campo As String
   
   If Not Me.GrdCampos.Visible Then Exit Sub
   
   With Me.GrdCampos
      LinAnt% = .Row
      If .Row = 0 Then Exit Sub
      '*****************************
      '* Preencher Linha de Edição *
      '*****************************
      DoEvents
      Call MyGrd.MoverLinha
      Campo = GrdCampos.TextMatrix(GrdCampos.Row, 1)
      On Error Resume Next
      '      ChkCampo.Value = IIf(GetTag(GrdCampos, "CEL(" & CStr(.Row) & ",0)") <> "False", vbChecked, vbUnchecked)
      ChkCampo.Value = MyPrj.FORMS(Tabela).CONTROLES(Campo).Flag
      ' IIf(.TextMatrix(.Row, 0) = "", vbUnchecked, vbChecked)
      TxtCampo = .TextMatrix(.Row, 1)
      If Me.LstTabForm.Selected(Me.LstTabForm.ListIndex) Then
         Call LocalizarCombo(CmbCtrl, .TextMatrix(.Row, 2))
      End If
      Call LocalizarCombo(CmbNotNull, .TextMatrix(.Row, 3))
      If TxtCampo <> "" Then
         'Me.CmbCtrl.Enabled = (GrpTipoCampo(xDb.dBase.TableDefs(UserDB & LstTabForm).Fields(TxtCampo).Type) <> 2)
         Me.CmbCtrl.Enabled = (GrpTipoCampo(XDb.dBase.TableDefs(Tabela).Fields(TxtCampo).Type) <> 2)
      End If
      '*****************************
      '*****************************
      Call MyGrd.MoverLinha
   End With
End Sub
Private Sub GrdCampos_LeaveCell()
   If GrdCampos.Row <> LinAnt Then
      GrdCampos.Col = 0
      GrdCampos.Row = GrdCampos.Row
      Exit Sub
   End If
'xxx   Call MyGrd.LeaveCell
   
   '   GrdCampos.c Matrix(2, 0) = IIf(n.Value = vbChecked, "`", "")
End Sub
Private Sub GrdCampos_Scroll()
   Call MyGrd.MoverLinha(True)
End Sub
Private Sub GrdCampos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '   Dim i%, AllChk As Boolean
   '   If Me.GrdCampos.MouseCol = 0 And Me.GrdCampos.MouseRow = 0 Then
   '      AllChk = True
   '      For i = 1 To Me.GrdCampos.Rows - 1
   '         If Me.GrdCampos.TextMatrix(i, 0) = "" Or Me.GrdCampos.TextMatrix(i, 0) = "0" Then
   '            AllChk = False
   '            Exit For
   '         End If
   '      Next
   '      For i = 1 To Me.GrdCampos.Rows - 1
   '         Me.GrdCampos.TextMatrix(i, 0) = IIf(AllChk, "", "1")
   '      Next
   '   End If
   '   Me.GrdCampos.Col = 0
End Sub
Private Sub Lbl_Click(Index As Integer)
   Select Case Index
      Case 7
         'Me.ChkSelectAll.Value = IIf(Me.ChkSelectAll.Value = vbChecked, vbUnchecked, vbChecked)
   End Select
End Sub
Private Sub LstCampoCls_Click()
   Dim i As Integer
   If Not Me.LstTabCls.Selected(Me.LstTabCls.ListIndex) Then Exit Sub
   Me.LstTabCls.ItemData(Me.LstTabCls.ListIndex) = Me.LstCampoCls.ListIndex
   If Me.LstCampoCls.SelCount > 1 Then
      For i = 0 To Me.LstCampoCls.ListCount - 1
         If Me.LstCampoCls.Selected(i) And Me.LstCampoCls.ListIndex <> i Then
            Me.LstCampoCls.Selected(i) = False
         End If
         If Me.LstCampoCls.SelCount = 1 Then Exit For
      Next
   ElseIf Me.LstCampoCls.SelCount = 0 Then Me.LstCampoCls.Selected(0) = False
   End If
End Sub
Private Sub LstCampoCls_ItemCheck(Item As Integer)
   If Not Me.LstTabCls.Selected(Me.LstTabCls.ListIndex) Then
      Call ExibirAviso("Selecione a Tabela.", LoadMsg(1))
      Me.LstCampoCls.Selected(Me.LstCampoCls.ListIndex) = False
   End If
End Sub

Private Sub LstTabCls_Click()
   Dim pTabela As String
   Dim i As Integer, j As Integer
   
   '* Montar Combo de Campo de Descrição
   If bSelecionado Then Exit Sub
   Me.LstCampoCls.Clear
   If Me.LstTabCls.Selected(Me.LstTabCls.ListIndex) Then
      Me.LstCampoCls.AddItem "  -- Em Branco -- "
   End If
   If UserDB = "" Or XDb.isADO Then
      pTabela = Me.LstTabCls
   Else
      pTabela = UserDB & "." & Me.LstTabCls
   End If
   
   
   '   On Error Resume Next
   Call DefineArrayID(pTabela)
   With XDb
      If .isADO Then
         With .Tables(pTabela)
            For i = 1 To .Fields.Count
               If Not .Fields(i).isSys Then
                  Me.LstCampoCls.AddItem .Fields(i).NOME
               End If
            Next
         End With
      Else
         With .dBase.TableDefs(pTabela)
            For i = 0 To .Fields.Count - 1
               If .Fields(i).Attributes < dbSystemField Then
                  Me.LstCampoCls.AddItem .Fields(i).Name
               End If
            Next
         End With
      End If
   End With
   If Me.LstTabCls.Selected(Me.LstTabCls.ListIndex) Then
      Me.LstCampoCls.Selected(Me.LstTabCls.ItemData(Me.LstTabCls.ListIndex)) = True
   End If
   
End Sub
Private Sub LstTabCls_ItemCheck(Item As Integer)
   Dim pTabela As String
   Dim i As Integer
   If bSelecionado Then Exit Sub
   
   Me.LstCampoCls.Clear
   Me.LstCampoCls.AddItem "  -- Em Branco -- "
   If UserDB = "" Or XDb.isADO Then
      pTabela = Me.LstTabCls
   Else
      pTabela = UserDB & "." & Me.LstTabCls
   End If
   On Error Resume Next
   Call DefineArrayID(pTabela)
   With XDb
      If .isADO Then
         With .Tables(pTabela)
            For i = 1 To .Fields.Count
               If Not .Fields(i).isSys Then
                  Me.LstCampoCls.AddItem .Fields(i).Name
               End If
            Next
         End With
      Else
         With .dBase.TableDefs(pTabela)
            For i = 0 To .Fields.Count - 1
               If .Fields(i).Attributes < dbSystemField Then
                  Me.LstCampoCls.AddItem .Fields(i).Name
               End If
            Next
         End With
      End If
   End With
End Sub
Private Sub LstTabForm_Click()
   Dim i%, j%, k%, Bool As Boolean
   Dim TxtAux$, Campo$
   Dim MyForm As New FORMULARIO
   Dim MyControl As New CONTROLE
   DoEvents
   TxtAux$ = IIf(UserDB = "", Me.LstTabForm, UserDB & Me.LstTabForm)
   
   Bool = Me.LstTabForm.Selected(Me.LstTabForm.ListIndex)
   Me.CmbCtrl.Enabled = True
   If Not Bool Then Me.CmbCtrl.ListIndex = -1
   Me.CmbCtrl.Enabled = Bool
   Me.ChkCampo.Enabled = Bool
   Me.CmbNotNull.Enabled = Bool
   
   '* Montar Classe de WizPrj
   If Bool Then
      With MyForm
         .FILENAME = "Frm" & IIf(InStr(TxtAux$, "TB_") = 0, TxtAux$, Mid(TxtAux, 4))
         .Flag = True
         .NOME = .FILENAME
      End With
      On Error Resume Next
      MyPrj.FORMS.Add MyForm, TxtAux$
      If Err = 457 Then '* This key is already associated with an element of this collection
         On Error GoTo 0
         Tabela = ""
      Else
         If Err <> 0 Then GoTo Fim
      End If
      If Tabela <> "" Then
         Call DefineArrayID(TxtAux$)
         k = 0
         For i = 1 To XDb.Tables(TxtAux$).Fields.Count
            If Not XDb.Tables(TxtAux$).Fields(i).isSys Then
               k = k + 1
               Campo$ = XDb.Tables(TxtAux$).Fields(i).NOME
               With MyControl
                  .Flag = True
                  .NOME = "Txt" & Campo$
                  If GrpTipoCampo(XDb.Tables(TxtAux$).Fields(i).Tipo) = 2 Then
                     .Tipo = "MaskEdit"
                  Else
                     .Tipo = "TextBox"
                  End If
                  Me.GrdCampos.TextMatrix(k, 1) = Campo$
                  Me.GrdCampos.TextMatrix(k, 2) = .Tipo
                  For j = LBound(ArrId) To UBound(ArrId)
                     If Campo = ArrId(j) Then
                        .NotNull = True
                        If Me.GrdCampos.Row = k Then
                           Call LocalizarCombo(Me.CmbNotNull, "Not Null")
                        Else
                           Me.GrdCampos.TextMatrix(k, 3) = "Not Null"
                        End If
                        
                        Exit For
                     End If
                  Next
               End With
               MyPrj.FORMS(TxtAux$).CONTROLES.Add MyControl, Campo$
               Set MyControl = Nothing
            End If
         Next
      End If
      Me.CmbCtrl.ListIndex = 0
   Else
      On Error Resume Next
      
      MyPrj.FORMS.Remove TxtAux$
      For i = 1 To Me.GrdCampos.Rows - 1
         If Me.GrdCampos.Row = i Then
            Call LocalizarCombo(Me.CmbNotNull, "")
         Else
            Me.GrdCampos.TextMatrix(i, 3) = ""
            Me.GrdCampos.TextMatrix(i, 2) = ""
         End If
      Next
      On Error GoTo 0
      On Error GoTo Fim
   End If
   Set MyForm = Nothing
   Set MyControl = Nothing
   
   
   If TxtAux$ = Tabela Then Exit Sub
   Tabela = TxtAux$
   TabName = Mid(TxtAux, Len(UserDB) + 1)
   
   '* Montar Grid de Campos
   Me.GrdCampos.Rows = 1
   
   Me.GrdCampos.TextMatrix(0, 1) = "Campo"
   Me.GrdCampos.TextMatrix(0, 2) = "Controle"
   Me.GrdCampos.TextMatrix(0, 3) = "Requerido"
   Me.GrdCampos.Visible = False
   With XDb.Tables(Tabela)
      '* Define Campos
      For i = 1 To .Fields.Count
         If Not .Fields(i).isSys Then
            Me.GrdCampos.Rows = Me.GrdCampos.Rows + 1
            Me.GrdCampos.Row = Me.GrdCampos.Rows - 1
            Me.GrdCampos.CellAlignment = flexAlignCenterCenter
            Campo$ = .Fields(i).NOME
            If Me.LstTabForm.Selected(Me.LstTabForm.ListIndex) Then
               Set MyControl = MyPrj.FORMS(Tabela).CONTROLES(Campo)
            Else
               With MyControl
                  .Flag = True
                  .NOME = "Txt" & Campo
                  .Tipo = ""
               End With
            End If
            With MyControl
               If .Flag Then
                  Set Me.GrdCampos.CellPicture = LoadResPicture("CHECKED", vbResBitmap)
                  Call SetTag(Me.GrdCampos, "CEL(" & CStr(Me.GrdCampos.Rows - 1) & ",0)", "True")
                  'Me.GrdCampos.TextMatrix(Me.GrdCampos.Rows - 1, 0) = "`"
               Else
                  Set Me.GrdCampos.CellPicture = LoadResPicture("UNCHECKED", vbResBitmap)
                  Call SetTag(Me.GrdCampos, "CEL(" & CStr(Me.GrdCampos.Rows - 1) & ",0)", "False")
                  'Me.GrdCampos.TextMatrix(Me.GrdCampos.Rows - 1, 0) = ""
               End If
               Me.GrdCampos.TextMatrix(Me.GrdCampos.Rows - 1, 1) = Campo$
               Me.GrdCampos.TextMatrix(Me.GrdCampos.Rows - 1, 2) = .Tipo
               Me.GrdCampos.TextMatrix(Me.GrdCampos.Rows - 1, 3) = IIf(.NotNull, "Not Null", "")
            End With
         End If
      Next
   End With
   
   Me.GrdCampos.Visible = True
   Me.GrdCampos.Row = 0
   Me.GrdCampos.Row = 1
   Exit Sub
Fim:
   Call ShowError
End Sub
Private Sub OptInPrj_Click(Index As Integer)
   Dim Bool As Boolean
   Me.Visible = False
   Me.Refresh
   Bool = (Index = 0)
   Me.TxtDrvDest.BackColor = IIf(Bool, &HC0FFFF, &HE0E0E0)
   Me.TxtSuperClasse.BackColor = IIf(Bool, &HC0FFFF, &HE0E0E0)
   Me.TxtDrvDest.Enabled = Bool
   Me.TxtSuperClasse.Enabled = Bool
   Me.CmdDrv(0).Enabled = Bool
   '   If Not Bool Then
   '      Me.Move Me.Left, Me.Top, 7300
   '      Me.TxtDrvDest.Text = ""
   '      Me.TxtDrvDest.Tag = ""
   '      Me.TxtSuperClasse.Text = ""
   '      Me.TxtSuperClasse.Tag = ""
   '   Else
   '      Me.Move Me.Left, Me.Top, 9270
   '   End If
   '   Call CentrarForm(me, Me)
   Me.Visible = True
   Me.Refresh
End Sub
Private Sub TabPrj_Click(PreviousTab As Integer)
   Select Case Me.TabPrj.Tab
      Case 1
         Call Formata_TabForm
         Call GrdCampos_RowColChange
         DoEvents
         Me.GrdCampos.Visible = True
   End Select
End Sub
Private Sub TxtCampo_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Public Sub MontaForm()
   Dim Drv$, Arq$, ArqName$
   Dim MyForm As New FORMULARIO
   Dim n, f&, i%, j%
   Call SetHourglass(hWnd)
   Drv$ = "C:\TMP\"
   With MyPrj
      For Each n In .FORMS
         Set MyForm = n
         Tabela = Mid(MyForm.FILENAME, 4)
         
         '* Define Chave
         Call DefineArrayID
         
         Arq = MyForm.FILENAME
         ArqName = Arq & ".FRM"
         Call Del(Drv$ & ArqName)
         f = FreeFile
         Open Drv & ArqName For Output As #f
         Print #f, "Version 5.00"
         Print #f, "Object = """ & "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0" & """" & "; " & """" & "MSMASK32.OCX" & """"
         Print #f, "Object = """ & "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0" & """" & "; " & """" & "COMCTL32.OCX" & """"
         '         Print #f, "Object = """ & "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0" & """" & "; " & """" & "TABCTL32.OCX" & """"
         '         Print #f, "Object = """ & "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0" & """" & "; " &  """" & "THREED32.OCX" & """"
         Print #f, "Begin VB.Form " & Arq
         Print #f, "   AutoRedraw = -1        'True"
         Print #f, "   BorderStyle = 3        'Fixed Dialog"
         Print #f, "   Caption = """ & "Cadastro " & Tabela$ & """"
         Print #f, "   ClientHeight = 6180"
         Print #f, "   ClientLeft = 2085"
         Print #f, "   ClientTop = 600"
         Print #f, "   ClientWidth = 11325"
         Print #f, "   ClipControls = 0       'False"
Print #f, "   BeginProperty Font"
   Print #f, "      Name           =  " & """" & "Times New Roman" & """"
   Print #f, "      Size           = 9.75"
   Print #f, "      Charset        = 0"
   Print #f, "      Weight         = 400"
   Print #f, "      Underline      = 0       'False"
   Print #f, "      Italic         = 0       'False"
   Print #f, "      Strikethrough  = 0       'False"
   Print #f, "   EndProperty"
   Print #f, "   KeyPreview = -1        'True"
   Print #f, "   LinkTopic = " & """" & "Form1" & """"
   Print #f, "   LockControls = -1      'True"
   Print #f, "   MaxButton = 0          'False"
   Print #f, "   MDIChild = -1          'True"
   Print #f, "   MinButton = 0          'False"
   Print #f, "   PaletteMode = 1        'UseZOrder"
   Print #f, "   ScaleHeight = 6180"
   Print #f, "   ScaleWidth = 11325"
   Print #f, "   WhatsThisButton = -1   'True"
   Print #f, "   WhatsThisHelp = -1     'True"
   '**********
   '* Montar CmdOper
   '**********
   Call MontaCmdOper(f)
   '**********
   '* Montar ImgFundo
   '**********
   '   Call MontaImgFundo(f)
   '**********
   '* Montar Objetos de Edição
   '**********
   Call MontaCampos(f, Tabela$)
   
   Print #f, "End"
   Print #f, "Attribute VB_Name = " & """" & Arq & """"
   Print #f, "Attribute VB_GlobalNameSpace = False"
   Print #f, "Attribute VB_Creatable = False"
   Print #f, "Attribute VB_PredeclaredId = True"
   Print #f, "Attribute VB_Exposed = False"
   '**************************
   '* Declaração de variáveis
   '**************************
   Print #f, "Option Explicit"
   'Print #f, "Public Suja as Boolean, PrimeiraVez As Boolean"
   'Print #f, "Public Acesso$, Oper%"
   'Print #f, "Public Grd As MSFlexGrid"
   '**************************
   '* ValidaCampos
   '**************************
   Call MontaValidaCampos(f, Tabela$)
   '**************************
   '* Popula_Tela
   '**************************
   Call MontaPopula_Tela(f, Tabela)
   '**************************
   '* F_INCLUIR
   '**************************
   '   Call MontaF_INCLUIR(f, Tabela)
   '**************************
   '* F_SALVAR
   '**************************
   Call MontaF_SALVAR(f, Tabela)
   '**************************
   '* F_EXCLUIR
   '**************************
   '   Call MontaF_EXCLUIR(f, Tabela)
   '**************************
   '* F_REFRESH
   '**************************
   '   Call MontaF_REFRESH(f, Tabela)
   '**************************
   '* F_PROCURAR
   '**************************
   Call MontaF_PROCURAR(f, Tabela)
   '**************************
   '* CmdOper_Click
   '**************************
   Call MontaCmdOper_Click(f, Tabela)
   '**************************
   '* Eventos do Form
   '**************************
   Call MontaEvtForm(f, Tabela)
   '**************************
   '* Eventos primários dos controle
   '**************************
   Call MontaEvtControl(f, Tabela)
   
   Close #f ' Close file.
Next
End With
Call ExibirAviso("Operação concluído com sucesso.", "Project Wizard")
Call SetDefault(hWnd)
End Sub
Public Sub MontaEvtControl(f, Table)
   Dim cName As String, n
'Print #f, "Private Sub LblId_Click(Index As Integer)"
Print #f, "Private Sub LblId_Click()"
   Print #f, "   Call F_PROCURAR()"
   Print #f, "End Sub"
   If Not IsEmpty(ArrId(0)) Then
      cName = MyPrj.FORMS(Table).CONTROLES(ArrId(LBound(ArrId))).NOME
   End If
   For Each n In MyPrj.FORMS(Table).CONTROLES
Print #f, "Private Sub " & n.NOME & "_GotFocus()"
   Print #f, "   Call SelecionarTexto(Me.ActiveControl)"
   Print #f, "End Sub"
   If cName = n.NOME Then
Print #f, "Private Sub " & cName & "_LostFocus()"
   Print #f, "   Call Popula_Tela"
   '         Print #f, "   If Trim(" & cName & ") = """" And Me.ActiveControl <> Me.CmdOper(3) Then"
   '         Print #f, "      Call LimparTela(Me)"
   '         Print #f, "   End If"
   Print #f, "End Sub"
End If
Next

End Sub
Public Sub MontaEvtForm(f&, Table$)
   Dim cName$
   If Not IsEmpty(ArrId(0)) Then
      cName = MyPrj.FORMS(Table).CONTROLES(ArrId(LBound(ArrId))).NOME
   End If
Print #f, "Private Sub Form_Activate()"
   Print #f, "   Screen.MousePointer = vbDefault"
   '   Print #f, "   Call SetHourglass(hWnd)"
   '   Print #f, "   Set MDIFilho = Me"
   '   Print #f, "   Call Popula_Tela"
   '   Print #f, "   If PrimeiraVez Then"
   '   Print #f, "      Me." & cName & ".SetFocus"
   '   Print #f, "      PrimeiraVez = False"
   '   Print #f, "   End If"
   '   Print #f, "   Call SetDefault(hWnd)"
   '   Print #f, "   If Not VerificaAcesso(Me.Acesso, LEITURA) Then"
   '   Print #f, "      Unload Me"
   '   Print #f, "   End If"
   Print #f, "End Sub"
   
Print #f, "Private Sub Form_Load()"
   '   Print #f, "   Call SetHourglass(hWnd)"
   '   Print #f, "   If Me.Oper = ALTERACAO Then"
   '   Print #f, "      Me." & cName & " = Grd.TextMatrix(Grd.Row, 0) 'IDSUPP"
   '   Print #f, "      Call Popula_Tela"
   '   Print #f, "   End If"
   Print #f, "   Call ConfigForm(Me, MDI.Icon, Sys.FundoTela)"
   '   Print #f, "   Call SetDefault(hWnd)"
   Print #f, "End Sub"
   
Print #f, "Private Sub Form_KeyPress(KeyAscii As Integer)"
   Print #f, "   Select Case KeyAscii"
   Print #f, "      Case vbKeyEscape: If Sys.SaiComESC Then Unload Me"
   Print #f, "      Case Else: KeyAscii = SendTab(Me, KeyAscii)"
   Print #f, "   End Select"
   Print #f, "End Sub"
   
Print #f, "Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)"
   Print #f, "   Select Case KeyCode"
   Print #f, "     '* Executar Lista de Valores ao teclar [F2]"
   Print #f, "      Case vbKeyF2: Call LblId_Click()"
   '   Print #f, "      Case vbKeyF2"
   '   Print #f, "         '* Executar Lista de Valores ao teclar [F2]"
   '   Print #f, "         Select Case Me.ActiveControl.Name"
   '   Print #f, "            Case Me." & cName & ".Name: Call LblId_Click(0)"
   '   Print #f, "         End Select"
   Print #f, "   End Select"
   Print #f, "End Sub"
   
'   Print #f, "Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)"
   '   Print #f, "   '============="
   '   Print #f, "   '=  Se nenhum campo foi alterado -> SAIR"
   '   Print #f, "   '============="
   '   Print #f, "   If Not Me.Suja Then Exit Sub"
   '   Print #f, "   '============="
   '   Print #f, "   '=   Se não deseja salvar -> SAIR"
   '   Print #f, "   '============="
   '   Print #f, "   If ExibirPergunta(LoadMsg(54), Me.Caption) = vbNo Then"
   '   Print #f, "      Exit Sub"
   '   Print #f, "   End If"
   '   Print #f, "   '============="
   '   Print #f, "   '=   Verificar e validar campos"
   '   Print #f, "   '============="
   '   Print #f, "   If ValidaCampos Then F_SALVAR"
   '   Print #f, "End Sub"
   
   Print #f, "Private Sub Form_Resize()"
   Print #f, "   Call PintarFundo(Me, Sys.FundoTela)"
   Print #f, "End Sub"
   
   Print #f, "Private Sub Form_Unload(Cancel As Integer)"
   '   Print #f, "   Set MDIFilho = Nothing"
   Print #f, "   Set BANCO.TB_" & Table & " = Nothing"
   Print #f, "   Call SetDefault(hWnd)"
   Print #f, "End Sub"
End Sub
Public Sub MontaF_INCLUIR(f&, pTabela As String)
   Dim cName$
   Print #f, "Public Sub F_INCLUIR()"
   Print #f, "   If Not VerificaAcesso(Me.Acesso, INCLUSAO) Then Exit Sub"
   Print #f, "   Call F_SALVAR"
   Print #f, "   Call LimparTela(Me)"
   cName = MyPrj.FORMS(pTabela).CONTROLES(ArrId(LBound(ArrId))).NOME
   Print #f, "   Me." & cName & ".SetFocus"
   Print #f, "End Sub"
End Sub
Public Sub MontaF_SALVAR(f&, pTabela As String)
   Dim i%, Txt$, cName$, cNameID$
   
   Print #f, "Public Sub F_SALVAR(Optional pInc_Alt_Exc = 0) "
   Print #f, " "
   '   Print #f, "   If Not VerificaAcesso(Me.Acesso, ALTERACAO) Then Exit Function"
   Print #f, "   If Not ValidaCampos() Then"
   Print #f, "      Exit Sub"
   Print #f, "   End If"
   Print #f, " "
   Print #f, "   With BANCO.TB_" & pTabela
   Txt = ""
   If Not IsEmpty(ArrId(0)) Then
      For i = LBound(ArrId) To UBound(ArrId)
         cName = MyPrj.FORMS(pTabela).CONTROLES(ArrId(i)).NOME
         Txt = Txt & IIf(i = LBound(ArrId), "", ",") & cName
      Next
   End If
   '   Print #f, "      Call .GetSelect(" & Txt$ & ")"
   With XDb.dBase.TableDefs(pTabela)
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            If Mid(.Fields(i).Name, 1, 2) = "ID" And cNameID = "" Then
               cNameID = MyPrj.FORMS(pTabela).CONTROLES(.Fields(i).Name).NOME
            End If
            cName = MyPrj.FORMS(pTabela).CONTROLES(.Fields(i).Name).NOME
            Print #f, "      ." & .Fields(i).Name & " = Me." & cName
         End If
      Next
   End With
   Print #f, " "
   Print #f, "      Select Case pInc_Alt_Exc"
   Print #f, "         Case 0: Call .Incluir(True)"
   Print #f, "         Case 1: Call .Alterar(True)"
   Print #f, "         Case 2: Call .Excluir(True)"
   Print #f, "      End Select"
   Print #f, "   End With"
   Print #f, "   Me." & cNameID & ".SetFocus"
   Print #f, "End Sub"
End Sub
Public Sub MontaF_EXCLUIR(f&, Table$)
   Dim i%, Txt$, cName$
   
Print #f, "Public Function F_EXCLUIR() As Boolean"
   Print #f, "   Dim Arr(0)"
   Print #f, "   If Not VerificaAcesso(Me.Acesso, EXCLUSAO) Then Exit Function"
   
   Txt = ""
   
   For i = LBound(ArrId) To UBound(ArrId)
      cName = MyPrj.FORMS(Table).CONTROLES(ArrId(i)).NOME
      Txt = Txt & IIf(i = LBound(ArrId), "", ",") & cName
   Next
   Print #f, "   Arr(0) = BANCO." & Table$; ".QryDelete(" & Txt & ")"
   
   Print #f, "   If xDb.Executa(Arr) Then"
   Print #f, "      Call LimparTela(Me)"
   Print #f, "      DoEvents"
   cName = MyPrj.FORMS(Table).CONTROLES(ArrId(LBound(ArrId))).NOME
   Print #f, "      Me." & cName & ".SetFocus"
   Print #f, "   End If"
   Print #f, "   F_EXCLUIR = True"
   Print #f, "End Function"
End Sub
Public Sub MontaF_REFRESH(f&, Table$)
Print #f, "Public Sub F_REFRESH()"
   Print #f, "'   Call Formata_Tela"
   Print #f, "   Call Popula_Tela"
   Print #f, "End Sub"
End Sub
Public Sub MontaF_PROCURAR(f&, Table$)
   Dim i%, k%, cName$
   Dim NumCampos As Integer
Print #f, " Public Sub F_PROCURAR(Optional Index = 0)"
   Print #f, "   Dim Arrid"
   Print #f, "   Select Case Index"
   
   With XDb.dBase.TableDefs(Tabela)
      NumCampos = 0
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField And Mid(.Fields(i).Name, 1, 2) = "ID" Then
            cName = MyPrj.FORMS(Table).CONTROLES(.Fields(i).Name).NOME
            NumCampos = NumCampos + 1
         End If
      Next
      If NumCampos <= 1 Then
         Print #f, "   Arrid = F_LOV(""" & Table & """)"
      Else
         k% = 0
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField And Mid(.Fields(i).Name, 1, 2) = "ID" Then
               Print #f, "      Case " & CStr(k%) & ": Arrid = F_LOV(""" & IIf(k = 0, Table, Mid(.Fields(i).Name, 3)) & """)"
               k = k + 1
            End If
         Next
         Print #f, "   End Select"
      End If
      Print #f, "   '======================="
      Print #f, "   If not IsEmpty(ArrId(0)) Then Exit Sub"
      Print #f, "   If UBound(Arrid) < 0 Then Exit Sub"
      Print #f, "   '======================="
      If NumCampos <= 1 Then
         Print #f, "   Me." & cName & ".SetFocus"
         Print #f, "   DoEvents"
         Print #f, "   Me." & cName & ".Text = Arrid(0)"
         Print #f, "   Me." & MyPrj.FORMS(Table).CONTROLES(NumCampos).NOME & ".SetFocus"
      Else
         Print #f, "   Select Case Index"
         k% = 0
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField And Mid(.Fields(i).Name, 1, 2) = "ID" Then
               cName = MyPrj.FORMS(Table).CONTROLES(.Fields(i).Name).NOME
               Print #f, "      Case " & CStr(k)
               Print #f, "         Me." & cName & ".SetFocus"
               Print #f, "         DoEvents"
               Print #f, "         Me." & cName & ".Text = Arrid(0)"
               '            If k = 0 Then
               '               Print #f, "'         Me.Msk = Arrid(1)"
               '               Print #f, "         Call Popula_Tela"
               '            Else
               Print #f, "         Me." & MyPrj.FORMS(Table).CONTROLES(NumCampos).NOME & ".SetFocus"
               '            End If
               k = k + 1
            End If
         Next
         Print #f, "   End Select"
      End If
      Print #f, "End Sub"
   End With
End Sub
Public Sub MontaValidaCampos(f, Table$)
   Dim i%, j%, k%
   Dim ContId%
Print #f, "Public Function ValidaCampos() As Boolean"
   With XDb.dBase.TableDefs(Tabela)
      k% = 0
      ContId = 0
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            With MyPrj.FORMS(Table).CONTROLES(.Fields(i).Name)
               If .NotNull Then
                  Print #f, "   If Trim(Me." & .NOME & ") = " & """""" & " Then"
                  If Mid(XDb.dBase.TableDefs(Tabela).Fields(i).Name, 1, 2) = "ID" Then
                     Print #f, "      Call ExibirAviso(LoadMsg(27) & vbCrLf & Me.LblId(" & CStr(ContId) & "), LoadMsg(1))"
                     ContId = ContId + 1
                  Else
                     Print #f, "      Call ExibirAviso(LoadMsg(27) & vbCrLf & Me.Lbl(" & CStr(k - ContId) & "), LoadMsg(1))"
                  End If
                  Print #f, "      Me." & .NOME & ".SetFocus"
                  Print #f, "      Exit Function"
                  Print #f, "   End If"
               End If
            End With
            k = k + 1
         End If
      Next
   End With
   Print #f, "   ValidaCampos = True"
   Print #f, "End Function"
End Sub
Public Sub MontaPopula_Tela(f&, Table$)
   Dim i%, Txt$
   Dim cName$
   
Print #f, "Public Sub Popula_Tela()"
   Txt = "   Dim "
   If Not IsEmpty(ArrId(0)) Then
      For i = LBound(ArrId) To UBound(ArrId)
         Txt = Txt & IIf(i = LBound(ArrId), "", ",") & "l" & ArrId(i) & "$"
      Next
   End If
   Print #f, Txt
   Print #f, " "
   Print #f, "   If Me.ActiveControl Is Me.CmdOper(3) Then"
   Print #f, "      Exit Sub"
   Print #f, "   End If"
   Print #f, " "
   If Not IsEmpty(ArrId(0)) Then
      For i = LBound(ArrId) To UBound(ArrId)
         cName = MyPrj.FORMS(Table).CONTROLES(ArrId(i)).NOME
         Print #f, "   If Trim(Me." & cName & ") = """" Then"
         Print #f, "      Me." & cName & ".SetFocus"
         Print #f, "      Exit Sub"
         Print #f, "   End If"
      Next
   End If
   
   Print #f, "   With BANCO.TB_" & Table$
   
   Txt = ""
   If Not IsEmpty(ArrId(0)) Then
      For i = LBound(ArrId) To UBound(ArrId)
         cName = MyPrj.FORMS(Table).CONTROLES(ArrId(i)).NOME
         Txt = Txt & IIf(i = LBound(ArrId), "", ",") & cName
      Next
   End If
   
   Print #f, "      If .Pesquisar(" & Txt & ") Then"
   Print #f, "         '* Popula Tela"
   With XDb.dBase.TableDefs(Tabela)
      For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            cName$ = MyPrj.FORMS(Table).CONTROLES(.Fields(i).Name).NOME
            Select Case GrpTipoCampo(.Fields(i).Type)
               Case 1, 3
                  Select Case MyPrj.FORMS(Table).CONTROLES(.Fields(i).Name).Tipo
                     Case "ComboBox"
                        Print #f, "         Call LocalizarCombo(Me." & cName & ", ." & .Fields(i).Name & ")"
                     Case "CheckBox"
                        Print #f, "         Me." & cName & ".Value = iif(." & .Fields(i).Name & "=0,vbUnchecked,vbChecked)"
                     Case "OptionButton"
                        Print #f, "         Me." & cName & "(." & .Fields(i).Name & ").value = True"
                     Case Else
                        Print #f, "         Me." & cName & " = ." & .Fields(i).Name
                  End Select
               Case 2:    Print #f, "         Me." & cName & " = DToMask(." & .Fields(i).Name & ", Me." & cName & ")"
            End Select
         End If
      Next
      Print #f, "  "
      Print #f, "         Me.CmdOper(1).Enabled = True"
      Print #f, "         Me.CmdOper(2).Enabled = True"
      Print #f, "      Else"
   End With
   If Not IsEmpty(ArrId(0)) Then
      For i = LBound(ArrId) To UBound(ArrId)
         cName = MyPrj.FORMS(Table).CONTROLES(ArrId(i)).NOME
         Print #f, "         l" & ArrId(i) & "$ = Me." & cName
      Next
   End If
   Print #f, "         Call LimparTela(Me)"
   If Not IsEmpty(ArrId(0)) Then
      For i = LBound(ArrId) To UBound(ArrId)
         cName = MyPrj.FORMS(Table).CONTROLES(ArrId(i)).NOME
         Print #f, "         Me." & cName & " = l" & ArrId(i) & "$"
      Next
   End If
   Print #f, "         Me.CmdOper(0).Enabled = True"
   Print #f, "      End If"
   Print #f, "   End With"
   Print #f, "End Sub"
End Sub
Public Sub MontaCmdOper_Click(f&, Table$)
Print #f, "Private Sub CmdOper_Click(Index As Integer)"
   Print #f, "   Select Case Index"
   Print #f, "      Case 3: Unload Me"
   Print #f, "      Case Else: Call F_SALVAR(Index)"
   Print #f, "   End Select"
   Print #f, "End Sub"
   
'   Print #f, "Private Sub CmdOper_Click(Index As Integer)"
   '   Print #f, "   Select Case Index"
   '   Print #f, "      Case 0: Call F_INCLUIR"
   '   Print #f, "      Case 1: Call F_SALVAR"
   '   Print #f, "      Case 2: Call F_EXCLUIR"
   '   Print #f, "      Case 3: Unload Me"
   '   Print #f, "   End Select"
   '   Print #f, "End Sub"
End Sub
Public Sub MontaCampos(f, Table$)
   Const MaxW = 11325
   Const Sep = 280
   
   Dim i%, j%, k%
   Dim TabInd%, TopCtrl%
   Dim ll&, PosX&
   
   With XDb.Tables(Table$)
      j = 0
      k = 0
      TabInd% = 1
      TopCtrl = 120
      ll& = 120
      PosX& = 0
      For i = 0 To .Fields.Count - 1
         If PosX& <> 0 Then
            '         PosX& = PosX& + (1680 + ll&) + (160 * .Fields(i).Size)
            If PosX& + (1680 + ll&) + (160 * .Fields(i).Size) > MaxW Then
               TopCtrl = TopCtrl + 320
               PosX& = 0
            End If
         End If
         
         If .Fields(i).Attributes < dbSystemField Then
            Select Case MyPrj.FORMS(Table).CONTROLES(.Fields(i).Name).Tipo
                  '**********
                  '* TextBox
                  '**********
               Case "TextBox"
                  Print #f, "   Begin VB.TextBox Txt" & .Fields(i).Name
Print #f, "      BeginProperty Font"
   Print #f, "         Name           = " & """" & "Times New Roman" & """"
   Print #f, "         Size           = 9.75"
   Print #f, "         Charset        = 0"
   Print #f, "         Weight         = 400"
   Print #f, "         Underline      = 0       'False"
   Print #f, "         Italic         = 0       'False"
   Print #f, "         Strikethrough  = 0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Height = 330"
   Print #f, "      Left = " & CStr(1680 + ll& + PosX&)
   Print #f, "      MaxLength = " & .Fields(i).Size
   Print #f, "      TabIndex = " & CStr(TabInd%)
   'Print #f, "      Tag = " & """" & """"
   Print #f, "      Top = " & CStr(TopCtrl)
   'Print #f, "      WhatsThisHelpID = 10249"
   Print #f, "      Width = " & CStr(160 * .Fields(i).Size)
   Print #f, "   End         "
   '**********
   '* MaskEdit
   '**********
Case "MaskEdit"
   Print #f, "   Begin MSMask.MaskEdBox Msk" & .Fields(i).Name
Print #f, "      BeginProperty Font"
   Print #f, "         Name = " & """" & "Times New Roman" & """"
   Print #f, "         Size           = 9.75"
   Print #f, "         Charset        = 0"
   Print #f, "         Weight         = 400"
   Print #f, "         Underline      = 0       'False"
   Print #f, "         Italic         = 0       'False"
   Print #f, "         Strikethrough  = 0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Height = 330"
   Print #f, "      Left = " & CStr(1680 + ll& + PosX&)
   Print #f, "      MaxLength = " & CStr(.Fields(i).Size)
   Print #f, "      TabIndex = " & CStr(TabInd%)
   Print #f, "      Top = " & CStr(TopCtrl)
   '                  Print #f, "      WhatsThisHelpID = 10247"
   Print #f, "      Width = " & CStr(160 * .Fields(i).Size)
   Print #f, "      _ExtentX       =   1296"
   Print #f, "      _ExtentY       =   582"
   Print #f, "      _Version       =   393216"
   Print #f, "      PromptInclude  =   0   'False"
   Select Case GrpTipoCampo(.Fields(i).Type)
      Case 1: Print #f, "      Mask = " & """" & String(.Fields(i).Size, "9") & """"
      Case 2
         Print #f, "      Mask = " & """" & "##/##/####" & """"
         Print #f, "      Format = " & """" & "dd/mm/yyyy" & """"
      Case 3: Print #f, "      Mask = " & """" & String(.Fields(i).Size, "A") & """"
   End Select
   Print #f, "      Mask = " & """" & String(.Fields(i).Size, "9") & """"
   Print #f, "      PromptChar = " & """" & "_" & """"
   Print #f, "   End"
   '**********
   '* CheckBox
   '**********
Case "CheckBox"
   Print #f, "   Begin VB.CheckBox Chk" & .Fields(i).Name
   Print #f, "      Caption = """ & .Fields(i).Name & """"
   Print #f, "      Height = 255"
   Print #f, "      Left = " & CStr(1680 + ll& + PosX&)
   'Print #f, "      Index = 0"
   Print #f, "      TabIndex = " & CStr(TabInd%)
   Print #f, "      Top = " & CStr(TopCtrl)
   Print #f, "      Width = 975"
   Print #f, "   End"
   '**********
   '* OptionButton
   '**********
Case "OptionButton"
   Print #f, "   Begin VB.OptionButton Opt" & .Fields(i).Name
   Print #f, "      Caption = """ & .Fields(i).Name & """"
   Print #f, "      Height = 375"
   Print #f, "      Left = " & CStr(1680 + ll& + PosX&)
   Print #f, "      TabIndex = " & CStr(TabInd%)
   Print #f, "      Top = " & CStr(TopCtrl)
   Print #f, "      Width = 1215"
   Print #f, "   End"
   '**********
   '* ComboBox
   '**********
Case "ComboBox"
   Print #f, "   Begin VB.ComboBox Cmb" & .Fields(i).Name
   Print #f, "      Height = 315"
   Print #f, "      Left = " & CStr(1680 + ll& + PosX&)
   'Print #f, "      Index = 0"
   Print #f, "      TabIndex = " & CStr(TabInd%)
   Print #f, "      Text = """ & .Fields(i).Name & """"
   Print #f, "      Top = " & CStr(TopCtrl)
   Print #f, "      Width = 1455"
   Print #f, "   End"
End Select
'**********
'* Label
'**********
If UCase(Mid(.Fields(i).Name, 1, 2)) = "ID" Then
Print #f, "   Begin VB.Label LblId"
'               Print #f, "      MouseIcon       =   " & """" & Table & ".frx" & ":003C"
'               Print #f, "      MousePointer = 99       'Custom"
'               Print #f, "      ForeColor = &H00FF0000&"
Print #f, "      ForeColor = &H00000000&"
If j% > 0 Then
   Print #f, "      Index = " & CStr(j%)
End If
j = j + 1
Else
Print #f, "   Begin VB.Label Lbl"
Print #f, "      ForeColor = &H0&"
If k > 0 Then
   Print #f, "      Index = " & CStr(k%)
End If
k = k + 1
End If
Print #f, "      Caption = " & """" & .Fields(i).Name & """"
Print #f, "      AutoSize = -1           'True"
Print #f, "      BackStyle = 0          'Transparent"
Print #f, "      BeginProperty Font"
   Print #f, "         Name           =  " & """" & "Times New Roman" & """"
   Print #f, "         Size           =  9.75"
   Print #f, "         Charset        =  0"
   Print #f, "         Weight         =  700"
   Print #f, "         Underline      =  0       'False"
   Print #f, "         Italic         =  0       'False"
   Print #f, "         Strikethrough  =  0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Height = 225"
   
   Print #f, "      Left = " & CStr(ll& + PosX&)
   Print #f, "      TabIndex = " & CStr(TabInd% - 1)
   Print #f, "      Top = " & CStr(TopCtrl)
   Print #f, "      WhatsThisHelpID = 10246"
   Print #f, "   End"
End If
TabInd% = TabInd% + 2
'TopCtrl = TopCtrl + 320
'* PosX& = Width + Left
PosX& = PosX& + CStr(1680 + ll&) + (160 * .Fields(i).Size)
'         TopCtrl = IIf(PosX& < MaxW, TopCtrl, TopCtrl + 320)
'         PosX& = IIf(PosX& < MaxW, PosX&, 0)
Next
'* Se ultrapassar Limite de Linhas da Tela (19),
'* Exibir Label com o número de Campos Existentes
If k + j > 19 Then
Print #f, "   Begin VB.Label Lbl"
Print #f, "      ForeColor = &H0&"
If k > 0 Then
   Print #f, "      Index = " & CStr(k%)
End If
Print #f, "      Caption = " & """" & CStr(k + j) & " Campos" & """"
Print #f, "      AutoSize = -1           'True"
Print #f, "      BackStyle = 0          'Transparent"
Print #f, "      BeginProperty Font"
   Print #f, "         Name           =  " & """" & "Times New Roman" & """"
   Print #f, "         Size           =  9.75"
   Print #f, "         Charset        =  0"
   Print #f, "         Weight         =  700"
   Print #f, "         Underline      =  0       'False"
   Print #f, "         Italic         =  0       'False"
   Print #f, "         Strikethrough  =  0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Height = 225"
   Print #f, "      Left = 0"
   Print #f, "      TabIndex = " & CStr(TabInd% - 1)
   Print #f, "      Top = 0"
   Print #f, "      WhatsThisHelpID = 10246"
   Print #f, "   End"
End If
End With
End Sub
Public Sub MontaCmdOper(f&, Optional Qtd = 4)
   '**********
   '* Botão 1
   '**********
   Print #f, "   Begin Threed.SSCommand CmdOper"
   Print #f, "      Height = 375"
   Print #f, "      Index = 0"
   Print #f, "      Left = 720"
   Print #f, "      TabIndex = 41"
   Print #f, "      Top = 5640"
   Print #f, "      Width = 1215"
   Print #f, "      _Version        =   65536"
   Print #f, "      _ExtentX        =   1667"
   Print #f, "      _ExtentY        =   714"
   Print #f, "      _StockProps     =   78"
   Print #f, "      Caption = " & """"; "&Incluir" & """"
   Print #f, "      ForeColor = 32768"
Print #f, "      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}"
   Print #f, "         Name = " & """" & "Times New Roman" & """"
   Print #f, "         Size           = 14.25"
   Print #f, "         Charset        = 0"
   Print #f, "         Weight         = 700"
   Print #f, "         Underline      = 0       'False"
   Print #f, "         Italic         = 0       'False"
   Print #f, "         Strikethrough  = 0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Enabled         = 0      'False"
   Print #f, "      Font3D          = 3"
   Print #f, "      RoundedCorners  = 0      'False"
   Print #f, "   End"
   '**********
   '* Botão 2
   '**********
   Print #f, "   Begin Threed.SSCommand CmdOper"
   Print #f, "      Height = 375"
   Print #f, "      Index = 1"
   Print #f, "      Left = 2520"
   Print #f, "      TabIndex = 42"
   Print #f, "      Top = 5640"
   Print #f, "      Width = 1215"
   Print #f, "      _Version        =   65536"
   Print #f, "      _ExtentX        =   1667"
   Print #f, "      _ExtentY        =   714"
   Print #f, "      _StockProps     =   78"
   Print #f, "      Caption = " & """" & "&Alterar" & """"
   Print #f, "      ForeColor = 8388608"
Print #f, "      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}"
   Print #f, "         Name = " & """" & "Times New Roman" & """"
   Print #f, "         Size           = 14.25"
   Print #f, "         Charset        = 0"
   Print #f, "         Weight         = 700"
   Print #f, "         Underline      = 0       'False"
   Print #f, "         Italic         = 0       'False"
   Print #f, "         Strikethrough  = 0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Enabled         = 0      'False"
   Print #f, "      Font3D          = 3"
   Print #f, "      RoundedCorners  = 0      'False"
   Print #f, "   End"
   '**********
   '* Botão 3
   '**********
   Print #f, "   Begin Threed.SSCommand CmdOper"
   Print #f, "      Height = 375"
   Print #f, "      Index = 2"
   Print #f, "      Left = 5040"
   Print #f, "      TabIndex = 43"
   Print #f, "      Top = 5640"
   Print #f, "      Width = 1215"
   Print #f, "      _Version        =   65536"
   Print #f, "      _ExtentX        =   1667"
   Print #f, "      _ExtentY        =   714"
   Print #f, "      _StockProps     =   78"
   Print #f, "      Caption = " & """" & "&Excluir" & """"
   Print #f, "      ForeColor = 8388608"
Print #f, "      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}"
   Print #f, "         Name = " & """" & "Times New Roman" & """"
   Print #f, "         Size           = 14.25"
   Print #f, "         Charset        = 0"
   Print #f, "         Weight         = 700"
   Print #f, "         Underline      = 0       'False"
   Print #f, "         Italic         = 0       'False"
   Print #f, "         Strikethrough  = 0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Enabled         = 0      'False"
   Print #f, "      Font3D          = 3"
   Print #f, "      RoundedCorners  = 0      'False"
   Print #f, "   End"
   '**********
   '* Botão 4
   '**********
   Print #f, "   Begin Threed.SSCommand CmdOper"
   Print #f, "      Height = 375"
   Print #f, "      Index = 3"
   Print #f, "      Left = 7440"
   Print #f, "      TabIndex = 44"
   Print #f, "      Top = 5640"
   Print #f, "      Width = 1215"
   Print #f, "      _Version        =   65536"
   Print #f, "      _ExtentX        =   1667"
   Print #f, "      _ExtentY        =   714"
   Print #f, "      _StockProps     =   78"
   Print #f, "      Caption = " & """"; "&Sair" & """"
   Print #f, "      ForeColor = 128"
Print #f, "      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}"
   Print #f, "         Name = " & """" & "Times New Roman" & """"
   Print #f, "         Size           = 14.25"
   Print #f, "         Charset        = 0"
   Print #f, "         Weight         = 700"
   Print #f, "         Underline      = 0       'False"
   Print #f, "         Italic         = 0       'False"
   Print #f, "         Strikethrough  = 0       'False"
   Print #f, "      EndProperty"
   Print #f, "      Enabled         = 0      'False"
   Print #f, "      Font3D          = 3"
   Print #f, "      RoundedCorners  = 0      'False"
   Print #f, "   End"
   '**********
   '* Frame
   '**********
   Print #f, "   Begin VB.Label LblFrme"
   Print #f, "      BackStyle = 0          'Transparent"
   Print #f, "      BorderStyle = 1        'Fixed Single"
   Print #f, "      Height = 615"
   Print #f, "      Left = 600"
   Print #f, "      TabIndex = 51"
   Print #f, "      Top = 5520"
   Print #f, "      WhatsThisHelpID = 10299"
   Print #f, "      Width = 7935"
   Print #f, "   End"
End Sub
Public Sub MontaImgFundo(f)
   Print #f, "   Begin VB.Image ImgFundo"
   Print #f, "      BorderStyle = 1        'Fixed Single"
   Print #f, "      Height = 990"
   Print #f, "      Left = 0"
   Print #f, "      Top = 0"
   Print #f, "      WhatsThisHelpID = 10244"
   Print #f, "      Width = 990"
   Print #f, "   End"
End Sub
Public Sub Formata_TabClass()
   Dim i As Integer
   Dim Pos As Integer
   
   Me.LstTabCls.Clear
   If XDb.isADO Then
      For i = 1 To XDb.Tables.Count
         If Not XDb.Tables(i).isSys Then
            UserDB = XDb.Tables(i).Owner
            If LocalizarCombo(Me.CmbOwner, UserDB, False) < 0 Then
               Me.CmbOwner.AddItem UserDB
            End If
         End If
      Next
   Else
      For i = 0 To XDb.dBase.TableDefs.Count - 1
         If (XDb.dBase.TableDefs(i).Attributes And dbSystemObject) = 0 Then
            Pos = InStr(XDb.dBase.TableDefs(i).Name, ".")
            If Pos > 0 Then
               UserDB = Mid(XDb.dBase.TableDefs(i).Name, 1, Pos - 1)
               If LocalizarCombo(Me.CmbOwner, UserDB, False) < 0 Then
                  Me.CmbOwner.AddItem UserDB
               End If
               '               Me.LstTabCls.AddItem Mid(xDb.dBase.TableDefs(i).Name, Pos + 1)
               '            Else
               '               Me.LstTabCls.AddItem xDb.dBase.TableDefs(i).Name
            End If
         End If
      Next
   End If
   
   If Me.CmbOwner.ListCount > 0 Then
      Me.CmbOwner.ListIndex = 0
   Else
      UserDB = ""
      Call MontarLstTabClss
   End If
   Call LstTabCls_Click
   Call CarregaLstOp
   Me.CmbOwner.Visible = (Me.CmbOwner.ListCount > 0)
   Me.Lbl(5).Visible = (Me.CmbOwner.ListCount > 0)
   
End Sub
Public Sub Formata_TabForm()
   '* Formatar Grid
   Me.GrdCampos.Cols = 4
   Me.GrdCampos.ColWidth(0) = 300
   Me.GrdCampos.ColWidth(1) = 1600
   Me.GrdCampos.ColWidth(2) = 1600
   Me.GrdCampos.ColWidth(3) = 1000
   Me.GrdCampos.Width = 4840
   Set MyGrd = New MSGrid
   With MyGrd
'xxx      Set .Grd = Me.GrdCampos
      .MaxLin = 23
      .CollLin.Add Me.ChkCampo, "0"
      .CollLin.Add Me.TxtCampo, "1"
      .CollLin.Add Me.CmbCtrl, "2"
      .CollLin.Add Me.CmbNotNull, "3"
'xxx      .Set_Linha_Item
   End With
   
   '* Montar Combo de Tabelas
   Call MontaLstTabForm
   'Call GrdCampos_RowColChange
End Sub
Public Sub MontarLstTabClss()
   Dim i As Integer, Pos As Integer
   Dim MyOwner As String
   Me.LstTabCls.Clear
   With XDb
      If XDb.isADO Then
         For i = 1 To .Tables.Count
            If Not .Tables(i).isSys Then
               If UserDB = .Tables(i).Owner Then
                  Me.LstTabCls.AddItem .Tables(i).NOME
               End If
            End If
         Next
      Else
         For i = 0 To .dBase.TableDefs.Count - 1
            If (.dBase.TableDefs(i).Attributes And dbSystemObject) = 0 Then
               Pos = InStr(.dBase.TableDefs(i).Name, ".")
               If Pos > 0 Then
                  MyOwner = Mid(.dBase.TableDefs(i).Name, 1, Pos - 1)
               Else
                  MyOwner = ""
               End If
               If UserDB = MyOwner Then
                  Me.LstTabCls.AddItem Mid(.dBase.TableDefs(i).Name, Pos + 1)
               End If
            End If
         Next
      End If
   End With
End Sub
Public Sub CarregaLstOp()
   With Me.LstOp
      .Clear
      .AddItem "Link Com XBanco"                '* 0
      .AddItem "Propriedade 'TipoQuery'"        '* 1
      .AddItem "Propriedade 'ItensExcluidos'"   '* 2
      .AddItem "Propriedade 'Existe'"           '* 3
      .AddItem "Classe Vb6.0"                   '* 4
      .AddItem "Classe Global"                  '* 5
      .AddItem "Propriedade 'IsDirt'"           '* 6
   End With
   Me.LstOp.Selected(0) = True
   Me.LstOp.Selected(4) = True
   Me.LstOp.Selected(5) = True
   Me.LstOp.Selected(6) = True
   
   Call CarregaOPs
End Sub
Public Sub CarregaOPs()
   bComDLL = Me.LstOp.Selected(0)
   bTipoQuery = Me.LstOp.Selected(1)
   bItensExcluidos = Me.LstOp.Selected(2)
   bExiste = Me.LstOp.Selected(3)
   bVB6 = Me.LstOp.Selected(4)
   bGlobalClass = Me.LstOp.Selected(5)
   bisDirt = Me.LstOp.Selected(6)
   NmDbObj = IIf(Trim(Me.TxtNmDbObj.Text) = "", "Dbase", Trim(Me.TxtNmDbObj))
   
   DrvLocal = "C:\Tmp\Classes\"
   On Error Resume Next
   Call Kill(DrvLocal & "*.*")
End Sub
Public Sub CriarSuperClasse(CLASSE$)
   Dim Arq$, Drv$, SuperClass$
   On Error Resume Next
   Drv$ = IIf(Me.TxtDrvDest = "", "C:\TMP\", Me.TxtDrvDest.Tag)
   SuperClass$ = "BANCO"
   Arq$ = SuperClass$ & ".cls"
   Call Del(Drv$ & SuperClass$)
   Open Drv & Arq$ For Output As #1
   Print #1, "VERSION 1.0 CLASS"
   Print #1, "BEGIN"
   Print #1, "  MultiUse = -1  'True"
   If bVB6 Then
      Print #1, "  Persistable = 0  'NotPersistable"
      Print #1, "  DataBindingBehavior = 0  'vbNone"
      Print #1, "  DataSourceBehavior = 0   'vbNone"
      Print #1, "  MTSTransactionMode = 0   'NotAnMTSObject"
   End If
   Print #1, "END"
   Print #1, "Attribute VB_Name = """ & SuperClass$ & """"
   Print #1, "Attribute VB_GlobalNameSpace = " & IIf(bGlobalClass, "True", "False")
   Print #1, "Attribute VB_Creatable = True"
   Print #1, "Attribute VB_PredeclaredId = False"
   Print #1, "Attribute VB_Exposed = " & IIf(bGlobalClass, "True", "False")
   Print #1, "Attribute VB_Ext_KEY = """ & "SavedWithClassBuilder"" ,""Yes"""
   If Me.OptInPrj(0) Then
      Print #1, "Attribute VB_Ext_KEY = """ & "Top_Level"" ,""No"""
   Else
      Print #1, "Attribute VB_Ext_KEY = """ & "Top_Level"" ,""Yes"""
   End If
   Print #1, "Attribute VB_Ext_KEY = ""Member0"", """ & CLASSE$ & """"
   'Print #1, "'local variable(s) to hold property value(s)"
   Print #1, "Option Explicit "
   Print #1, "Private mvar" & NmDbObj & " As Object "
   Print #1, " "
   Print #1, "Private mvar" & CLASSE$ & " As " & CLASSE$
Print #1, "Public Property Get " & CLASSE$ & "() As " & CLASSE$
   Print #1, "   If mvar" & CLASSE$ & " Is Nothing Then"
   Print #1, "      Set mvar" & CLASSE$ & " = New " & CLASSE$
   Print #1, "      mvar" & CLASSE$ & "." & NmDbObj & " = mvar" & NmDbObj
   Print #1, "   End If"
   Print #1, "   Set " & CLASSE$ & " = mvar" & CLASSE$
   Print #1, "End Property"
Print #1, "Public Property Set " & CLASSE$ & "(vData As " & CLASSE$ & ")"
   Print #1, "   Set mvar" & CLASSE$ & " = vData"
   Print #1, "End Property"
Print #1, "Public Property Set " & NmDbObj & "(ByVal vData As Object)"
   Print #1, "   Set mvar" & NmDbObj & " = vData"
   Print #1, "End Property"
Print #1, "Public Property Let " & NmDbObj & "(ByVal vData As Object)"
   Print #1, "   Set mvar" & NmDbObj & " = vData"
   Print #1, "End Property"
Print #1, "Public Property Get " & NmDbObj & "() As Object"
   Print #1, "   Set " & NmDbObj & " = mvar" & NmDbObj
   Print #1, "End Property"
Print #1, "Private Sub Class_Terminate()"
   Print #1, "  Set mvar" & CLASSE$ & " = Nothing"
   Print #1, "End Sub"
   Close #1
End Sub
Public Sub MontarSuperClasse(ByRef CLASSE$)
   Dim ExtKEY As Boolean, ExtKEY_In As Boolean, ExtKEY_Cont%
   Dim Mvar As Boolean, Mar_In As Boolean
   Dim Terminate As Boolean, Terminate_In As Boolean
   Dim TextLine$, SuperClass$, Drv$
   Dim Mvar_In As Boolean
   
   Drv$ = IIf(Me.TxtDrvDest = "", "C:\TMP\", Me.TxtDrvDest.Tag)
   SuperClass$ = IIf(Me.TxtSuperClasse = "", "BANCO.CLS", Me.TxtSuperClasse.Tag)
   
   Call SetHourglass(hWnd)
   Call Del(Drv$ & "TOOL.TMP")
   If FileExists(Drv & SuperClass$) Then
      Open Drv & SuperClass$ For Input As #1
      Do While Not EOF(1)
         Line Input #1, TextLine
         If InStr(TextLine, CLASSE$) <> 0 Then
            Call SetDefault(hWnd)
            Exit Sub
         End If
      Loop
      Close #1 ' Close file.
   Else
      '* Criar SuperClasse
      Call CriarSuperClasse(CLASSE$)
      Call SetDefault(hWnd)
      Exit Sub
   End If
   ExtKEY_Cont% = -2
   Open Drv & SuperClass$ For Input As #1
   Open Drv & "TOOL.TMP" For Output As #2
   Do While Not EOF(1)
      Line Input #1, TextLine
      If Not Terminate And Mvar Then
         If InStr(TextLine, "Private Sub Class_Terminate()") = 0 Then
               If Terminate_In Then
                  Print #2, "   Set mvar" & CLASSE$ & " = Nothing"
                  Print #2, TextLine
                  Terminate_In = False
               Else
                  Print #2, TextLine
               End If
            Else
               Print #2, TextLine
               Terminate_In = True
            End If
         End If
         '* Private mvarTB_PAIS As TB_PAIS
         If Not Mvar And ExtKEY Then
            If InStr(TextLine, "Private mvar") = 0 Then
               If Mvar_In Then
                  Print #2, "Private mvar" & CLASSE$ & " As " & CLASSE$
                  Print #2, "Public Property Get " & CLASSE$ & "() As " & CLASSE$
                  Print #2, "   If mvar" & CLASSE$ & " Is Nothing Then"
                  Print #2, "      Set mvar" & CLASSE$ & " = New " & CLASSE$
                  Print #2, "      mvar" & CLASSE$ & "." & NmDbObj & " = mvar" & NmDbObj
                  Print #2, "   End If"
                  Print #2, "   Set " & CLASSE$ & " = mvar" & CLASSE$
                  Print #2, "End Property"
                  Print #2, "Public Property Set " & CLASSE$ & "(vData As " & CLASSE$ & ")"
                  Print #2, "   Set mvar" & CLASSE$ & " = vData"
                  Print #2, "End Property"
                  Print #2, TextLine
                  Mvar_In = False
                  Mvar = True
               Else
                  Print #2, TextLine
               End If
            Else
               Print #2, TextLine
               Mvar_In = True
            End If
         End If
         '* Attribute VB_Ext_KEY = "Member3" ,"TB_PAIS"
         If Not ExtKEY Then
            If InStr(TextLine, "VB_Ext_KEY") = 0 Then
               If ExtKEY_In Then
                  Print #2, "Attribute VB_Ext_KEY = ""Member" & CStr(ExtKEY_Cont) & """, """ & CLASSE$ & """"
                  Print #2, TextLine
                  ExtKEY_In = False
                  ExtKEY = True
               Else
                  Print #2, TextLine
               End If
            Else
               Print #2, TextLine
               ExtKEY_Cont = ExtKEY_Cont + 1
               ExtKEY_In = True
            End If
         End If
   Loop
   Close #1
   Close #2
   Call Del(Drv & SuperClass$)
   Call Copy(Drv & "TOOL.TMP", Drv & SuperClass$)
   Call SetDefault(hWnd)
End Sub
Public Function ValidaCampos()
   ValidaCampos = False
   If Me.LstTabCls.SelCount = 0 Then
      Call ExibirAviso("Escolha pelo menos uma tabela.", LoadMsg(1))
      Me.LstTabCls.SetFocus
      Exit Function
   End If
   ValidaCampos = True
End Function
Private Sub TxtCampo_KeyUp(KeyCode As Integer, Shift As Integer)
   Call MyGrd.UpDownGrid(KeyCode, Shift)
End Sub
Public Sub DefineArrayID(Optional TxtAux)
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   
   If IsMissing(TxtAux) Then TxtAux = Tabela
   
   If XDb.isADO Then
      With XDb.Tables(TxtAux)
         '* Define Chave
         Set .PrimaryKey = Nothing
         k = .PrimaryKey.Count
         ReDim ArrId(k)
         If .PrimaryKey.Count = 0 Then
            ArrId(0) = Empty
         Else
            For i = 0 To .PrimaryKey.Count - 1
               ArrId(i) = .PrimaryKey(i + 1).NOME
            Next
         End If
      End With
   Else
      With XDb.dBase.TableDefs(TxtAux)
         '* Define Chave
         For i = 0 To .Indexes.Count - 1
            If .Indexes(i).Primary Then
               ReDim ArrId(.Indexes(i).Fields.Count - 1)
               For j = 0 To .Indexes(i).Fields.Count - 1
                  ArrId(j) = .Indexes(i).Fields(j).Name
               Next
               Exit For
            Else
               ReDim ArrId(0)
               ArrId(0) = Empty
            End If
         Next
      End With
   End If
End Sub

