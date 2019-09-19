VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmMontaCls 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Montagem padrão de um arquivo .Cls"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6135
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox TxtNmDbObj 
      Height          =   285
      Left            =   5520
      TabIndex        =   18
      Text            =   "XDb"
      Top             =   2280
      Width           =   1935
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
      ItemData        =   "MontaCls.frx":0000
      Left            =   5520
      List            =   "MontaCls.frx":0007
      Style           =   1  'Checkbox
      TabIndex        =   17
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CheckBox ChkSelectAll 
      Caption         =   "Selecionar Todos os Itens"
      Height          =   200
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   200
   End
   Begin VB.ComboBox CmbOwner 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton CmdCarregaVetor 
      Caption         =   "Carrega Vetor"
      Height          =   375
      Left            =   7800
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
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
      TabIndex        =   11
      Top             =   600
      Value           =   -1  'True
      Width           =   735
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
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox TxtSuperClasse 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   5520
      LinkTimeout     =   30
      TabIndex        =   8
      Top             =   1560
      Width           =   3555
   End
   Begin VB.ListBox LstCampoCls 
      Height          =   5235
      ItemData        =   "MontaCls.frx":002A
      Left            =   2760
      List            =   "MontaCls.frx":0031
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   600
      Width           =   2535
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
      TabIndex        =   6
      Top             =   480
      Width           =   1755
   End
   Begin VB.TextBox TxtDrvDest 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   5520
      LinkTimeout     =   30
      TabIndex        =   5
      Top             =   960
      Width           =   3555
   End
   Begin VB.ListBox LstTabCls 
      Height          =   5235
      ItemData        =   "MontaCls.frx":003D
      Left            =   120
      List            =   "MontaCls.frx":0044
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin Threed.SSCommand CmdOperCls 
      Height          =   435
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Criar Classe"
      Top             =   5400
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   767
      _StockProps     =   78
      ForeColor       =   -2147483635
      Font3D          =   3
      Picture         =   "MontaCls.frx":0051
   End
   Begin Threed.SSCommand CmdOperSair 
      Height          =   435
      Left            =   7440
      TabIndex        =   3
      Top             =   5400
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   767
      _StockProps     =   78
      ForeColor       =   12632256
      Font3D          =   4
      Picture         =   "MontaCls.frx":09CB
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
      Left            =   360
      TabIndex        =   20
      Top             =   5880
      Width           =   2085
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
      TabIndex        =   19
      Top             =   2040
      Width           =   1785
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
      TabIndex        =   16
      Top             =   60
      Width           =   660
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
      TabIndex        =   12
      Top             =   360
      Width           =   1650
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
      TabIndex        =   9
      Top             =   1320
      Width           =   1875
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
      TabIndex        =   0
      Top             =   360
      Width           =   2175
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
      TabIndex        =   1
      Top             =   360
      Width           =   1650
   End
End
Attribute VB_Name = "FrmMontaCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Suja%
Public Grd As MSFlexGrid
Public Oper%
Public UserDB$

Public bSelectAll As Boolean

Dim bComDLL As Boolean
Dim bTipoQuery As Boolean
Dim bItensExcluidos As Boolean
Dim bExiste As Boolean
Dim bVB6 As Boolean
Dim bGlobalClass As Boolean

Dim NmDbObj As String
Public Sub MontarClasse(Tabela$, DscExclusao$)
   Dim Arq$, CLASSE$
   Dim TabName$, Campos$, Txt$
   Dim Chave
   Dim ChaveOptional As String
   Dim ChaveParam As String
   Dim mvarChaves As String
   Dim isKey As Boolean
   Dim i%, j%, File&
   Dim Drv As String
   
'On Error Resume Next
   On Error GoTo Fim
      
   If UserDB <> "" Then
      TabName$ = UCase(Tabela$)
      Tabela = UserDB & "." & UCase(Tabela$)
   Else
      TabName$ = UCase(Tabela$)
      Tabela = UCase(Tabela$)
   End If
   Drv$ = IIf(Me.TxtDrvDest = "", "C:\TMP\", Me.TxtDrvDest.Tag)
   CLASSE$ = IIf(Mid(TabName$, 1, 3) = "TB_", TabName$, "TB_" & TabName$)
   Arq = CLASSE & ".cls"
   Call SetHourglass(hWnd)
   Call Del(Drv$ & Arq$)
'   AbrirTxt% = FreeFile()
   Close #1
   Open Drv & Arq For Output As #1
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
      Print #1, "Attribute VB_Name = """ & CLASSE$ & """"
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
      Print #1, "Option Explicit "
      Print #1, "Private mvar" & NmDbObj & IIf(bComDLL, " As DS_BANCO", " As Object ")
      Print #1, "Private mvarRS As Recordset "
      If bExiste Then
         Print #1, "Private mvarEXISTE As Integer "
      End If
      Print #1,
      Print #1, "Private mvarQryInsert As String"
      Print #1, "Private mvarQryUpDate As String"
      Print #1, "Private mvarQryDelete As String"
      Print #1,
      If bItensExcluidos Then
         Print #1, "Private mvarItensExcluidos As Collection"
      End If
      If bTipoQuery Then
         Print #1, "Private mvarTipoQuery As String"
      End If
      If bItensExcluidos Or bTipoQuery Then
         Print #1,
      End If
      With DB.dBase.TableDefs(Tabela)
         '* Define Chave e a seguencia de parametros opcionais
         On Error Resume Next
         For i = 0 To .Indexes.Count - 1
            If .Indexes(i).Primary Then
               ReDim Chave(.Indexes(i).Fields.Count)
               'If Err = 3365 Then 'Property value not valid for REMOTE objects.
               For j = 0 To .Indexes(i).Fields.Count - 1
                  Chave(j) = .Indexes(i).Fields(j).Name
               Next
               Exit For
            End If
         Next
         ChaveOptional = ""
         ChaveParam = ""
         mvarChaves = ""
         If Not IsEmpty(Chave) Then
            For i = LBound(Chave) To UBound(Chave) - 1
               ChaveOptional = ChaveOptional & IIf(i = LBound(Chave), Space(1), ", ") & "Optional Ch_" & Chave(i)
               ChaveParam = ChaveParam & IIf(i = 0, Space(1), ", ") & "Ch_" & Chave(i) & " As String"
               mvarChaves = mvarChaves & IIf(i = 0, Space(1), ", ") & "mvar" & Chave(i)
            Next

         End If
         
         On Error GoTo Fim
         '* Define Campos
         j = 0
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField Then
               Select Case GrpTipoCampo(.Fields(i).Type)
                  Case 1: Txt = " As Double" '* Numérico
                  Case 2: Txt = " As String" '* Data
                  Case 3: Txt = " As String" '* Caracter
               End Select
               Print #1, "Private mvar" & .Fields(i).Name & Txt
               If (i Mod 5) = 0 And i <> 0 Then
                  Campos = Campos & vbNewLine
                  Campos = Campos & "   Sql = Sql & """ & IIf(j = 0, "", ", ") & .Fields(i).Name
               Else
                  Campos = Campos & IIf(j = 0, "", ", ") & .Fields(i).Name
               End If
               j = 1
            End If
         Next
         
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField Then
               Select Case GrpTipoCampo(.Fields(i).Type)
                  Case 1: Txt = " As Double" '* Numérico
                  Case 2: Txt = " As String" '* Data
                  Case 3: Txt = " As String" '* Caracter
               End Select
               Print #1, "Public Property Let " & .Fields(i).Name & "(ByVal vData" & Txt & ")"
               Print #1, "   mvar" & .Fields(i).Name & " = vData"
               Print #1, "End Property"
               Print #1, "Public Property Get " & .Fields(i).Name & "()" & Txt
               Print #1, "   " & .Fields(i).Name & " = mvar" & .Fields(i).Name
               Print #1, "End Property"
            End If
         Next
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
         Print #1, "Public Property Set " & NmDbObj & "(ByVal vData As Object)"
         Print #1, "   Set mvar" & NmDbObj & " = vData"
         Print #1, "End Property"
         Print #1, "Public Property Let " & NmDbObj & "(ByVal vData As Object)"
         Print #1, "   Set mvar" & NmDbObj & " = vData"
         Print #1, "End Property"
         Print #1, "Public Property Get " & NmDbObj & "() As Object"
         Print #1, "   Set " & NmDbObj & " = mvar" & NmDbObj
         Print #1, "End Property"
         Print #1, "Public Property Get RS() As Recordset"
         Print #1, "   Set RS = mvarRS"
         Print #1, "End Property"

'* QryInsert
         Print #1, "Public Property Get QryInsert() As String"
         Print #1, "   Dim Sql As String"
         Print #1, " "
         Print #1, "   Sql = """ & "insert into " & TabName & " (" & Campos & ") """
         Print #1, "   Sql = Sql & """ & " Values " & """"
         Print #1, "   Sql = Sql & """ & "(" & """"
         j = 0
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField Then
               Select Case GrpTipoCampo(.Fields(i).Type)
                  Case 1: Print #1, "   Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ") & "SqlNum(mvar" & .Fields(i).Name & ")"
                  Case 2: Print #1, "   Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ") & "SqlDate(mvar" & .Fields(i).Name & ")"
                  Case 3: Print #1, "   Sql = Sql & " & IIf(j = 0, Space(1), """, """ & " & ") & "SqlStr(mvar" & .Fields(i).Name & ")"
               End Select
               j = 1
            End If
         Next
         Print #1, "   Sql = Sql & """ & ")" & """"
         Print #1, "   mvarQryInsert = Sql"
         Print #1, "   QryInsert = mvarQryInsert"
         Print #1, "End Property"
'* QryDelete
         Print #1, "Public Property Get QryDelete(" & ChaveOptional & ") As String"
         Print #1, "   Dim Sql As String"
         Print #1, " "
'         For i = LBound(Chave) To UBound(Chave) - 1
'            If i = LBound(Chave) Then
'               Txt = "   If isMissing(Ch_" & Chave(i) & ")"
'            Else
'               Txt = Txt & " And isMissing(Ch_" & Chave(i) & ")"
'            End If
'         Next
'         Txt = Txt & " Then"
'         Print #1, Txt
'         For i = LBound(Chave) To UBound(Chave) - 1
'            Print #1, "      If Trim(mvar" & Chave(i) & ") = """" Then Exit Property"
'         Next
'         For i = LBound(Chave) To UBound(Chave) - 1
'            Print #1, "      Ch_" & Chave(i) & " = mvar" & Chave(i)
'         Next
'         Print #1, "   End If"
         Print #1, "   Sql = """ & "Delete From " & TabName & """"
         Print #1, "   Sql = Sql & """ & " Where " & """"
         If Not IsEmpty(Chave) Then
            For i = LBound(Chave) To UBound(Chave) - 1
               Txt = "   if Not isMissing(Ch_" & Chave(i) & ") Then "
               Txt = Txt & "Sql = Sql & """ & Chave(i) & " = """
               Select Case GrpTipoCampo(.Fields(Chave(i)).Type)
                  Case 1: Txt = Txt & " & SqlNum(Cstr(Ch_" & Chave(i) & "))"
                  Case 2: Txt = Txt & " & SqlDate(Cstr(Ch_" & Chave(i) & "))"
                  Case 3: Txt = Txt & " & SqlStr(Cstr(Ch_" & Chave(i) & "))"
               End Select
               Txt = Txt & " & "" AND """
               Print #1, Txt
            Next
         End If
         Print #1, "   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len("" AND "")))"
         Print #1, "   mvarQryDelete = Sql"
         Print #1, "   QryDelete = mvarQryDelete"
         Print #1, "End Property"
'* QryUpDate
         Print #1, "Public Property Get QryUpDate() As String"
         Print #1, "   Dim Sql As String"
         Print #1, " "
         Print #1, "   Sql = """ & "update " & TabName & " set " & """"
         j = 0
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField Then
               Txt = "   Sql = Sql & """ & IIf(j = 0, Space(1), " , ") & .Fields(i).Name & " = """
               Select Case GrpTipoCampo(.Fields(i).Type)
                  Case 1:  Txt = Txt & " & SqlNum(mvar" & .Fields(i).Name & ")"
                  Case 2: Txt = Txt & " & SqlDate(mvar" & .Fields(i).Name & ")"
                  Case 3: Txt = Txt & " & SqlStr(mvar" & .Fields(i).Name & ")"
               End Select
               Print #1, Txt
               j = 1
            End If
         Next
         Print #1, "   Sql = Sql & """ & " Where "
         If Not IsEmpty(Chave) Then
            For i = LBound(Chave) To UBound(Chave) - 1
               Txt = "   Sql = Sql & """ & IIf(i = 0, Space(1), " and ") & Chave(i) & " = """
               Select Case GrpTipoCampo(.Fields(Chave(i)).Type)
                  Case 1:  Txt = Txt & " & mvar" & Chave(i)
                  Case 2: Txt = Txt & " & SqlDate(mvar" & Chave(i) & ")"
                  Case 3: Txt = Txt & " & SqlStr(mvar" & Chave(i) & ")"
               End Select
               Print #1, Txt
            Next
         End If
         Print #1, "   mvarQryUpDate = Sql"
         Print #1, "   QryUpDate = mvarQryUpDate"
         Print #1, "End Property"
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
'* PESQUISAR
         Print #1, "Public Function Pesquisar(" & ChaveOptional & ") As Boolean" '& IIf(bComDLL, "Integer", "Boolean")
         Print #1, "   Dim Sql As String"

'         If Not IsEmpty(Chave) Then
'            For i = LBound(Chave) To UBound(Chave) - 1
'               Print #1, "   mvar" & Chave(i) & " = Ch_" & Chave(i)
'            Next
'         End If
         Print #1,
         Print #1, "   Sql = """ & "select distinct " & Campos
         Print #1, "   Sql = Sql & """ & " From " & TabName
         Print #1, "   Sql = Sql & """ & " Where " & """"
         If Not IsEmpty(Chave) Then
            For i = LBound(Chave) To UBound(Chave) - 1
               Txt = "   if Not isMissing(Ch_" & Chave(i) & ") Then "
               Txt = Txt & "Sql = Sql & """ & Chave(i) & " = """
               Select Case GrpTipoCampo(.Fields(Chave(i)).Type)
                  Case 1: Txt = Txt & " & SqlNum(Cstr(Ch_" & Chave(i) & "))"
                  Case 2: Txt = Txt & " & SqlDate(Cstr(Ch_" & Chave(i) & "))"
                  Case 3: Txt = Txt & " & SqlStr(Cstr(Ch_" & Chave(i) & "))"
               End Select
               Txt = Txt & " & "" AND """
               Print #1, Txt
            Next
         End If
         Print #1, "   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len("" AND "")))"
         If bComDLL Then
            Print #1, "   Call mvar" & NmDbObj & ".AbreTabela(Sql, mvarRS)"
         Else
            Print #1, "   Set mvarRS = mvar" & NmDbObj & ".OpenRecordset(Sql, dbOpenSnapshot, dbExecDirect)"
         End If
         Print #1, "   With mvarRS"
         Print #1, "      If Not .EOF Then"
         
         j = 0
         For i = 0 To .Fields.Count - 1
            If .Fields(i).Attributes < dbSystemField Then
               If Not IsEmpty(Chave) Then
                  isKey = InArray(.Fields(i).Name, Chave) And (j = 0)
               Else
                  isKey = False
               End If
               Select Case GrpTipoCampo(.Fields(i).Type)
                  Case 1: Txt = "      mvar" & .Fields(i).Name & " = XVal(!" & .Fields(i).Name & " & """")"
                  Case 2: Txt = "      mvar" & .Fields(i).Name & " = Format(!" & .Fields(i).Name & " & """"" & ", """ & "DD/MM/YYYY" & """" & ")"
                  Case 3: Txt = "      mvar" & .Fields(i).Name & " = !" & .Fields(i).Name & " & """""
               End Select
               If bComDLL Then
                  Print #1, "   " & Txt
               Else
                  Print #1, Txt
               End If
            End If
            j = 1
         Next
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
'* INCLUIR
         If bComDLL Then
            Print #1, "Public Function Incluir(Optional ComCOMMIT = False) As Boolean"
            Print #1, "   Incluir = mvar" & NmDbObj & ".Executa(Me.QryInsert, ComCOMMIT)"
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
'* EXCLUIR
         If bComDLL Then
            Print #1, "Public Function Excluir(Optional ComCOMMIT = False) As Boolean"
            DscExclusao = IIf(DscExclusao = "", "", ", mvar" & DscExclusao)
'            If DscExclusao = "" Then
               Print #1, "   Excluir = mvar" & NmDbObj & ".Executa(Me.QryDelete(" & mvarChaves & "), ComCOMMIT)"
'            Else
'               Print #1, "      Incluir = mvar" & NmDbObj & ".Executa(Me.QryInsert)"
'               Print #1, "      Incluir = mvar" & NmDbObj & ".Executa(Me.QryInsert)"
'               Print #1, "   End if"
'            End If
            
            
         Else
            Print #1, "Public Function Excluir() As Boolean"
            Print #1, "   Dim Sql As String"
            Print #1, "   On Error GoTo Fim"
            Print #1, "   Sql = QryDelete(" & mvarChaves & ")"
            Print #1, "   mvar" & NmDbObj & ".Execute Sql, dbExecDirect"
            Print #1, "   Excluir = True"
            Print #1, "Exit Function"
            Print #1, "Fim:"
            Print #1, "   Excluir = False"
            Print #1, "   msgInformacao " & """" & "Problema na exclusão." & """" & " & vbNewLine & Errors(0).Description"
         End If
         Print #1, "End Function"
'* ALTERAR
         If bComDLL Then
            Print #1, "Public Function Alterar(Optional ComCOMMIT = False) As Boolean"
            Print #1, "   Alterar =  mvar" & NmDbObj & ".Executa(Me.QryUpDate, ComCOMMIT)"
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

'* ALTERARCHAVE
         If bComDLL Then
            Print #1, "Public Function AlterarChave(" & ChaveParam & ", Optional ComCOMMIT = False) As Integer"
         Else
            Print #1, "Public Function AlterarChave(" & ChaveParam & ") As Integer"
         End If
         Print #1, "   Dim Sql As String"
         Print #1, " "
         If Not bComDLL Then
            Print #1, "   On Error GoTo Fim"
         End If
         Print #1, "   Sql = """ & "update " & TabName & " set " & """"
         If Not IsEmpty(Chave) Then
            For i = LBound(Chave) To UBound(Chave) - 1
               Txt = "   Sql = Sql & """ & IIf(i = 0, Space(1), " , ") & .Fields(Chave(i)).Name & " = """
               Select Case GrpTipoCampo(.Fields(Chave(i)).Type)
                  Case 1:  Txt = Txt & " & Ch_" & .Fields(Chave(i)).Name
                  Case 2: Txt = Txt & " & SqlDate(Ch_" & .Fields(Chave(i)).Name & ")"
                  Case 3: Txt = Txt & " & SqlStr(Ch_" & .Fields(Chave(i)).Name & ")"
               End Select
               Print #1, Txt
            Next
         End If
         Print #1, "   Sql = Sql & """ & " Where "
         If Not IsEmpty(Chave) Then
            For i = LBound(Chave) To UBound(Chave) - 1
               Txt = "   Sql = Sql & """ & IIf(i = 0, Space(1), " and ") & Chave(i) & " = """
               Select Case GrpTipoCampo(.Fields(Chave(i)).Type)
                  Case 1:  Txt = Txt & " & mvar" & Chave(i)
                  Case 2: Txt = Txt & " & SqlDate(mvar" & Chave(i) & ")"
                  Case 3: Txt = Txt & " & SqlStr(mvar" & Chave(i) & ")"
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
'* INITIALIZE
         If bTipoQuery Then
            Print #1, "Private Sub Class_Initialize()"
            If bTipoQuery Then
               Print #1, "   mvarTipoQuery = ""I"""
            End If
      '         Print #1, "   If mvar" & NmDbObj & " Is Nothing Then"
      '         Print #1, "      set mvar" & NmDbObj & " = BANCO." & NmDbObj
      '         Print #1, "   End If"
            Print #1, "End Sub"
         End If
'* TERMINATE
         Print #1, "Private Sub Class_Terminate()"
         Print #1, "   Set mvar" & NmDbObj & " = Nothing"
         Print #1, "   Set mvarRS = Nothing"
         If bItensExcluidos Then
            Print #1, "   Set mvarItensExcluidos = Nothing"
         End If
         Print #1, "End Sub"
      End With
   Close #1
   
   Call MontarSuperClasse(CLASSE$)
   
   Call SetDefault(hWnd)
Exit Sub
Fim:
   'If Err = 55 Then
   Close #1
   Call ShowError
'   MsgBox CStr(Err) & " - " & CStr(Error)
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
      Print #1, "      Set mvar" & CLASSE$ & " = New T" & CLASSE$
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

Public Function ValidaCampos()
   ValidaCampos = False
   If Me.LstTabCls.SelCount = 0 Then
      Call ExibirAviso("Escolha pelo menos uma tabela.", LoadMsg(1))
      Me.LstTabCls.SetFocus
      Exit Function
   End If
   ValidaCampos = True
End Function

Private Sub ChkSelectAll_Click()
   Dim i As Integer
   If Me.ChkSelectAll.Value = vbChecked Then
      Me.LstTabCls.Visible = False
      bSelectAll = True
      For i = Me.LstTabCls.ListCount - 1 To 0 Step -1
         Me.LstTabCls.Selected(i) = True
      Next
      Me.LstTabCls.Visible = True
   End If
   bSelectAll = False
End Sub

Private Sub CmbOwner_Click()
   Me.ChkSelectAll.Value = vbUnchecked
   UserDB = Me.CmbOwner.Text
   Call MontarLstTabClss
End Sub
Private Sub CmdDrv_Click(Index As Integer)
   Dim PATH$
   Dim Tit$, Filtro$, Arq$, Ind%
   
   Tit$ = "Find Project"
   Filtro = "Project Files (*.vbp)|*.vbp"
   Ind% = 1
   SysMdi.CmDialog.InitDir = "C:\SISTEMAS\"
   Arq$ = ProcurarArquivo(SysMdi.CmDialog, Tit$, Arq$, Filtro$, Ind%)
   If Arq$ = "" Then
      Me.OptInPrj(1).Value = True
      Exit Sub
   End If
   Me.TxtDrvDest.Text = UCase(SysMdi.CmDialog.Tag) & Arq
   Me.TxtDrvDest.Tag = UCase(SysMdi.CmDialog.Tag)
   
   Tit$ = "Find DataBase Class"
   Filtro = "Project Files (*.cls)|*.cls"
   Ind% = 1
   SysMdi.CmDialog.InitDir = Me.TxtDrvDest.Tag
   Arq$ = ProcurarArquivo(SysMdi.CmDialog, Tit$, Arq$, Filtro$, Ind%)
   If Arq$ = "" Then
      Exit Sub
   End If
   Me.TxtSuperClasse.Text = UCase(SysMdi.CmDialog.Tag) & Arq
   Me.TxtSuperClasse.Tag = Arq
End Sub

Private Sub CmdOperCls_Click()
   Dim Sql As String, DscExclusao$, i%, j%
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
            Call MontarClasse(Me.LstTabCls.List(i), DscExclusao$)
            j = j + 1
            If j = Me.LstTabCls.SelCount Then Exit For
         End If
      Next
      Call ExibirAviso(LoadMsg(34), LoadMsg(1))
   End If
   Call SetDefault(hWnd)
End Sub

Private Sub CmdCarregaVetor_Click()
   Dim Arq$, CLASSE$, Tabela$, Drv$
   Dim TabName$, Campos$, Txt$
   Dim Chave, isKey%
   Dim i%, j%, File&
   On Error GoTo Fim
   Tabela$ = Me.LstTabCls
   If UserDB <> "" Then
      TabName$ = UCase(Tabela$)
      Tabela = UserDB & "." & UCase(Tabela$)
   Else
      TabName$ = UCase(Tabela$)
      Tabela = UCase(Tabela$)
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
      Print #1, "    '" & Tabela
      With DB.dBase.TableDefs(Tabela)
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

Private Sub CmdOperSair_Click()
   Unload Me
End Sub
Private Sub Form_Activate()
   If Not DB.Conectado Then
      Unload Me
      Exit Sub
   End If
   Me.Visible = True
   Set MDIFilho = Me
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Initialize()
   Dim Arq As String
   Dim i As Integer
   With DB
      If Not .Conectado Then
         FrmOpBanco.Show vbModal
         If Not Sys.isODBC Then
            If Dir("C:\DSR\", vbDirectory) <> "" Then
               SysMdi.CmDialog.InitDir = "C:\DSR\"
            End If
            Arq$ = .Alias
            If .Alias = "" Then
               Arq$ = ProcurarArquivo(SysMdi.CmDialog, "Abrir Banco de Dados Access", , "Microsoft Access MDBs (*.mdb)|*.mdb")
               .isODBC = False
               .dbDrive = SysMdi.CmDialog.Tag
               .dbName = Arq$
            End If
            If Arq$ <> "" Then
               Call .SrvConecta(.dbDrive, .dbName, "", "", "", "")
            End If
         Else
            .isODBC = True
            frmODBCLog.Show vbModal
         End If
         For i = 2 To 5
            Arq$ = Trim(GetSetting(Sys.AppName, "Outros", "BDRecente" & CStr(i - 1), ""))
            If DB.Alias <> Arq$ And Arq$ <> "" Then
               Call SaveSetting(Sys.AppName, "Outros", "BDRecente" & CStr(i), Arq$)
            End If
         Next
         If DB.Alias <> "" Then
            Call SaveSetting(Sys.AppName, "Outros", "BDRecente1", DB.Alias)
         End If
'         If Not .Conectado Then
'            Unload Me
'         End If
      End If
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete, vbKeyBack
   End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: Unload Me
      Case Else
         If Not (Me.ActiveControl Is Me.TxtNmDbObj) Or KeyAscii = 13 Then
            KeyAscii = SendTab(Me, KeyAscii)
         End If
   End Select
End Sub
Private Sub Form_Load()
   Dim i%, Pos%
   Call SetHourglass(hWnd)
   '* Montar LISTA de Tabelas
   Me.Visible = False
   If Not DB.Conectado Then
      Exit Sub
   End If
   
   With DB.dBase
      Me.LstTabCls.Clear
      
      For i = 0 To DB.dBase.TableDefs.Count - 1
         If (.TableDefs(i).Attributes And dbSystemObject) = 0 Then
            Pos = InStr(DB.dBase.TableDefs(i).Name, ".")
            If Pos > 0 Then
               UserDB = Mid(DB.dBase.TableDefs(i).Name, 1, Pos - 1)
               If LocalizarCombo(Me.CmbOwner, UserDB, False) < 0 Then
                  Me.CmbOwner.AddItem UserDB
               End If
               '               Me.LstTabCls.AddItem Mid(DB.dBase.TableDefs(i).Name, Pos + 1)
               '            Else
               '               Me.LstTabCls.AddItem DB.dBase.TableDefs(i).Name
            End If
         End If
      Next
   End With
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
   Call ConfigForm(Me, SysMdi.Icon, FundoTela)
   Call SetDefault(hWnd)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Suja = False
   '=============
   '=  Se nenhum campo foi alterado -> SAIR
   '=============
   If Not Me.Suja Then Exit Sub
   '=============
   '=   Se não deseja salvar -> SAIR
   '=============
   If ExibirPergunta(LoadMsg(54), Me.Caption) = vbNo Then
      Exit Sub
   End If
   '=============
   '=   Verificar e validar campos
   '=============
'   If ValidaCampos Then  F_SALVAR
End Sub

Private Sub Form_Resize()
   If DB.Conectado Then
      Call PintarFundo(Me, FundoTela)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set MDIFilho = Nothing
End Sub

Private Sub Lbl_Click(Index As Integer)
   Select Case Index
      Case 7
         Me.ChkSelectAll.Value = IIf(Me.ChkSelectAll.Value = vbChecked, vbUnchecked, vbChecked)
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
   Dim Tabela$, i As Integer
'* Montar Combo de Campo de Descrição
   If bSelectAll Then Exit Sub
   Me.ChkSelectAll.Value = vbUnchecked
   Me.LstCampoCls.Clear
   If Me.LstTabCls.Selected(Me.LstTabCls.ListIndex) Then
      Me.LstCampoCls.AddItem "  -- Em Branco -- "
   End If
   If UserDB = "" Then
      Tabela = Me.LstTabCls
   Else
      Tabela = UserDB & "." & Me.LstTabCls
   End If
On Error Resume Next
   With DB.dBase.TableDefs(Tabela)
       For i = 0 To .Fields.Count - 1
         If .Fields(i).Attributes < dbSystemField Then
            Me.LstCampoCls.AddItem .Fields(i).Name
         End If
      Next
   End With
   If Me.LstTabCls.Selected(Me.LstTabCls.ListIndex) Then
      Me.LstCampoCls.Selected(Me.LstTabCls.ItemData(Me.LstTabCls.ListIndex)) = True
   End If
End Sub

Private Sub LstTabCls_ItemCheck(Item As Integer)
   Dim Tabela$, i As Integer
   If bSelectAll Then Exit Sub
   Me.LstCampoCls.Clear
   Me.LstCampoCls.AddItem "  -- Em Branco -- "
   If UserDB = "" Then
      Tabela = Me.LstTabCls
   Else
      Tabela = UserDB & "." & Me.LstTabCls
   End If
   On Error Resume Next
   With DB.dBase.TableDefs(Tabela)
      For i = 0 To .Fields.Count - 1
         Me.LstCampoCls.AddItem .Fields(i).Name
      Next
   End With
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
'   Call CentrarForm(SysMdi, Me)
   Me.Visible = True
   Me.Refresh
End Sub
Public Sub MontarLstTabClss()
   Dim i As Integer, Pos As Integer
   Dim MyOwner As String
   Me.LstTabCls.Clear
   With DB.dBase
      For i = 0 To .TableDefs.Count - 1
         If (.TableDefs(i).Attributes And dbSystemObject) = 0 Then
            Pos = InStr(.TableDefs(i).Name, ".")
            If Pos > 0 Then
               MyOwner = Mid(.TableDefs(i).Name, 1, Pos - 1)
            Else
               MyOwner = ""
            End If
            If UserDB = MyOwner Then
               Me.LstTabCls.AddItem Mid(.TableDefs(i).Name, Pos + 1)
            End If
         End If
      Next
   End With
End Sub
Public Sub CarregaLstOp()
   With Me.LstOp
      .Clear
      .AddItem "Link Com DSR100.DLL"            '* 0
      .AddItem "Propiedade 'TipoQuery'"         '* 1
      .AddItem "Propiedade 'ItensExcluidos'"    '* 2
      .AddItem "Propiedade 'Existe'"            '* 3
      .AddItem "Classe Vb6.0"                   '* 4
      .AddItem "Classe Global"                  '* 5
   End With
   Me.LstOp.Selected(0) = True
   Me.LstOp.Selected(4) = True
   Me.LstOp.Selected(5) = True
   
   Call CarregaOPs
End Sub
Public Sub CarregaOPs()
   bComDLL = Me.LstOp.Selected(0)
   bTipoQuery = Me.LstOp.Selected(1)
   bItensExcluidos = Me.LstOp.Selected(2)
   bExiste = Me.LstOp.Selected(3)
   bVB6 = Me.LstOp.Selected(4)
   bGlobalClass = Me.LstOp.Selected(5)
   NmDbObj = IIf(Trim(Me.TxtNmDbObj.Text) = "", "Dbase", Trim(Me.TxtNmDbObj))
End Sub

