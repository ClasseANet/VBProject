VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmLoto 
   AutoRedraw      =   -1  'True
   Caption         =   "Loto"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGridToExcel 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00D8E9EC&
      Picture         =   "FrmLoto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Gerar um arquivo Excel"
      Top             =   5760
      Width           =   420
   End
   Begin VB.TextBox TxtNum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "06"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CmbSorteio 
      Height          =   360
      ItemData        =   "FrmLoto.frx":06C2
      Left            =   1800
      List            =   "FrmLoto.frx":06CC
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox LstNumeros 
      BackColor       =   &H80000018&
      Height          =   5580
      ItemData        =   "FrmLoto.frx":06D8
      Left            =   120
      List            =   "FrmLoto.frx":0718
      TabIndex        =   11
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton CmdOper 
      Caption         =   "Calcular"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton CmdOper 
      Caption         =   "Sair"
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox TxtCercado 
      Height          =   315
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "05"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox TxtJogado 
      Height          =   315
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "06"
      Top             =   360
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCartao 
      Height          =   4335
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
   End
   Begin VB.Label LblGrd 
      BackStyle       =   0  'Transparent
      Caption         =   "1 / 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label LblContador 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1 / 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Nºs Cercados"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Nºs Jogados "
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Nºs de Sorteio"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmLoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ExisteArqMestre As Boolean
Public sArquivo As String
Public ArqMestre As String

Private Sub cmdGridToExcel_Click()
   Call GridToExcel(Me.GrdCartao, "Loto" & Me.CmbSorteio & Me.TxtJogado)
End Sub

Private Sub CmdOper_Click(Index As Integer)
   Select Case Index
      Case 0
         End
      Case 1
         Call CalcularCartoes
         Me.cmdGridToExcel.Enabled = True
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      SendKeys "{TAB}"
   End If
End Sub
Private Sub Form_Load()
   Dim VetAtrib(), VetInd()
   Dim i As Integer
   
   Me.CmbSorteio.ListIndex = 1
   
   sArquivo = App.Path & "\Loto.mdb"
   If FileExists(sArquivo) Then
      On Error Resume Next
      Kill sArquivo
   End If
   If Not FileExists(sArquivo) Then
      Call CriarBancoDeDados(xDb, sArquivo)
      
   Else
      Set DBEngine = Nothing
      Set WS = DBEngine.CreateWorkspace("WsEngine", "admin", "")
      Set xDb = WS.OpenDatabase(sArquivo, False, False, ";")
   End If
   Me.LstNumeros.Clear
   For i = 1 To Val(Me.TxtJogado)
      Me.LstNumeros.AddItem Right$("00" & i, 2)
   Next
   Call PosTxtNum
End Sub

Private Sub Form_Unload(Cancel As Integer)
   xDb.Close
   WS.Close
   Set xDb = Nothing
   Set WS = Nothing
            
   On Error Resume Next
   Kill sArquivo
End Sub
Private Sub LstNumeros_Click()
   Call PosTxtNum
End Sub
Private Sub LstNumeros_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
      Call PosTxtNum
   End If
End Sub
Public Sub PosTxtNum()
   If Me.TxtNum.Tag = "" Then
      Me.TxtNum.Tag = LstNumeros.ListIndex
      Me.TxtNum.Text = LstNumeros.List(LstNumeros.ListIndex)
   End If
   LstNumeros.List(Val(Me.TxtNum.Tag)) = Me.TxtNum.Text
   Me.TxtNum.Left = LstNumeros.Left
   'Me.TxtNum.Height = 220
   Me.TxtNum.Top = (240 * LstNumeros.ListIndex) + LstNumeros.Top
   Me.TxtNum.Text = LstNumeros.List(LstNumeros.ListIndex)
   Me.TxtNum.Tag = LstNumeros.ListIndex
   If Me.TxtNum.Visible Then
      Me.TxtNum.SetFocus
   '   Me.TxtNum.SelStart = 0
   '   Me.TxtNum.SelLength = Len(Me.TxtNum)
   End If
   On Error Resume Next
   If Me.TxtNum.Tag <> "-1" Then
      Me.TxtNum.Visible = True
   End If
End Sub
Private Sub TxtCercado_GotFocus()
   Me.TxtCercado.SelStart = 0
   Me.TxtCercado.SelLength = Len(Me.TxtCercado)
End Sub
Private Sub TxtCercado_LostFocus()
   TxtCercado.Text = Right$("00" & TxtCercado.Text, 2)
End Sub
Private Sub TxtJogado_GotFocus()
   Me.TxtJogado.SelStart = 0
   Me.TxtJogado.SelLength = Len(Me.TxtJogado)
End Sub
Private Sub TxtJogado_LostFocus()
   Dim i As Integer
   
   TxtJogado.Text = Right$("00" & TxtJogado.Text, 2)
   Me.LstNumeros.Clear
   For i = 1 To Val(Me.TxtJogado)
      Me.LstNumeros.AddItem Right$("00" & i, 2)
   Next
End Sub
Private Sub CmbSorteio_GotFocus()
   'Me.CmbSorteio.SelStart = 0
   'Me.CmbSorteio.SelLength = Len(Me.CmbSorteio)
End Sub
Private Sub CmbSorteio_LostFocus()
   Me.GrdCartao.Cols = Val(Me.CmbSorteio)
   CmbSorteio.Text = Right$("00" & CmbSorteio.Text, 2)
End Sub
Public Sub CalcularCartoes()
   Dim i As Double, j As Double
   Dim Comb As Double
   Dim Sql As String
   Dim Rs As DAO.Recordset
   Dim Min As Double
   Dim Max As Double
   Dim sAux As String
   Dim CollNum As Collection
   Dim Cartoes As Collection
   Dim MyCartao As Cartao
   Dim n As Variant
   Dim Pos As Collection, PosA()
   Dim Contador As Double
   Dim nAux As Double
   Dim IDAnterior As Double
   
   Screen.MousePointer = vbHourglass
   
   Me.GrdCartao.Rows = 0
   Me.LblGrd.Caption = ""
   Me.LblContador = ""
   
   Set CollNum = New Collection
   For i = 0 To Val(Me.TxtJogado) - 1
      CollNum.Add 2 ^ i
   Next
   Comb = 1
   For i = 0 To Val(Me.CmbSorteio) - 1
      Comb = Comb * (Val(Me.TxtJogado) - i)
   Next
   For i = 0 To Val(Me.CmbSorteio) - 1
      Comb = Comb / (Val(Me.CmbSorteio) - i)
   Next
   
   Set Cartoes = New Collection
   Set MyCartao = New Cartao
   
   Set Pos = New Collection
   ReDim PosA(CollNum.Count)
   For i = 1 To CollNum.Count
      Pos.Add i
      PosA(i) = 0
   Next

   Set MyCartao = New Cartao
   Contador = 0
   
   ArqMestre = App.Path & "\" & Me.CmbSorteio & Me.TxtJogado & ".mdb"
   ExisteArqMestre = FileExists(ArqMestre)
   If ExisteArqMestre Then
      'On Error Resume Next
      'xDb.Close
      'Set xDb = Nothing
      'Call CriarBancoDeDados(xDb, sArquivo)
      
      'xDbMestre.Close
      Set xDbMestre = Nothing
      Call Copy(ArqMestre, sArquivo)
      Set DBEngine = Nothing
      Set WS = DBEngine.CreateWorkspace("WsEngine", "admin", "")
      If FileExists(sArquivo) Then
         Set xDb = Nothing
         Set xDb = WS.OpenDatabase(sArquivo, False, False, ";")
      Else
         MsgBox "Arquivo abaixo não encontrado." & vbNewLine & vbNewLine & "[ " & sArquivo & " ] "
         Exit Sub
      End If
      
   Else
      Call CriarBancoDeDados(xDbMestre, ArqMestre)

      Set xDb = Nothing
      Set xDb = WS.OpenDatabase(sArquivo, False, False, ";")
            
      Sql = "Delete from SOMAS "
      Call xDb.Execute(Sql, dbFailOnError)
      
      Sql = "Delete from NUMEROS "
      Call xDb.Execute(Sql, dbFailOnError)
      
      Sql = "Delete from CARTOES "
      Call xDb.Execute(Sql, dbFailOnError)
      Call DBEngine.Idle(dbFreeLocks)
   End If
   Set xDbMestre = Nothing
   Set xDbMestre = WS.OpenDatabase(ArqMestre, False, False, ";")
   
         
   While PosA(1) <= CollNum.Count - Val(Me.CmbSorteio) + 1
      Set MyCartao = New Cartao
      
      Contador = Contador + 1
      
      Me.LblContador.Caption = CStr(Contador) & "/" & CStr(Comb)
      Me.LblContador.Refresh
      
      
      For i = 1 To Val(Me.CmbSorteio)
         MyCartao.Numeros.Add CollNum(Pos(i) + PosA(i))
      Next
      Call MyCartao.MontarSomas(Val(Me.TxtCercado))
      Call MyCartao.SomaPos(PosA, Val(Me.CmbSorteio), CollNum.Count - Val(Me.CmbSorteio))
      
      If Not ExisteArqMestre Then
         Sql = "Insert into CARTOES (IDJOGO, IDCARTAO, VALIDO, VERIFICADO )"
         Sql = Sql & " values "
         Sql = Sql & "('" & Me.CmbSorteio & Me.TxtJogado & Me.TxtCercado & "'"
         Sql = Sql & ", " & CStr(Contador)
         Sql = Sql & ", 'S',  'N')"
         
         Call xDb.Execute(Sql, dbFailOnError)
         Call xDbMestre.Execute(Sql, dbFailOnError)
         Call DBEngine.Idle(dbFreeLocks)
         
         For Each n In MyCartao.Numeros
            Sql = "Insert into NUMEROS (IDJOGO, IDCARTAO, NUMERO )"
            Sql = Sql & " values "
            Sql = Sql & "('" & Me.CmbSorteio & Me.TxtJogado & Me.TxtCercado & "'"
            Sql = Sql & ", " & CStr(Contador)
            Sql = Sql & ", " & CStr(Val(n))
            Sql = Sql & ")"
            Call xDb.Execute(Sql, dbFailOnError)
            Call xDbMestre.Execute(Sql, dbFailOnError)
            Call DBEngine.Idle(dbFreeLocks)
         Next
         For Each n In MyCartao.Somas
            Sql = "Insert into SOMAS (IDJOGO, IDCARTAO, SOMA )"
            Sql = Sql & " values "
            Sql = Sql & "('" & Me.CmbSorteio & Me.TxtJogado & Me.TxtCercado & "'"
            Sql = Sql & ", " & CStr(Contador)
            Sql = Sql & ", " & CStr(n)
            Sql = Sql & ")"
            Call xDb.Execute(Sql, dbFailOnError)
            Call xDbMestre.Execute(Sql, dbFailOnError)
            Call DBEngine.Idle(dbFreeLocks)
         Next
      End If
      
      Cartoes.Add MyCartao, CStr(Cartoes.Count + 1)
      Set MyCartao = Nothing
   Wend
   
   If Not ExisteArqMestre Then
      xDbMestre.Close
      Set xDbMestre = Nothing
   End If
   
   Min = 1
   Max = Comb
   i = Min
   While Min <= Max
      Sql = "Delete From CARTOES C"
'      Sql = Sql & " Set VALIDO = 'N'"
      If i = Min Then
         Sql = Sql & " where C.IDCARTAO > " & CStr(i)
         Sql = Sql & " and  C.IDCARTAO <= " & CStr(Max)
      Else
         Sql = Sql & " where C.IDCARTAO < " & CStr(i)
         Sql = Sql & " and  C.IDCARTAO >= " & CStr(Min)
      End If
      Sql = Sql & " and C.IDCARTAO in ("
      Sql = Sql & "    Select S.IDCARTAO from SOMAS S "
      If i = Min Then
         Sql = Sql & " where S.IDCARTAO > " & CStr(i)
         Sql = Sql & " and  S.IDCARTAO <= " & CStr(Max)
      Else
         Sql = Sql & " where S.IDCARTAO < " & CStr(i)
         Sql = Sql & " and  S.IDCARTAO >= " & CStr(Min)
      End If
      Sql = Sql & "    and S.SOMA in ("
      For Each n In Cartoes(i).Somas
         Sql = Sql & CStr(n)
         If n <> Cartoes(i).Somas(Cartoes(i).Somas.Count) Then
            Sql = Sql & ","
         End If
      Next
      Sql = Sql & ")"
      Sql = Sql & ")"
      
      Call xDb.Execute(Sql, dbFailOnError)
      Call DBEngine.Idle(dbFreeLocks)
      
      If i = Min Then
         Sql = "Select Min(IDCARTAO)  "
         Sql = Sql & " from CARTOES "
         Sql = Sql & " Where IDCARTAO > " & CStr(Min)
         Set Rs = xDb.OpenRecordset(Sql)
         Min = Val(Rs(0) & "")
         i = Max
      Else
         Sql = "Select Max(IDCARTAO)  "
         Sql = Sql & " from CARTOES "
         Sql = Sql & " Where IDCARTAO < " & CStr(Max)
         Set Rs = xDb.OpenRecordset(Sql)
         Max = Val(Rs(0) & "")
         i = Min
      End If
      DBEngine.Idle dbRefreshCache
      DBEngine.Idle dbFreeLocks
      Rs.Close
      Set Rs = Nothing
      
      Me.LblContador.Caption = CStr(Min) & "/" & CStr(Max)
      Me.LblContador.Refresh
      If Min = 0 And Max = 0 Then Max = -1
   Wend
     
   Sql = "Select C.IDCARTAO, N.NUMERO "
   Sql = Sql & " from NUMEROS N , CARTOES C"
   Sql = Sql & " Where N.IDCARTAO = C.IDCARTAO"
   Sql = Sql & " ORDER BY C.IDCARTAO, N.NUMERO"
   
   Set Rs = xDb.OpenRecordset(Sql)
   DBEngine.Idle dbRefreshCache
   DBEngine.Idle dbFreeLocks
   
   Me.GrdCartao.Rows = Rs.RecordCount / Val(Me.CmbSorteio)
   Me.GrdCartao.Cols = Val(Me.CmbSorteio) + 1
   Me.GrdCartao.FixedCols = 1
   For i = 0 To Me.GrdCartao.Cols - 1
      Me.GrdCartao.ColWidth(i) = 800
   Next
   
   For i = 0 To Me.GrdCartao.Rows - 1
      j = 0
      
      IDAnterior = Rs("IDCARTAO") & ""
      Me.GrdCartao.Redraw = False
      Me.GrdCartao.TextMatrix(i, 0) = IDAnterior
      Do While IDAnterior = Rs("IDCARTAO") & ""
         j = j + 1
         
         nAux = Val(Rs("NUMERO") & "")
         Contador = 1
         While nAux <> 1
            Contador = Contador + 1
            nAux = nAux / 2
         Wend
         Contador = Contador - 1
         Me.GrdCartao.TextMatrix(i, j) = Me.LstNumeros.List(Contador)

         IDAnterior = Rs("IDCARTAO") & ""
         Rs.MoveNext
         If Rs.EOF Then Exit Do
      Loop
      Me.GrdCartao.Redraw = True
      If Not Rs.EOF Then
         sAux = Rs("IDCARTAO") & ""
      End If
      Me.LblContador.Caption = CStr(i + 1) & "/" & CStr(Me.GrdCartao.Rows)
      Me.LblContador.Refresh
   Next
   Me.LblGrd.Caption = CStr(Me.GrdCartao.Rows) & " Cartões"
   Me.LblGrd.Refresh
   Screen.MousePointer = vbDefault
End Sub
Private Sub TxtNum_GotFocus()
   Me.TxtNum.SelStart = 0
   Me.TxtNum.SelLength = Len(Me.TxtNum)
End Sub

Private Sub TxtNum_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      Me.LstNumeros.ListIndex = Me.LstNumeros.ListIndex + 1
      Call PosTxtNum
   ElseIf KeyCode = vbKeyUp Then
      Me.LstNumeros.ListIndex = Me.LstNumeros.ListIndex - 1
      Call PosTxtNum
   ElseIf KeyCode = vbKeyPageDown Then
      Call PosTxtNum
   ElseIf KeyCode = vbKeyPageUp Then
      Call PosTxtNum
   ElseIf KeyCode = vbKeyEnd Then
      Me.LstNumeros.ListIndex = Me.LstNumeros.ListCount - 1
      Call PosTxtNum
   ElseIf KeyCode = vbKeyHome Then
      Me.LstNumeros.ListIndex = 0
      Call PosTxtNum
   End If
End Sub
