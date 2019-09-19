VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADRBATIDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents mvarMe  As FrmCADRBATIDA
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   Dim bBiometria As Boolean
   'Dim sMsg       As String
   
   bBiometria = (xVal(mvarSys.Propriedades("BIOMETRIA")) = 1)
   '* Verifica Biometria
   If bBiometria Then
      bBiometria = ExisteArquivo(Environ("PROGRAMFILES") & "\NITGEN eNBSP\SDK\Bin\NBioBSPCOM.dll") And Not ExisteArquivo(Environ("WINDIR") & "\System32\NBioBSPCOM.dll")
      'If Not bBiometria Then sMsg = "Sistema não encontrou o arquivo '" & Environ("PROGRAMFILES") & "\NITGEN eNBSP\SDK\Bin\NBioBSPCOM.dll" & "'" & vbNewLine & " A biometria será desligada."
      mvarSys.Propriedades("BIOMETRIA") = IIf(bBiometria, 1, 0)
      Call Sys.SaveParam(pCODPARAM:="BIOMETRIA", pVLPARAM:=CStr(IIf(bBiometria, 1, 0)))
   End If
   If bBiometria Then
      Dim oCaBio     As Object
      
      Set oCaBio = CriarObjeto("CABio.TL_CABio")
      Set oCaBio.Sys = mvarSys
      If oCaBio.F_Identificar < 0 Then
         mvarMe.Show vbModal
      End If
      Set oCaBio = Nothing
   Else
      mvarMe.Show vbModal
   End If
   
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADRBATIDA
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
   End If
   mvarMe.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmbSENTIDOKeyPress(KeyAscii As Integer)
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_CmbUNIDADEKeyPress(KeyAscii As Integer)
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   Call Salvar
End Sub
Private Sub mvarMe_Load()
   Call MontarTela
   Call LimpaTela
End Sub
Private Sub mvarMe_Timer()
   mvarMe.TxtHORA.Text = mvarSys.xdb.SysDate(2)
End Sub
Private Sub mvarMe_TxtCHAPAGotFocus()
   'Call SelecionarTexto(Me.ActiveControl)
   Call LimpaTela
End Sub
Private Sub LimpaTela()
   Dim nIDLOJA As Integer

   Call LimparTela(mvarMe)
   Call LocalizarCombo(mvarMe.CmbUnidade, mvarSys.IDLOJA, True, True)
   
   mvarMe.CmdSalvar.Enabled = True
   mvarMe.TxtSenha.Tag = ""
   mvarMe.Timer1.Enabled = True
   mvarMe.PctReg.Visible = False
   mvarMe.TxtHORA.ForeColor = vbBlack
   mvarMe.TxtHORA.BackColor = vbWhite
   
   With mvarMe.TxtData
      .Text = Format(mvarSys.xdb.SysDate, "dd/mm/yy")
      Select Case Weekday(.Text)
         Case 1: .Text = .Text & ", Domingo"
         Case 2: .Text = .Text & ", Segunda-Feira"
         Case 3: .Text = .Text & ", Terça-Feira"
         Case 4: .Text = .Text & ", Quarta-Feira"
         Case 5: .Text = .Text & ", Quinta-Feira"
         Case 6: .Text = .Text & ", Sexta-Feira"
         Case 7: .Text = .Text & ", Sábado-Feira"
      End Select
   End With
   mvarMe.TxtHORA.Text = mvarSys.xdb.SysDate(2)
End Sub
Private Sub mvarMe_TxtCHAPAKeyPress(KeyAscii As Integer)
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_TxtCHAPALostFocus()
   Dim Sql As String
   Dim Rs  As Object
   Dim sAux As String
   
   If mvarMe.ActiveControl Is mvarMe.CmdSair Then Exit Sub
   
   If Len(Trim(mvarMe.TxtCHAPA)) < mvarMe.TxtCHAPA.MaxLength Then
      If IsNumeric(Mid(mvarMe.TxtCHAPA, 1, 1)) Then
         mvarMe.TxtCHAPA.Text = StrZero(mvarMe.TxtCHAPA.Text, mvarMe.TxtCHAPA.MaxLength)
      Else
         mvarMe.TxtCHAPA.Text = Mid(mvarMe.TxtCHAPA, 1, 1) & StrZero(Mid(mvarMe.TxtCHAPA, 2), mvarMe.TxtCHAPA.MaxLength - 1)
      End If
   End If
   
   Sql = "Select *"
   Sql = Sql & " From RFUNCIONARIO"
   Sql = Sql & " Where IDLOJA=" & mvarMe.CmbUnidade.ItemData(mvarMe.CmbUnidade.ListIndex)
   Sql = Sql & " And CHAPA=" & SqlStr(mvarMe.TxtCHAPA.Text)
   If mvarSys.xdb.AbreTabela(Sql, Rs) Then
      mvarMe.TxtID.Text = Rs("IDFUNCIONARIO")
      mvarMe.TxtNOME = Rs("NOME")
      mvarMe.TxtSenha.Tag = Decrypt2(Rs("SENHA"))
      If Decrypt2(Rs("SENHA")) = "" Then
         sAux = "Funcionário: " & Rs("CHAPA") & " - " & Rs("NOME") & vbNewLine & vbNewLine
         sAux = sAux & "Senha não cadastrada." & vbNewLine
         sAux = sAux & "O Funcionário deve cadastrar sua senha para registrar sua batida."
         Call ExibirAviso(sAux & vbNewLine, "Senha")
         mvarMe.TxtCHAPA.SetFocus
      Else
         Sql = "Select *"
         Sql = Sql & " From RBATIDA"
         Sql = Sql & " Where IDLOJA=" & mvarMe.CmbUnidade.ItemData(mvarMe.CmbUnidade.ListIndex)
         Sql = Sql & " And IDFUNCIONARIO=" & SqlStr(Rs("IDFUNCIONARIO"))
         Sql = Sql & " And YEAR(DTBATIDA)=" & SqlStr(Year(Mid(mvarMe.TxtData.Text, 1, 8)))
         Sql = Sql & " And MONTH(DTBATIDA)=" & SqlStr(Month(Mid(mvarMe.TxtData.Text, 1, 8)))
         Sql = Sql & " And DAY(DTBATIDA)=" & SqlStr(Day(Mid(mvarMe.TxtData.Text, 1, 8)))
         If mvarSys.xdb.AbreTabela(Sql, Rs) Then
            If xVal(Rs("SENTIDO") & "") = 0 Then
               mvarMe.CmbSENTIDO.ListIndex = 1
            Else
               mvarMe.CmbSENTIDO.ListIndex = 0
            End If
         Else
            mvarMe.CmbSENTIDO.ListIndex = 0
         End If
         mvarMe.TxtSenha.SetFocus
      End If
   Else
      Call ExibirAviso("Registro Inválido")
      mvarMe.TxtCHAPA.SetFocus
   End If
   
End Sub
Private Sub mvarMe_TxtSENHAKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call mvarMe_CmdSalvarClick
   End If
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub MontarTela()
  Dim Sql As String
  Dim MyRs As Object
  With mvarMe
     .CmbSENTIDO.Clear
     .CmbUnidade.Clear
     
     .CmbSENTIDO.AddItem "Entrada"
     .CmbSENTIDO.ItemData(.CmbSENTIDO.NewIndex) = 0
     .CmbSENTIDO.AddItem "Saida"
     .CmbSENTIDO.ItemData(.CmbSENTIDO.NewIndex) = 1
     
      Sql = "Select IDLOJA, NOME "
      Sql = Sql & " From OLOJA"
      Sql = Sql & " Where IDCOLIGADA=" & mvarSys.IDCOLIGADA
      
      If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            .CmbUnidade.AddItem MyRs("NOME")
            .CmbUnidade.ItemData(.CmbUnidade.NewIndex) = MyRs("IDLOJA")
            MyRs.MoveNext
         Wend
         Call LocalizarCombo(.CmbUnidade, mvarSys.IDLOJA, True, True)
      End If
  End With
  Set MyRs = Nothing
End Sub
Private Function Salvar() As Boolean
   Dim MyBatida As Object
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   Dim Queries As Collection
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Set MyBatida = CriarObjeto("BANCO_3R.TB_RBATIDA")
      With MyBatida
         Set .xdb = mvarSys.xdb
         .IDLOJA = mvarSys.IDLOJA
         .IDFUNCIONARIO = mvarMe.TxtID.Text
         .DTBATIDA = Format(Mid(mvarMe.TxtData.Text, 1, 8), "DD/MM/YYYY") & " " & Format(mvarMe.TxtHORA.Text, "hh:mm:ss")
         .SENTIDO = mvarMe.CmbSENTIDO.ItemData(mvarMe.CmbSENTIDO.ListIndex)
         .FLGMANUAL = 1
      End With
      
      bOk = True
      bIsDirt = MyBatida.isDirt
      If bIsDirt Then
         bOk = MyBatida.Salvar()
      End If
      If bOk Then
         mvarMe.Timer1.Enabled = False
         mvarMe.PctReg.Visible = True
         mvarMe.TxtHORA.BackColor = &HC000&
         mvarMe.TxtHORA.ForeColor = vbWhite
      End If
      mvarMe.CmdSalvar.Enabled = Not bOk
   End If
   
   Salvar = bOk
End Function
Private Function ValidaCampos() As Boolean
   If mvarMe.TxtSenha.Tag = mvarMe.TxtSenha.Text Then
      ValidaCampos = True
   Else
      DoEvents
      Call ExibirInformacao("Senha inválida!")
      mvarMe.TxtSenha.SetFocus
   End If
End Function
