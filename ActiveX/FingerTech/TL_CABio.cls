VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CABio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCABio
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys      As Object
Private mvarClFinger As Object
Private mvarIDFUNC   As Integer
Public Sub Show()
   Set mvarClFinger = CriarObjeto("FingerTech.NG_Bio")
   With mvarClFinger
      '.Biometria = (xVal(Sys.GetParam(pCODPARAM:="BIOMETRIA", pCODSIS:=gCODSIS)) = 1)
      .Biometria = True
      If .Biometria Then
         Call .Init_Finger
         .fdbPath = ResolvePathName(mvarSys.EXEPATH) & "Digitais\"
         .FDbName = mvarSys.CODSIS & ".fdb"
          
          mvarMe.Show
          'mvarIDFUNC = .IDFUNC
      End If
   End With
End Sub
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Set ClFinger(ByVal vData As Object)
   Set mvarClFinger = vData
End Property
Public Property Get ClFinger() As Object
    Set ClFinger = mvarClFinger
End Property
Public Property Get IDFUNC() As Integer
    IDFUNC = mvarIDFUNC
End Property

Private Sub Class_Initialize()
   Set mvarMe = New FrmCABio
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   mvarMe.Hide
End Sub
Private Sub mvarMe_ChkServicoClick()
   Call HabilitarServico
End Sub
Private Sub mvarMe_CmdApagarClick()
   Dim nIDUSER As Long
   
   If nIDUSER = 0 Then
      nIDUSER = xVal(InputBox("Entre com número do usuário da Biometria.", "Biometria", 0))
   End If
   Call mvarClFinger.RemoveData(nIDUSER)
End Sub

Private Sub mvarMe_CmdIdentClick()
   Call F_Identificar
End Sub
Private Sub mvarMe_CmdLoadDBClick()
   Dim sFileName As String
   Dim sPath As String
   Dim sFile As String
   
   If mvarClFinger.Carregado Then
      If vbYes = ExibirPergunta("Exsite um Banco em memória. Deseja descarregá-lo?", "Biometria") Then
         '* Descarregar Banco
         Call mvarClFinger.UnloadFdb
      Else
         Exit Sub
      End If
   End If
         
   mvarMe.CommonDialog.CancelError = True
   mvarMe.CommonDialog.Filter = "FDB Files (*.fdb)|*.fdb"
   On Error Resume Next
   mvarMe.CommonDialog.ShowOpen
   If Err.Number = cdlCancel Then
      On Error GoTo 0
      Exit Sub
   End If
    
   sFileName = ""
   sFileName = mvarMe.CommonDialog.FileName
   If sFileName <> "" Then
      sFile = GetNameFromPath(sFileName, sPath)
      '* Carregar Banco
      If mvarClFinger.LoadFDb(sPath, sFile) Then
         Call PopulaTela
         'Call ExibirInformacao("Banco(.fdb) carregado com sucesso!", "Biometria")
      End If
    Else
        Call ExibirStop("Banco(.fdb) não carregado.", "Biometria")
    End If
    Call HabilitarServico
End Sub
Private Sub mvarMe_CmdRegClick()
   Dim nIDUSER As Long
   If nIDUSER = 0 Then
      nIDUSER = xVal(InputBox("Entre com número do usuário da Biometria.", "Biometria", 1))
   End If
   Call mvarClFinger.RegisterFinger(nIDUSER)
End Sub

Private Sub mvarMe_CmdSaveDBClick()
   Call mvarClFinger.SaveFDb
   Call HabilitarServico
End Sub
Private Sub mvarMe_CmdUnloadDbClick()
   Call mvarClFinger.UnloadFdb
   Call HabilitarServico
End Sub
Private Sub mvarMe_Load()
   Screen.MousePointer = vbHourglass
   Call mvarClFinger.LoadFDb
   Call HabilitarServico
   mvarMe.Move 0, 0, 0, 0
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Timer1Timer()
   '* Consulta o valor da tecla por Api e se a resposta for -32767
   '* , então mostra tecla em Chr(i) e o valor de i
   Dim i As Integer
   For i = 0 To 255
      If GetAsyncKeyState(i) = -32767 Then
         If i = 123 Then
            Call F_Identificar
         'ElseIf i = 122 Then
         '   End
         End If
      End If
   Next
End Sub
Public Sub HabilitarServico()
   With mvarClFinger
      If mvarMe.ChkServico.Visible Then
         .Biometria = (mvarMe.ChkServico.Value = vbChecked)
      Else
         mvarMe.ChkServico.Value = IIf(.Biometria, vbChecked, vbUnchecked)
      End If
   
      mvarMe.Timer1.Enabled = .Biometria
      mvarMe.CmdLoadDB.Enabled = .Biometria And Not .Carregado
      mvarMe.CmdUnloadDb.Enabled = .Biometria And .Carregado
      mvarMe.CmdSaveDB.Enabled = .Biometria And .Carregado
      
      mvarMe.CmdIdent.Enabled = .Biometria And .Carregado
      mvarMe.CmdReg.Enabled = .Biometria And .Carregado
      mvarMe.CmdApagar.Enabled = .Biometria And .Carregado
      
      mvarMe.TxtStrConect.Enabled = .Biometria And Sys.xDb.Conectado
      mvarMe.Frame1.Enabled = .Biometria
   End With
End Sub
Public Function F_Identificar() As Integer
   Dim nIDFDB     As Long
   Dim Sql        As String
   Dim MyRs       As Object
   
   Dim nNMFUNC    As String
   Dim nIDLOJA    As Long
   Dim sNMLOJA    As String
   Dim nIDBATIDA  As Long
   Dim sDTBATIDA  As String
   Dim nSENTIDO   As Long

   '*********
   '* Testa se Drive está instalado.
   On Error Resume Next
   Dim objNBioBSP  As Object
   If mvarSys.Propriedades("BIOMETRIA") = 1 Then
      Set objNBioBSP = CriarObjeto("NBioBSPCOM.NBioBSP", False)
      If objNBioBSP Is Nothing Then
         Call ExibirAviso("Sistema não identificou driver instalado." & vbNewLine & "A biometria será desligada.", "[CABio]")
         Call mvarSys.SaveParam(pCODPARAM:="BIOMETRIA", pVLPARAM:=0, pCODSIS:=mvarSys.CODSIS)
         mvarSys.Propriedades("BIOMETRIA") = 0
      Else
         objNBioBSP.Close
      End If
      Set objNBioBSP = Nothing
   End If
   
   On Error GoTo TrataErro
   If mvarSys.Propriedades("BIOMETRIA") = 0 Then
      F_Identificar = -1
      Exit Function
   End If
   If mvarClFinger Is Nothing Then
      Set mvarClFinger = CriarObjeto("FingerTech.NG_Bio")
      mvarClFinger.fdbPath = ResolvePathName(mvarSys.EXEPATH) & "Digitais\"
      mvarClFinger.FDbName = mvarSys.CODSIS & ".fdb"
      mvarClFinger.Biometria = (mvarSys.Propriedades("Biometria") = 1)
   End If
   With mvarClFinger
      mvarIDFUNC = -1
      If ExisteArquivo(.fdbPath & .FDbName) Then
         mvarIDFUNC = 0
         Call .Init_Finger
         Call .LoadFDb
         nIDFDB = .Identify_Finger
      End If
   End With
'   nIDFDB = 14
   If nIDFDB <> 0 Then
      Sql = "Select L.IDLOJA, L.NOME [NMLOJA], F.IDFUNCIONARIO, F.NOME " & vbNewLine
      Sql = Sql & " From " & gTBNAME & " F" & vbNewLine
      Sql = Sql & " Join OLOJA L On L.IDLOJA=F.IDLOJA" & vbNewLine
      Sql = Sql & " Where F.IDLOJA= " & SqlNum(Sys.IDLOJA) & vbNewLine
      Sql = Sql & " And F.IDFINGER = " & SqlNum(nIDFDB) & vbNewLine
      
      If Sys.xDb.AbreTabela(Sql, MyRs) Then
         nIDLOJA = MyRs("IDLOJA") & ""
         sNMLOJA = MyRs("NMLOJA") & ""
         mvarIDFUNC = MyRs("IDFUNCIONARIO") & ""
         nNMFUNC = MyRs("NOME") & ""
         sDTBATIDA = Sys.xDb.SysDate(3)
         
         Sql = "Select IDBATIDA, SENTIDO" & vbNewLine
         Sql = Sql & " From RBATIDA" & vbNewLine
         Sql = Sql & " Where IDLOJA= " & SqlNum(nIDLOJA) & vbNewLine
         Sql = Sql & " And IDFUNCIONARIO = " & SqlNum(mvarIDFUNC) & vbNewLine
         Sql = Sql & " And Day(DTBATIDA) = " & Day(sDTBATIDA) & vbNewLine
         Sql = Sql & " And Month(DTBATIDA)= " & Month(sDTBATIDA) & vbNewLine
         Sql = Sql & " And Year(DTBATIDA)= " & Year(sDTBATIDA) & vbNewLine
         Sql = Sql & " Order DTBATIDA"
         If Sys.xDb.AbreTabela(Sql, MyRs) Then
            If CDate(Format(sDTBATIDA, "hh:mm")) < CDate("12:00") Then
               nSENTIDO = 0
            Else
               nSENTIDO = 1
            End If
            
'            MyRs.MoveLast
'            If xVal(MyRs("SENTIDO") & "") = 0 Then
'               nSENTIDO = 1
'            Else
'               nSENTIDO = 0
'            End If
         Else
            nSENTIDO = 0
         End If
         
         
         Sql = "Insert Into RBATIDA "
         Sql = Sql & " (IDLOJA, IDFUNCIONARIO, IDBATIDA, DTBATIDA, SENTIDO) "
         Sql = Sql & " Select"
         Sql = Sql & " " & SqlNum(nIDLOJA)
         Sql = Sql & "," & SqlNum(mvarIDFUNC)
         Sql = Sql & ",(Select isNull(Max(IDBATIDA),0)+1 From RBATIDA Where IDLOJA=" & SqlNum(nIDLOJA) & " And IDFUNCIONARIO=" & SqlNum(mvarIDFUNC) & ")"
         Sql = Sql & "," & SqlDate(sDTBATIDA)
         Sql = Sql & "," & SqlNum(nSENTIDO)
         If Sys.xDb.Executa(Sql) Then
            Dim TlIdentidade As TL_Identidade
            Set TlIdentidade = New TL_Identidade
            With TlIdentidade
               Set .Sys = mvarSys
            
               .IDLOJA = Trim(CStr(nIDLOJA))
               .NMLOJA = sNMLOJA
               .IDFUNC = Trim(CStr(mvarIDFUNC))
               .NMFUNC = nNMFUNC
               .DTHORA = sDTBATIDA
               Set .ClFinger = mvarClFinger
               .Show
            End With
            Set TlIdentidade = Nothing
         Else
            Call ExibirInformacao("Erro ao gravar registro.")
         End If
      Else
         Call ExibirInformacao("Registro não encontrato no Banco de Dados.")
      End If
   End If
TrataErro:
   F_Identificar = mvarIDFUNC
   If Err <> 0 Then
      MsgBox Err.Number & " - " & Err.Description
   End If
End Function
Private Sub PopulaTela()
    Dim val
    Dim sFile As String
    Dim sFileName
'    Dim szTemp As String
'    Dim szFileName As String
'    Dim nUserID, nFingerID, nSampleNumber As Long
   
   '* converter arquivo para Txt
   sFile = mvarClFinger.FDbName
   sFileName = mvarClFinger.fdbPath & Left(sFile, Len(sFile) - 3) & "fid"
   If ExisteArquivo(sFileName) Then
'      Call ListSearchDB.ListItems.Clear
'
'      Open sFileName For Input As #1
'
'      Do While Not EOF(1)
'         Line Input #1, szTemp
'
'         If szTemp <> "" Then
'             val = Split(szTemp, vbTab)
'
'             Set ListItem = ListSearchDB.ListItems.Add
'             ListItem.Text = val(0)
'             ListItem.SubItems(1) = val(1)
'             ListItem.SubItems(2) = val(2)
'             Set ListItem = Nothing
'         End If
'      Loop
'
'      Close #1
   Else
      Call ExibirAviso("Arquivo não existe!", "Biometria")
   End If
End Sub

