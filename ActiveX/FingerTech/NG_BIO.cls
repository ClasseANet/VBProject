VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_BIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' Declaration global variables
Public mvarBiometria       As Boolean
Public mvarFDbPath         As String
Public mvarFDbName         As String
Public mvarCarregado       As Boolean

Public objNBioBSP       As Object   'NBioBSPCOMLib.NBioBSP
Public objDevice        As Object   'IDevice                ' Device object
Public objExtraction    As Object   'IExtraction            ' Extraction object
Public objNSearch       As Object   'INSearch               ' NSearch object
Public szTextEncodeFIR  As String
Public Property Let Carregado(ByVal vData As Boolean)
    mvarCarregado = vData
End Property
Public Property Get Carregado() As Boolean
    Carregado = mvarCarregado
End Property
Public Property Let FDbName(ByVal vData As String)
    mvarFDbName = vData
End Property
Public Property Get FDbName() As String
    FDbName = mvarFDbName
End Property
Public Property Let FDbPath(ByVal vData As String)
    mvarFDbPath = ResolvePathName(vData)
End Property
Public Property Get FDbPath() As String
    FDbPath = mvarFDbPath
End Property
Public Property Let Biometria(ByVal vData As Boolean)
    mvarBiometria = vData
End Property
Public Property Get Biometria() As Boolean
    Biometria = mvarBiometria
End Property
Public Sub Init_Finger()
   If Not mvarBiometria Then Exit Sub
   On Error Resume Next
   
   'Set objNBioBSP = New NBioBSPCOMLib.NBioBSP
   Set objNBioBSP = CriarObjeto("NBioBSPCOM.NBioBSP")
   If objNBioBSP Is Nothing Then Exit Sub
   Set objDevice = objNBioBSP.Device
   Set objExtraction = objNBioBSP.Extraction        ' Extraction object
   Set objNSearch = objNBioBSP.NSearch
   
   ' Check initialize object
   If objNSearch.ErrorCode <> NBioAPIERROR_NONE Then
       MsgBox objNSearch.ErrorDescription & " [" & objNSearch.ErrorCode & "]", vbOKOnly
   End If
End Sub
Public Sub Terminate_Finger()
   If Not mvarBiometria Then Exit Sub
   On Error Resume Next
   Set objDevice = Nothing
   Set objExtraction = Nothing
   Set objNSearch = Nothing
   Set objNBioBSP = Nothing
End Sub
Public Function Identify_Finger(Optional pTentativas As Integer = 3) As Long
   If Not mvarBiometria Then Exit Function
       
   Dim i As Integer
   Dim szTextEncodeFIR As String
   Dim nUserID            As Long
   Dim nTent As Integer

   On Error Resume Next
        
   If objNBioBSP Is Nothing Then
      Call Init_Finger
   End If
   nUserID = 0
   szTextEncodeFIR = ""
'    Call ListResult.ListItems.Clear
   
   ' Get FIR data
   Call objDevice.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
   Call objExtraction.Capture(NBioAPI_FIR_PURPOSE_VERIFY + NBioAPI_WINDOW_STYLE_TOPMOST)
   If objExtraction.ErrorCode <> NBioAPIERROR_NONE Then
      If objExtraction.ErrorCode = 513 Then     '* Cancelado
      ElseIf objExtraction.ErrorCode = 515 Then '* Timeout
      Else
        MsgBox objExtraction.ErrorDescription & " [" & objExtraction.ErrorCode & "]"
        Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
      End If
      Exit Function
    End If

    Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)

    szTextEncodeFIR = objExtraction.TextEncodeFIR

    ' Identify FIR to NSearch DB
    Call objNSearch.IdentifyUser(szTextEncodeFIR, NBioAPI_FIR_SECURITY_LEVEL_NORMAL)

    If objNSearch.ErrorCode <> 0 Then
      If objNSearch.ErrorCode = 777 Then
         Call ExibirStop("Digital não identificada." & vbNewLine & vbNewLine & "Tente novamente.")
         nTent = nTent + 1
         While nTent < pTentativas
            Identify_Finger = Identify_Finger(1)
         Wend
      Else
         Call ExibirStop("Erro na identificalção biométrica." & vbNewLine & " [" & objNSearch.ErrorCode & "]- " & objNSearch.ErrorDescription)
      End If
      If Identify_Finger = 0 Then
         Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
      End If
      Exit Function
   Else
      '* Add item to list of result
      'MsgBox "IDFINGER" & objNSearch.UserID
      Identify_Finger = objNSearch.UserID
   End If
End Function
'Private Sub Register_Finger()
'
'    Dim i, j As Integer
'    Dim nUserID As Long
'    Dim szFir As String
'    Dim ListItem As ListItem
'
'    Dim objResult As ICandidateList         ' CandidateList or Result object
'
'    nUserID = 0
'    szFir = ""
'
'    ' Get User ID
'    If Not IsNumeric(txtUserID.Text) Then
'        MsgBox "User ID must be have numeric type and greater than 0.", vbOKOnly, "Error"
'        Exit Sub
'    End If
'
'    nUserID = CLng(txtUserID.Text)
'
'    ' Get FIR data
'    Call objDevice.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
'    Call objExtraction.Enroll(Null)
'    If objExtraction.ErrorCode <> NBioAPIERROR_NONE Then
'        MsgBox objExtraction.ErrorDescription & " [" & objExtraction.ErrorCode & "]"
'        Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
'        Exit Sub
'    End If
'
'    Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
'
'    szFir = objExtraction.TextEncodeFIR
'
'    ' Regist FIR to NSearch DB
'    Call objNSearch.AddFIR(szFir, nUserID)
'    If objNSearch.ErrorCode <> NBioAPIERROR_NONE Then
'        MsgBox objNSearch.ErrorDescription & " [" & objNSearch.ErrorCode & "]"
'        Exit Sub
'    End If
'
'
'    ' Add item to list of SearchDB
'    For Each objResult In objNSearch
'
'        Set ListItem = ListSearchDB.ListItems.Add
'        ListItem.Text = objResult.UserID
'        ListItem.SubItems(1) = objResult.FingerID
'        ListItem.SubItems(2) = objResult.SampleNumber
'        Set ListItem = Nothing
'
'    Next
'
'    txtUserID.Text = CLng(txtUserID.Text) + 1
'    ListSearchDB.SelectedItem.Selected = False
'
'End Sub
Private Sub Class_Initialize()
   Call Init_Finger
End Sub
Public Function RegisterFinger(pIDUSER As Long) As Boolean
   'Dim nFingerID, nSampleNumber As Long
    Dim i, j As Integer
    Dim nUserID As Long
    Dim szFir As String
    
   If pIDUSER = 0 Then
      Call ExibirStop("Usuário de Biometria não identificado.", "Biometria")
      Exit Function
   End If
   If Not mvarCarregado Then
      If LoadFDb Then
         If Not mvarCarregado Then
            Call ExibirStop("Banco não carregado", "Biometria")
            Exit Function
         End If
      Else
         Exit Function
      End If
   End If
   
   '* Recuperar Informação FIR
   Call objDevice.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
   Call objExtraction.Enroll(Null)
   If objExtraction.ErrorCode <> NBioAPIERROR_NONE Then
      Call ExibirStop("Erro ao registrar impressão." & vbNewLine & " [" & objExtraction.ErrorCode & "]- " & objExtraction.ErrorDescription)
      Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
      Exit Function
   End If
   
   Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
   szFir = objExtraction.TextEncodeFIR
    
   '* Registra Informação FIR no Banco
   Call objNSearch.AddFIR(szFir, pIDUSER)
   If objNSearch.ErrorCode <> NBioAPIERROR_NONE Then
      If objNSearch.ErrorCode = 781 Then
         Call RemoveData(pIDUSER)
         Call objNSearch.AddFIR(szFir, pIDUSER)
      End If
      If objNSearch.ErrorCode <> NBioAPIERROR_NONE Then
         Call ExibirStop("Erro ao gravar informação no Banco(.fdb)." & vbNewLine & " [" & objNSearch.ErrorCode & "]- " & objNSearch.ErrorDescription)
         Exit Function
      End If
   End If
   If SaveFdb Then
      RegisterFinger = True
   End If
End Function
Public Function SaveFdb(Optional pPath As String, Optional pFile As String) As Boolean
   Dim szFileName As String

   If pPath = "" Then pPath = mvarFDbPath
   If pFile = "" Then pFile = mvarFDbName
   
   szFileName = mvarFDbPath & mvarFDbName
   If InStr(szFileName, ".") = 0 Then szFileName = szFileName & ".fdb"

   If Len(szFileName) > "4" Then
      ' Salvar Arquivo FDB
      Call objNSearch.SaveDBToFile(szFileName)
      
      SaveFdb = ExisteArquivo(szFileName)
   End If
End Function
Public Function RemoveData(pUserID As Long, Optional pFingerID As Integer = -1, Optional pSampleNumber As Integer = -1) As Boolean
   Dim i As Integer

   If pFingerID = -1 Then
      For i = 1 To 10
         Call RemoveData(pUserID, i, pSampleNumber)
      Next
   End If
   If pSampleNumber = -1 And pFingerID >= 0 Then
      For i = 0 To 1
         Call RemoveData(pUserID, pFingerID, i)
      Next
   End If
   If pFingerID <> -1 And pSampleNumber <> -1 Then
      Call objNSearch.RemoveData(pUserID, pFingerID, pSampleNumber)
      If objNSearch.ErrorCode = 0 Then
         RemoveData = True
      Else
         'Call ExibirStop("Erro ao Remover" & vbNewLine & " [" & objNSearch.ErrorCode & "]- " & objNSearch.ErrorDescription, "Biometria")
      End If
   End If
End Function
Public Sub UnloadFDb()
   On Error Resume Next
   Call objNSearch.ClearDB
   mvarCarregado = False
End Sub
Public Function LoadFDb(Optional pPath As String, Optional pFile As String) As Boolean
   Dim szFileName As String

   If pPath = "" Then pPath = mvarFDbPath
   If pFile = "" Then pFile = mvarFDbName
      
   '* Descarregar Banco
   Call UnloadFDb
       
   If Dir(pPath & pFile) = "" Then
      Call SaveFdb(pPath, pFile)
   End If
       
   szFileName = pPath & pFile
   If szFileName = "" Then
      Call ExibirStop("Arquivo inválido ou inexistente!" & vbNewLine & szFileName)
   Else
      'Carregar Banco do Arquivo
      Call objNSearch.LoadDBFromFile(szFileName)
      If objNSearch.ErrorCode = 0 Then
         mvarFDbPath = pPath
         mvarFDbName = pFile
         
         LoadFDb = True
         mvarCarregado = True
      Else
         Call ExibirStop("Erro ao carregar arquivo." & vbNewLine & " [" & objNSearch.ErrorCode & "]" & objNSearch.ErrorDescription)
      End If
   End If
    
End Function
Private Sub Class_Terminate()
   Call Terminate_Finger
End Sub
