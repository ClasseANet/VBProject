Attribute VB_Name = "WizPrj"
Option Explicit
Global Const SecaoBase = "Conection "

Global mvarLocalReg
Global XDb As New DS_BANCO

Public Sub Main()
   On Error Resume Next
   FrmWizPrj.Show vbModal
End Sub
Public Sub ObjetoParaRegistro(MyXdb As Object, Index)
   Dim sSecao  As String
   
   sSecao = SecaoBase & Index
   
   Call WriteIniFile(mvarLocalReg, sSecao, "ALIAS", MyXdb.Alias)
   Call WriteIniFile(mvarLocalReg, sSecao, "isODBC", False)
   Call WriteIniFile(mvarLocalReg, sSecao, "DBTIPO", MyXdb.dbTipo)
   Call WriteIniFile(mvarLocalReg, sSecao, "isADO", True)
   If MyXdb.dbTipo = 0 Then
      Call WriteIniFile(mvarLocalReg, sSecao, "DBDRIVE", MyXdb.dbDrive)
   Else
      If UCase(Mid(MyXdb.Server, 1, Len(Environ("COMPUTERNAME")))) = UCase(Environ("COMPUTERNAME")) Then
         Call WriteIniFile(mvarLocalReg, sSecao, "SERVER", "[Local]" & Mid(MyXdb.Server, Len(Environ("COMPUTERNAME")) + 1))
      Else
         Call WriteIniFile(mvarLocalReg, sSecao, "SERVER", MyXdb.Server)
      End If
   End If
   Call WriteIniFile(mvarLocalReg, sSecao, "DBNAME", MyXdb.dbName)
   Call WriteIniFile(mvarLocalReg, sSecao, "UID", MyXdb.UID)
   Call WriteIniFile(mvarLocalReg, sSecao, "PWD", Encrypt2(MyXdb.PWD))
End Sub
Public Sub RegistroParaObjeto(ByRef MyXdb As Object, pAlias As String)
   Dim sSecao  As String
   
   sSecao = GetSecao(pAlias)
   If sSecao = "" Then Exit Sub
   With MyXdb
      .Alias = ReadIniFile(mvarLocalReg, sSecao, "ALIAS")
      .isODBC = ReadIniFile(mvarLocalReg, sSecao, "isODBC", False)
      .dbTipo = ReadIniFile(mvarLocalReg, sSecao, "DBTIPO", 1)
      .dbVersao = ReadIniFile(mvarLocalReg, sSecao, "DBVERSAO")
      .isADO = ReadIniFile(mvarLocalReg, sSecao, "isADO", True)
      .dbDrive = ""
      .Server = ""
      If MyXdb.dbTipo = 0 Then
         .dbDrive = ReadIniFile(mvarLocalReg, sSecao, "DBDRIVE")
      Else
         .Server = ReadIniFile(mvarLocalReg, sSecao, "SERVER")
      End If
      .dbName = ReadIniFile(mvarLocalReg, sSecao, "DBNAME")
      .UID = ReadIniFile(mvarLocalReg, sSecao, "UID")
      .PWD = Decrypt2(ReadIniFile(mvarLocalReg, sSecao, "PWD"))
   End With
End Sub
Public Sub RenumerarSetting()
   Dim i       As Integer
   Dim sAlias  As String
   Dim MyXdb   As Object
      
   For i = 5 To 1 Step -1
      Set MyXdb = Nothing
      Set MyXdb = CreateObject("XBANCO01.DS_BANCO")
      
      sAlias = ReadIniFile(mvarLocalReg, SecaoBase & i - 1, "Alias")
      If sAlias <> "" Then
         Call RegistroParaObjeto(MyXdb, sAlias)
         Call ObjetoParaRegistro(MyXdb, i)
      End If
   Next
   Set MyXdb = Nothing
End Sub
Public Function GetSecao(pAlias As String) As String
   Dim i       As Integer
   Dim j       As Integer
   Dim sAlias  As String
   Dim bAchou  As Boolean
   On Error GoTo TrataErro
   i = 0
   j = 0
   sAlias = ReadIniFile(mvarLocalReg, SecaoBase & i, "Alias")
   While UCase(pAlias) <> UCase(sAlias) And j <> 2
      If sAlias = "" Then j = j + 1
      i = i + 1
      sAlias = Trim(ReadIniFile(mvarLocalReg, SecaoBase & i, "Alias", ""))
      If sAlias <> "" Then j = 0
   Wend
   GetSecao = SecaoBase & i - j
Exit Function
TrataErro:
   ShowError "Splash.GetSecao"
End Function

