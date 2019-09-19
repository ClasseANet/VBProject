Attribute VB_Name = "P3R"
Option Explicit
Public gCODSIS As String
Public gDBTipo As Integer
Public gALIAS  As String
Public gSERVER As String
Public gDBNAME As String
Public gDBUSER As String
Public gDBPWD  As String

Global MdiMenu As Object

'Public Enum eIMGMenu
'
'End Enum

Public Function LoadVersion() As String
   Dim sVer    As String
   LoadVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function MyLoadPicture() As Object
   On Error GoTo TrataErro
   'Set MyLoadPicture = LoadResPicture("P3R", 1)
   If ExisteArquivo(App.Path & "\" & gCODSIS & ".ico") Then
      Set MyLoadPicture = LoadPicture(App.Path & "\" & gCODSIS & ".ico")
   End If
   
   
   '* Erro ao carrebar MDI
   'Set MyLoadPicture = MDI.imgList.ListImages(gCODSIS).Picture
   'Set MyLoadPicture = MDI.Picture
   
Exit Function
TrataErro:
   MsgBox Err & "-" & Error
End Function
Public Sub MyLoadgCODSIS(Optional bCODSIS As Boolean)
   Dim sPathSetup As String
   Dim sConn As String
      
   gDBTipo = -1
   gCODSIS = "P3R"
   If bCODSIS Then Exit Sub
   
   'Call WriteIniFile(gLocalReg, "Conections", "Last", "1")
   
   If gLocalReg <> "" Then
      sConn = "Conection " & ReadIniFile(gLocalReg, "Conections", "Last", "0")
      gALIAS = ReadIniFile(gLocalReg, sConn, "Alias", "")
      gDBTipo = ReadIniFile(gLocalReg, sConn, "dbTipo", "-1")
      gSERVER = ReadIniFile(gLocalReg, sConn, "Server", "")
      gDBNAME = ReadIniFile(gLocalReg, sConn, "dbName", "")
      gDBUSER = ReadIniFile(gLocalReg, sConn, "UID", "")
      gDBPWD = Decrypt2(ReadIniFile(gLocalReg, sConn, "Pwd", ""))
      
      sPathSetup = ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "")
      sPathSetup = ResolvePathName(sPathSetup)
   End If
   
   If sPathSetup <> "" Then
      sPathSetup = sPathSetup & gSetupFile
      If gALIAS = "" Then gSERVER = ReadIniFile(sPathSetup, "Database Format", "Alias", "")
      If gDBTipo = -1 Then gDBTipo = ReadIniFile(sPathSetup, "Database Format", "dbTipo", "-1")
      If gSERVER = "" Then gSERVER = ReadIniFile(sPathSetup, "Database Format", "Server", "")
      If gDBNAME = "" Then gDBNAME = ReadIniFile(sPathSetup, "Database Format", "dbName", "")
      If gDBUSER = "" Then
         gDBUSER = ReadIniFile(sPathSetup, "Database Format", "UID", "")
         If gDBPWD = "" Then gDBPWD = Decrypt2(ReadIniFile(sPathSetup, "Database Format", "Pwd", ""))
      End If
   End If
      
'   gIDUSU = "DIO"   '* Para não exibir Splash
   If gALIAS = "" Then gALIAS = "PRODUCAO"
   If gDBTipo = -1 Then gDBTipo = 1
   If gDBNAME = "" Then gDBNAME = "G3R"
   If gDBUSER = "" Then
      gDBUSER = "USU_VERIF"
      If gDBPWD = "" Then gDBPWD = Decrypt2("7B75787D7A776274616A7A")
   End If

   gCaption1 = "Projeto 3R"
   gCaption2 = "Módulo"
   gCaption3 = "Loja"
End Sub
Public Sub MyLimpaInstaciaObj()
   If gDebug Then MsgBox "LimpaInstaciaObj"
   On Error GoTo TrataErro
         
   Set MdiMenu = Nothing
         
   Exit Sub
TrataErro:
   MsgBox "Limpa Instacia de Objetos" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
End Sub

Public Sub MyInstaciaObj()
   Dim TpErro  As String
   Dim Erro429 As Boolean
   
   If gDebug Then MsgBox "MyInstaciaObj..."
   
   On Error GoTo TrataErro
   TpErro = "MdiMenu"
   '#If ComRef Then
   '   Set MdiMenu = New Menu.ControlMenu
   '#Else
      Set MdiMenu = CreateObject("Menu.ControlMenu")
   '#End If
   If gDebug Then MsgBox "Criou Menu"
   
   GoTo Saida
TrataErro:
   If Err = 429 Then
      If gDebug Then MsgBox "Não Criou Menu"
      Erro429 = True
      Resume Next
   Else
      MsgBox "Instacia de Objetos [" & TpErro & "]" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
   End If
Saida:
   If Erro429 Then
      Err = 429
      MsgBox "Instacia de Objetos [" & TpErro & "]" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
   End If
End Sub
