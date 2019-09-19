Attribute VB_Name = "Bl_Ponto"
Option Explicit
Global gCODSIS As String
Global gSetupFile As String
Global gLocalReg As String
Global Sys     As Object
Sub main()
   Dim MyPonto As Object
   
   
   Set Sys = GetSys("Projeto3R", "P3R", "2")
   
   Set MyPonto = New TL_Ponto
   With MyPonto
      Set .Sys = Sys
      .Show
   End With
   
   
End Sub
Private Function GetSys(Projeto As String, CodSis As String, Optional Conn As String = "0") As Object
   Dim sConn  As String
   Dim Splash As Object
   Dim MySys  As Object
    
   gSetupFile = "SETUP.INI"
   gCODSIS = CodSis
   gLocalReg = Environ("programfiles") & "\ClasseA\" & Projeto & "\" & gCODSIS & ".reg"
   'sConn = "Conection " & ReadIniFile(gLocalReg, "Conections", "Last", "0")
   sConn = "Conection " & Conn

   If Not ExisteArquivo(gLocalReg) Then Exit Function

   Set MySys = CriarObjeto("SysA.SetA")
   
   Set Splash = CriarObjeto("CONEXAO.Splash")
   With Splash
      Set .Sys = MySys
      .DebugSys = False
      .CodSis = gCODSIS
      .Alias = ReadIniFile(gLocalReg, sConn, "Alias", "")
      .DbTipo = ReadIniFile(gLocalReg, sConn, "dbTipo", "")
      .Server = ReadIniFile(gLocalReg, sConn, "Server", "")
      .dbName = ReadIniFile(gLocalReg, sConn, "dbName", "")
      .UID = ReadIniFile(gLocalReg, sConn, "UID", "")
      .PWD = Decrypt2(ReadIniFile(gLocalReg, sConn, "Pwd", ""))
      
      If Not .VerificaLicenca1 Then
         End
      End If
   End With
   
   With MySys
      Set .XDb = Splash.XDb
      .IDUSU = Trim(Splash.IDUSU)
      .CodSis = gCODSIS
      .LocalReg = gLocalReg
      If .IDUSU = "" Then End
      '.GetIniVars pCODSIS:=gCODSIS, pIniFile:=gSetupFile, pAppPath:=App.Path
   End With
   Set GetSys = MySys
   Set Splash = Nothing
End Function
