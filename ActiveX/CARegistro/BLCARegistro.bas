Attribute VB_Name = "BLCARegistro"
Public Sub Main()
   Dim MyReg      As Object
   Dim sTAG       As String
   Dim sCommand   As String
   
   sCommand = Command$
   'sCommand = "/RUNAS:|USERNAME=dsr|PASSWORD=dolphinmm|DOMAINNAME=MAUAJURONG|COMMANDLINE=C:\WINDOWS\system32\ClasseA\CAReg.exe|CURRENTDIR=C:\WINDOWS\system32\ClasseA\"
   'sCommand = "/RUNAS:|USERNAME=teste|PASSWORD=teste|DOMAINNAME=HAMYLTON|COMMANDLINE=C:\WINDOWS\system32\ClasseA\CARegistro.exe|CURRENTDIR=C:\WINDOWS\system32\ClasseA\"
   
   Set MyReg = CriarObjeto("CARegistro.RunAs")
   'Set MyReg = New RunAs
   With MyReg
      If InStr(sCommand, "RUNAS:") <> 0 Then
         sTAG = Mid(sCommand, InStr(sCommand, "RUNAS") + Len("RUNAS:"))
         .UserName = GetTag(sTAG, "USERNAME", "")
         .Password = GetTag(sTAG, "PASSWORD", "")
         .DomainName = GetTag(sTAG, "DOMAINNAME", "")
         .CommandLine = GetTag(sTAG, "COMMANDLINE", "")
         .CurrentDir = ResolvePathName(GetTag(sTAG, "CURRENTDIR", ""))
         .Command = sCommand
      End If
      .Show
   End With
   Set MyReg = Nothing
End Sub

