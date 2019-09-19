Attribute VB_Name = "BlCAVersao"
Public Sub Main()
   Dim MyVerif As Object
   
   If InStr(App.Path, "\Sistemas\ActiveX\CAVersao") <> 0 Then
      Set MyVerif = CriarObjeto("VersaoFTP.TL_VerifVersao")
      MyVerif.CODSIS = "P3R"
      MyVerif.FileList = App.Path & "\" & "Files.txt"
      MyVerif.ShowFTP
      Set MyVerif = Nothing
   Else
      Set MyVerif = CriarObjeto("VersaoFTP.TL_VerifVersao")
      'If ExibirPergunta("Exibir FTP?", "CAVs") = vbYes Then
      '   MyVerif.ShowFTP
      'Else
         MyVerif.ShowCAVs
      'End If
      Set MyVerif = Nothing
      End
   End If
End Sub
