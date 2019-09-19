Attribute VB_Name = "Rel"
Sub Main()
   Dim sCommand   As String
   Dim sNmRel     As String
   
   Screen.MousePointer = vbHourglass
   sCommand = UCase(Trim(Command$))
   
   If Trim(sCommand) <> "" Then
      sNmRel = sCommand
      If Right(sCommand, 4) <> ".rpt" Then
         sNmRel = sCommand & ".rpt"
      End If
      Call ExtractResData(sCommand, "RPT", App.Path & "\" & sNmRel)
      Screen.MousePointer = vbDefault
      End
   End If
   FrmRel.Show vbModal
   Screen.MousePointer = vbDefault
End Sub
Public Sub ExtractResData(ID, Tipo, Arquivo As String)
   Dim nInt As Integer
   Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
   
   On Error GoTo Fim
   
   'Call ClsDos.Del(Arquivo)
   Call Kill(Arquivo)
   
   nInt = FreeFile
   Open Arquivo$ For Binary Access Write As nInt
      byteFileBuf = LoadResData(ID, Tipo)
      Put nInt, , byteFileBuf
   GoTo Saida
Fim:
   If Err = 53 Then '* File Not Found
     Resume Next
   Else
      Resume
   End If
Saida:
   Close nInt
   Err = 0
   Exit Sub
End Sub

