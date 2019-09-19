Attribute VB_Name = "BLVersaoFTP"
Public Sub Zip2()
   Dim oZip As CGZipFiles
   On Error GoTo vbErrorHandler
    
    Set oZip = New CGZipFiles
    
    With oZip
'
' Give Zip File a Name / Path
'
        .ZipFileName = "\ZIPTEST.ZIP"
'
' Are we updating a Zip File ?
' - This doesn't seem to work - check InfoZip
' homepage for more info.
'
        .UpdatingZip = False ' ensures a new zip is created
'
' Add in the files to the zip - in this case, we
' want all the ones in the current directory
'
        .AddFile App.Path & "\*.*"
'
' Make the zip file & display any errors
'
        If .MakeZipFile <> 0 Then
            MsgBox .GetLastMessage ' any errors
        End If
    End With
    
    Set oZip = Nothing
    
    MsgBox "\ZIPTEST.ZIP Created Successfully"
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1::cmdZip_Click" & " " & Err.Description
End Sub

