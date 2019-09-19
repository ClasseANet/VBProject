Attribute VB_Name = "DSR"
Option Explicit
Public Function FormateDate(pStrDate As String) As String
   Dim sText As String

   sText = pStrDate
   
   sText = Mid(Replace(sText, "/", ""), 1, 8)
   If Len(sText) <= 2 Then
      sText = StrZero(sText, 2) + Format(Now(), "/mm/yyyy")
   ElseIf Len(sText) <= 4 Then
      sText = Mid(sText, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Format(Now(), "/yyyy")
   ElseIf Len(sText) <= 8 Then
      If Mid(sText, 5, 4) >= Left(Year(Now), Len(Mid(sText, 5, 4))) Then
         sText = Mid(sText, 1, 2) + "/" + Mid(sText, 3, 2) + "/" + Left(Left(Year(Now) - 100, 4 - Len(Mid(sText, 5, 4))) + Mid(sText, 5, 4), 4)
      Else
         sText = Mid(sText, 1, 2) + "/" + Mid(sText, 3, 2) + "/" + Left(Left(Year(Now), 4 - Len(Mid(sText, 5, 4))) + Mid(sText, 5, 4), 4)
      End If
   End If
   FormateDate = sText
End Function

