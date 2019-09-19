Attribute VB_Name = "BlScript"
Sub Main()
   Dim MySql As SqlScript
   
   Set MySql = New SqlScript
   With MySql
      .ShowRun
   End With
End Sub
