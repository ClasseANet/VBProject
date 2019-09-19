Attribute VB_Name = "Bl_Restore3R"
Global MDI As FrmRestoreP3R
Global gLocalPath As String

Global gLocalReg  As String
Global gSetupFile As String
Global gCaption1  As String
Global gCaption2  As String
Global gCaption3  As String
Global gDebug     As Boolean
Public Sub Main()
   On Error Resume Next
  
   If AppAtiva(App) Then End
   
   On Error GoTo TrataErro
   gDebug = (InStr(UCase(Command$), "DEBUG") <> 0)
   Screen.MousePointer = vbHourglass
   
   gLocalPath = Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\"

   Set MDI = New FrmRestoreP3R
   MDI.Show
   Exit Sub
TrataErro:
   MsgBox Err.Number & " - " & Err.Description
   End
End Sub
'*********
'* Testa se já existe uma cópia da aplicação rodando e define formato Data e número.
Public Function AppAtiva(pApp As App) As Boolean
   Dim MyLoad As Object
   Dim bAtiva As Boolean
   
   bAtiva = False
   Set MyLoad = CriarObjeto("DSACTIVE.DS_LOAD")
   If Not MyLoad Is Nothing Then
      MyLoad.Aplic = App
      If MyLoad.Ativa Then
         bAtiva = True
      End If
   End If
   Set MyLoad = Nothing
   AppAtiva = bAtiva
End Function
