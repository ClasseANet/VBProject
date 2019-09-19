Attribute VB_Name = "P3R"
Option Explicit
Public gCODSIS    As String

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
   Set MyLoadPicture = LoadPicture(App.Path & "\" & gCODSIS & ".ico")
   
   
   '* Erro ao carrebar MDI
   'Set MyLoadPicture = MDI.imgList.ListImages(gCODSIS).Picture
   'Set MyLoadPicture = MDI.Picture
   
Exit Function
TrataErro:
   MsgBox Err & "-" & Error
End Function
Public Function MyLoadgCODSIS() As Object
   gCODSIS = "P3R"
   gIDUSU = ""   '* Para não exibir Splash se define gIDUSU="DIO"
   gCaption1 = "Projeto 3R"
   gCaption2 = "Fotodepilação"
   gCaption3 = "Inteligente"
End Function
Public Sub MyLimpaInstaciaObj()
   If gDebug Then MsgBox "LimpaInstaciaObj"
   On Error GoTo TrataErro
         
   Set MdiMenu = Nothing
         
   Exit Sub
TrataErro:
   DsMsg.ShowError "Limpa Instacia de Objetos"
End Sub

Public Sub MyInstaciaObj()
   Dim TpErro  As String
   Dim Erro429 As Boolean
   
   If gDebug Then MsgBox "MyInstaciaObj..."
   
   On Error GoTo TrataErro
   TpErro = "MdiMenu"
   Set MdiMenu = CreateObject("Menu.ControlMenu")
   If gDebug Then MsgBox "Criou Menu"
   
   GoTo Saida
TrataErro:
   If Err = 429 Then
      If gDebug Then MsgBox "Não Criou Menu"
      Erro429 = True
      Resume Next
   Else
      DsMsg.ShowError "Instacia de Objetos [" & TpErro & "]"
   End If
Saida:
   If Erro429 Then
      Err = 429
      DsMsg.ShowError "Instacia de Objetos [" & TpErro & "]"
   End If
End Sub
