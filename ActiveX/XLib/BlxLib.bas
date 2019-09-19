Attribute VB_Name = "BlxLib"
Option Explicit
Global Const SubFolder = "ClasseA"
Public xDb           As Object
'********************
'** Classes Comuns **
'********************
Public ClsAmbiente      As New Ambiente
Public ClsBanco         As New Banco
Public ClsDetalhes      As New Detalhes
Public ClsGeneral       As New General
Public ClsMensagem      As New Mensagem
Private Sub InstaciaObj()
   If xDb Is Nothing Then
      Set xDb = ClsAmbiente.CriarObjeto("XBANCO01.DS_BANCO")
   End If
End Sub
Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
   Const EM_SETPASSWORDCHAR As Long = &HCC
   Dim EditHwnd As Long

   '* CHANGE APP.TITLE TO YOUR INPUT BOX TITLE.
   EditHwnd = FindWindowEx(FindWindow("#32770", App.Title), 0, "Edit", "")
   Call SendMessage(EditHwnd, EM_SETPASSWORDCHAR, Asc("*"), 0)
   KillTimer hwnd, idEvent
End Sub

