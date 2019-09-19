VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORDERBY"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Private WithEvents mvarMe As FrmOrderBy
Attribute mvarMe.VB_VarHelpID = -1
Private mvarCaption$

Private mvarFundoTela As String
Private mvarIdioma As Integer
Private mvarOrderBy As String
Private mvarGrdLov As Object

Private VetTabela As Collection

Public Property Set GrdLov(ByVal vData As Object)
    Set mvarGrdLov = vData
End Property
Public Property Get GrdLov() As Object
    Set GrdLov = mvarGrdLov
End Property
Public Property Let Idioma(ByVal vData As Integer)
    mvarIdioma = vData
End Property
Public Property Get Idioma() As Integer
    Idioma = mvarIdioma
End Property
Public Property Let OrderBy(ByVal vData As Integer)
    mvarOrderBy = vData
End Property
Public Property Get OrderBy() As Integer
    OrderBy = mvarOrderBy
End Property
Public Sub Show()
   Set mvarMe = New FrmOrderBy
   mvarMe.Show vbModal
End Sub
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Let Caption(Txt$)
    mvarCaption = Txt$
End Property
Private Sub mvarMe_Activate()
   '* Definir o Textos Form
   Call FormataTela
      mvarMe.Visible = True
   mvarMe.Move mvarMe.Left, mvarMe.Top + 660

End Sub

Private Sub mvarMe_ChkSelectedClick()
   Dim i%
   If mvarMe.ChkSelected.Value = vbChecked Then
      mvarMe.LstCampos.Visible = False
      mvarMe.LstCampos.Refresh
      Set VetTabela = New Collection
      For i = mvarMe.LstCampos.ListCount - 1 To 0 Step -1
         If mvarMe.LstCampos.Selected(i) Then
            VetTabela.Add mvarMe.LstCampos.ItemData(i)
         Else
            mvarMe.LstCampos.RemoveItem i
        End If
      Next
      mvarMe.LstCampos.Visible = True
      mvarMe.LstCampos.Refresh
   Else
      Call MontaListCampos
   End If
End Sub
Private Sub mvarMe_FrmKeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyReturn: Call mvarMe_CmdOper(0)
'      Case vbKeyDelete: Call mvarMe_Excluir
      Case vbKeyEscape: UnLoad mvarMe
      Case Else: KeyAscii = ClsDsr.SendTab(mvarMe, KeyAscii)
   End Select
   DoEvents
End Sub
Private Sub mvarMe_FrmKeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn: mvarMe_CmdOper (0)
      Case vbKeyInsert
      Case vbKeyEscape: UnLoad mvarMe
   End Select
End Sub
Private Sub mvarMe_Load()
   Screen.MousePointer = vbHourglass
   Call MontaListCampos
   Call ConfigLov
   Call ClsCtrl.ConfigForm(mvarMe, , mvarFundoTela)
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdOper(Index As Integer)
   mvarOrderBy = ""
   Select Case Index
      Case 0
         For i = 0 To mvarMe.LstCampos.ListCount - 1
            If mvarMe.LstCampos.Selected(i) Then
               mvarOrderBy = IIf(mvarOrderBy = "", "", mvarOrderBy & ", ")
               mvarOrderBy = mvarOrderBy & mvarMe.LstCampos.List(i)
            End If
         Next
         UnLoad mvarMe
      Case 1: UnLoad mvarMe
   End Select
End Sub
Private Sub FormataTela()
   mvarMe.Caption = mvarCaption
   mvarMe.CmdLovOper(1).Caption = ClsMsg.LoadMsg(mvarIdioma + 5015)  'Cancela
   mvarMe.CmdLovOper(0).Caption = ClsMsg.LoadMsg(mvarIdioma + 5018) 'OK
   On Error Resume Next
   DSR100.Idioma = mvarIdioma
End Sub
Private Sub ConfigLov()
'================================================================
'= Última Alteração : 10/07/98                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Configurar Form de Lista de valores.            **
'**                                                            **
'** Recebe: F - Form a ser configurado                         **
'**                                                            **
'** Retorna:                                                   **
'**                                                            **
'****************************************************************
   Dim t&, L&
   On Error GoTo Fim
   Exit Sub
Fim:
   ClsMsg.ShowError
End Sub
Private Sub mvarMe_QueryUnload(Cancel As Variant, UnloadMode As Variant)
   If UnloadMode = 0 Then mvarOrderBy = ""
End Sub

Private Sub mvarMe_Unload(Cancel As Integer)
   DoEvents
   Call Class_Terminate
End Sub
Public Sub MontaListCampos()
   Dim i%
   Dim Grd As MSFlexGrid
   mvarMe.LstCampos.Clear
   For i = 0 To mvarGrdLov.Cols - 1
      'If mvarGrdLov.ColWidth(i) > 0 And mvarGrdLov.ColIsVisible(i) Then
      If mvarGrdLov.ColWidth(i) Then
         mvarMe.LstCampos.AddItem mvarGrdLov.TextMatrix(0, i)
         mvarMe.LstCampos.ItemData(i) = i
      End If
   Next
   If Not VetTabela Is Nothing Then
      For Each n In VetTabela
         mvarMe.LstCampos.Selected(n) = True
      Next
   End If
End Sub
Public Function SetOrderBy(Grd As Object)
   Set GrdLov = Grd
   Call Show
   SetOrderBy = ""
   If mvarOrderBy <> "" Then
      SetOrderBy = " Order by " & mvarOrderBy
  End If
End Function
Private Sub Class_Initialize()
   mvarIdioma = 5000 'Português
   mvarFundoTela = "FUNDO"
   mvarCaption = "Ordenação Múltipla"
   Set mvarMe = New FrmOrderBy
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarMe = Nothing
   Set VetTabela = Nothing
End Sub


