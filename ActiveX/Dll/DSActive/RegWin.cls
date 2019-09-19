VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegWin"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents mvarMe As FrmRegister
Attribute mvarMe.VB_VarHelpID = -1

Private mvarAplic As Object
Private mvarFundoTela As String
Private mvarNomeArq As String
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let NomeArq(ByVal vData As String)
    mvarNomeArq = vData
End Property
Public Property Get NomeArq() As String
    NomeArq = mvarNomeArq
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Public Property Let Aplic(ByVal vData As Object)
Attribute Aplic.VB_Description = "Objeto ""App"" relacionada a sua Aplicação."
    mvarAplic = vData
End Property
Public Property Get Aplic() As Object
    Set Aplic = mvarAplic
End Property

Private Sub Class_Initialize()
   Set mvarMe = New FrmRegister
End Sub

Private Sub mvarMe_Activate()
   mvarMe.Move 0, 0
End Sub

Private Sub mvarMe_CmdOperClick(Index As Integer)
   Select Case Index
      Case 0: Call ClsDsr.RegServer(mvarMe.TxtArq, True)
      Case 1: Call ClsDsr.RegServer(mvarMe.TxtArq, False)
      Case 2: Unload mvarMe
   End Select
End Sub

Private Sub mvarMe_ImgOpen()
   Dim Filtro$, Arq$
   Filtro = "Files (*.dll;*.ocx;*.exe)|*.dll;*.ocx;*.exe"
   Arq$ = ProcurarArquivo(mvarMe.CommonDialog1, "Finde File", "", Filtro)
   mvarMe.TxtArq = mvarMe.CommonDialog1.Tag & Arq$
End Sub

Private Sub mvarMe_Load()
   Call ClsDsr.SetHourglass(mvarMe.hWnd)
   mvarMe.ImgOpen.BorderStyle = 0
   
   Call ClsCtrl.ConfigForm(mvarMe, , mvarFundoTela)
   Call ClsDsr.SetDefault(mvarMe.hWnd)
End Sub
