VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_VIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe As FrmView
Attribute mvarMe.VB_VarHelpID = -1

Public mvarxDb       As Object
Public mvarFundoTela As String
Public mvarIcon      As Variant
Public mvarArquivo   As String
Public mvarUsuario   As String
Public mvarCaption   As String
Public Property Let Caption(ByVal vData As String)
    mvarCaption = vData
End Property
Public Property Get Caption() As String
    Caption = mvarCaption
End Property
Public Property Let Arquivo(ByVal vData As String)
    mvarArquivo = vData
End Property
Public Property Get Arquivo() As String
    Arquivo = mvarArquivo
End Property
Public Property Let Usuario(ByVal vData As String)
    mvarUsuario = vData
End Property
Public Property Get Usuario() As String
    Usuario = mvarUsuario
End Property
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let ICON(ByVal vData As Variant)
    mvarIcon = vData
End Property
Public Property Get ICON() As Variant
    ICON = mvarIcon
End Property
Public Property Set xDb(ByVal vData As Object)
    Set mvarxDb = vData
End Property
Public Property Get xDb() As Object
    Set xDb = mvarxDb
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmView
   mvarFundoTela = "FUNDO"
End Sub
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub mvarMe_Activate()
   mvarMe.ImageAss.Stretch = False
   mvarMe.Width = 1.03 * (mvarMe.ImageAss.Width + mvarMe.ImageAss.Left)
   mvarMe.Height = 1.2 * (mvarMe.ImageAss.Height + mvarMe.ImageAss.Top)
   
   mvarMe.Move mvarMe.Left, mvarMe.Top, mvarMe.ImageAss.Width + mvarMe.ImageAss.Left + 860, mvarMe.ImageAss.Height + mvarMe.ImageAss.Top + 1200
   
   Call ClsCtrl.CentrarObj(Me, mvarMe.ImageAss)
   Call ClsCtrl.PintarFundo(Me, mvarFundoTela)
   
   mvarMe.ImageAss.Left = mvarMe.ImageAss.Left - 1
   
End Sub
Private Sub mvarMe_Load()
   If Trim(mvarCaption) <> "" Then
      mvarMe.Caption = mvarCaption
   End If
   
   Call BinaryToFile(mvarxDb, , mvarArquivo, "USUARIO", "ASSINATURA", "IDUSU=" & SqlStr(mvarUsuario))
   If FileLen(mvarArquivo) <> 0 Then
      mvarMe.ImageAss.Stretch = False
      mvarMe.ImageAss.Picture = LoadPicture(mvarArquivo)
      mvarMe.ImageAss.Tag = mvarArquivo
   Else
      mvarMe.ImageAss.Tag = ""
      mvarMe.ImageAss = Nothing
   End If
   Call ClsDos.Del(mvarArquivo)
   
   Call ClsCtrl.ConfigForm(mvarMe, mvarIcon, mvarFundoTela)
End Sub
Private Sub mvarMe_Resize()
   Call ClsCtrl.CentrarObj(Me, mvarMe.ImageAss)
   Call ClsCtrl.PintarFundo(Me, mvarFundoTela)
   mvarMe.ImageAss.Left = mvarMe.ImageAss.Left - 1
End Sub
