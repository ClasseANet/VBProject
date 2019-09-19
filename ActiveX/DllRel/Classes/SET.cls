VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SETTING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarDrvRpt            As String
Private mvarDrvDrive          As String

Private mvarDbTimeOut         As Integer

Private mvardbDrive_Orig      As String
Private mvarDSVM              As Boolean
Private mvarFundoTela         As String

Private mvarDrvTmp            As String
Private mvarDrvTmpErro        As String
Private mvarDrvTmpImagem      As String

Private mvarMDIFilho          As Object
Private mvarResize            As Boolean

Private mvarIdioma            As eIdioma

Private mvarCODSIS            As String
Private mvarAppExeName        As String
Private mvarAppName           As String
Private mvarAppVer            As String
Private mvarAppTitle          As String
Private mvarAppTimeStamp      As String
Private mvarAppDate           As String

Private mvarNomeEmpresa       As String
Private mvarNomeFilial        As String
Private mvarCodLocal          As String
Private mvarDscLocal          As String

Private mvarIDUSU             As String
Private mvarNMUSU             As String
Private mvarMaskCODCCUSTO     As String
Private mvarMaskCODDESP       As String

Private mvarGrpSistema        As String
Private mvarGrpMaster         As String

Private mvarCancelado         As Boolean

Private mvarExibeLstCadastro  As Boolean
Private mvarExibeToolBar      As Boolean
Private mvarSaiComESC         As Boolean

Private mvarPROJColor         As Double
Private mvarPROJColorSel      As Double

Private mvarNVCARGO           As Collection
Private mvarUSER              As TB_USUARIO

Private mvarCollIDPROJ        As Collection

Private mvarSysA              As Object

Public Property Let DbTimeOut(ByVal vData As Integer)
   mvarDbTimeOut = vData
End Property
Public Property Get DbTimeOut() As Integer
   DbTimeOut = mvarDbTimeOut
End Property
Public Property Let DrvTmp(ByVal vData As String)
   mvarDrvTmp = vData
End Property
Public Property Get DrvTmp() As String
   If Trim(mvarDrvTmp) = "" Then
      mvarDrvTmp = "C:\TMP\" & AppName
   End If
   DrvTmp = mvarDrvTmp
End Property
Public Property Let DrvTmpErro(ByVal vData As String)
   mvarDrvTmpErro = vData
End Property
Public Property Get DrvTmpErro() As String
   If Trim(mvarDrvTmpErro) = "" Then
      mvarDrvTmpImagem = DrvTmp & "\Erro"
   End If
   DrvTmpErro = mvarDrvTmpErro
End Property
Public Property Let DrvTmpImagem(ByVal vData As String)
   mvarDrvTmpImagem = vData
End Property
Public Property Get DrvTmpImagem() As String
   If Trim(mvarDrvTmpImagem) = "" Then
      mvarDrvTmpImagem = DrvTmp & "\Imagem"
   End If
   DrvTmp = mvarDrvTmp
End Property
Public Property Let PROJColor(ByVal vData As Double)
   mvarPROJColor = vData
End Property
Public Property Get PROJColor() As Double
   If mvarPROJColor = 0 Then
      mvarPROJColor = vbWhite
   End If
   PROJColor = mvarPROJColor
End Property
Public Property Let PROJColorSel(ByVal vData As Double)
   mvarPROJColorSel = vData
End Property
Public Property Get PROJColorSel() As Double
      mvarPROJColorSel = &H80FF&
   PROJColorSel = mvarPROJColorSel
End Property
Public Function FormatRM(ByVal Codigo) As String
   Dim sAux As String
   
   Codigo = UnFormat(Codigo)
   sAux = Mid(Codigo, 1, 2)
   sAux = sAux & "-" & Mid(Codigo, 3, 4)
   sAux = sAux & "." & Mid(Codigo, 7, 2)
   sAux = sAux & "-" & Mid(Codigo, 9, 4)
   sAux = sAux & "-" & Mid(Codigo, 13, 3)
   sAux = sAux & "-" & Mid(Codigo, 16, 3)
   sAux = sAux & "-" & Mid(Codigo, 19, 3)
   FormatRM = sAux
End Function
Public Property Get Caption() As String
   Dim cVerNum As String
   If InStr(mvarAppVer, " ") <> 0 Then
      cVerNum = " Vs." & Mid(mvarAppVer, InStr(mvarAppVer, " ") + 1)
   End If
   Caption = AppTitle & cVerNum
End Property
Public Property Let SaiComESC(ByVal vData As Boolean)
   mvarSaiComESC = vData
End Property
Public Property Get SaiComESC() As Boolean
   SaiComESC = mvarSaiComESC
End Property
Public Property Let ExibeLstCadastro(ByVal vData As Boolean)
   mvarExibeLstCadastro = vData
End Property
Public Property Get ExibeLstCadastro() As Boolean
   ExibeLstCadastro = mvarExibeLstCadastro
End Property
Public Property Let ExibeToolBar(ByVal vData As Boolean)
   mvarExibeToolBar = vData
End Property
Public Property Get ExibeToolBar() As Boolean
   ExibeToolBar = mvarExibeToolBar
End Property
Public Property Get GrpMaster() As String
   GrpMaster = mvarGrpMaster
End Property
Public Property Get GrpSistema() As String
   GrpSistema = mvarGrpSistema
End Property
Public Property Let Cancelado(ByVal vData As String)
   mvarCancelado = vData
End Property
Public Property Get Cancelado() As String
   Cancelado = mvarCancelado
End Property
Public Property Let Idioma(ByVal vData As eIdioma)
   mvarIdioma = vData
End Property
Public Property Get Idioma() As eIdioma
   Idioma = mvarIdioma
End Property
Public Property Let NomeEmpresa(ByVal vData As String)
   mvarNomeEmpresa = vData
End Property
Public Property Get NomeEmpresa() As String
   NomeEmpresa = mvarNomeEmpresa
End Property
Public Property Let NomeFilial(ByVal vData As String)
   mvarNomeFilial = vData
End Property
Public Property Get NomeFilial() As String
   NomeFilial = mvarNomeFilial
End Property
Public Property Let CodLocal(ByVal vData As String)
   mvarCodLocal = vData
End Property
Public Property Get CodLocal() As String
   CodLocal = mvarCodLocal
End Property
Public Property Let DscLocal(ByVal vData As String)
   mvarDscLocal = vData
End Property
Public Property Get DscLocal() As String
   DscLocal = mvarDscLocal
End Property
Public Property Let AppTimeStamp(ByVal vData As String)
   mvarAppTimeStamp = vData
End Property
Public Property Get AppTimeStamp() As String
   mvarAppTimeStamp = mvarAppName & "-" & mvarAppDate & " "
   AppTimeStamp = mvarAppTimeStamp
End Property
Public Property Let AppTitle(ByVal vData As String)
   mvarAppTitle = vData
End Property
Public Property Get AppTitle() As String
   AppTitle = mvarAppTitle
End Property
Public Property Let AppVer(ByVal vData As String)
   mvarAppVer = vData
End Property
Public Property Get AppVer() As String
   AppVer = mvarAppVer
End Property
Public Property Let AppName(ByVal vData As String)
   mvarAppName = vData
End Property
Public Property Get AppName() As String
   mvarAppName = IIf(mvarAppName = "", App.ProductName, mvarAppName)
   AppName = mvarAppName
End Property
Public Property Let AppExeName(ByVal vData As String)
   mvarAppExeName = vData
End Property
Public Property Get AppExeName() As String
   mvarAppName = IIf(mvarAppName = "", App.EXEName, mvarAppName)
   AppExeName = mvarAppExeName
End Property
Public Property Let AppDate(ByVal vData As String)
   mvarAppDate = vData
End Property
Public Property Get AppDate() As String
   AppDate = mvarAppDate
End Property
Public Property Set MDIFilho(ByVal vData As Object)
   Set mvarMDIFilho = vData
End Property
Public Property Get MDIFilho() As Object
   Set MDIFilho = mvarMDIFilho
End Property
Public Property Let FundoTela(ByVal vData As String)
   mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
   FundoTela = mvarFundoTela
End Property
Public Property Let Resize(ByVal vData As Boolean)
   mvarResize = vData
End Property
Public Property Get Resize() As Boolean
   Resize = mvarResize
End Property
Public Property Let DSVM(ByVal vData As Boolean)
   mvarDSVM = vData
End Property
Public Property Get DSVM() As Boolean
   DSVM = mvarDSVM
End Property
Public Property Let dbDrive_Orig(ByVal vData As String)
   mvardbDrive_Orig = vData
End Property
Public Property Get dbDrive_Orig() As String
   dbDrive_Orig = mvardbDrive_Orig
End Property
Public Property Let DrvRpt(ByVal vData As String)
   mvarDrvRpt = vData
End Property
Public Property Get DrvRpt() As String
   DrvRpt = mvarDrvRpt
End Property
Public Property Let DrvDrive(ByVal vData As String)
   mvarDrvDrive = vData
End Property
Public Property Get DrvDrive() As String
   DrvDrive = mvarDrvDrive
End Property
Private Sub Class_Initialize()
   Dim MyCls As New DS_SEGURANCA
   mvarGrpSistema = MyCls.GrpSistema
   mvarGrpMaster = MyCls.GrpMaster
   
   mvarIdioma = PORTUGUES
   mvarCodLocal = "RJ"
   mvarDscLocal = "Rio de Janeiro"
   
   
   mvarDbTimeOut = 15
      
   mvarExibeLstCadastro = False
   Set MyCls = Nothing
   
End Sub
