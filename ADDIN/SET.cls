VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Set_Proj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'Private mvardbDrive As String
'Private mvarDrvRpt As String
'Private mvardbName As String
'Private mvardbODBC As Boolean
'Private mvardbVersao As String
'Private mvardbVersion As String
'Private mvarDrvDrive As String
Private mvarDrvError As String
Private mvardbDrive_Orig As String
Private mvarMICRO As String
Private mvarFundoTela As String
Private mvarMDIFilho As Object

Private mvarAppExe As String
Private mvarAppName As String
Private mvarAppVer As String
Private mvarAppTitle As String
Private mvarAppTimeStamp As String
Private mvarAppDate As String
Private mvarNomeEmpresa As String
Private mvarIdioma As Long
Public Property Let Idioma(ByVal vData As Long)
    mvarIdioma = vData
End Property
Public Property Get Idioma() As Long
    Idioma = mvarIdioma
End Property
Public Property Let NomeEmpresa(ByVal vData As String)
    mvarNomeEmpresa = vData
End Property
Public Property Get NomeEmpresa() As String
    NomeEmpresa = mvarNomeEmpresa
End Property
Public Property Let AppTimeStamp(ByVal vData As String)
    mvarAppTimeStamp = vData
End Property
Public Property Get AppTimeStamp() As String
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
    AppName = "VbEditorUtil " & mvarAppName
End Property
Public Property Let AppExe(ByVal vData As String)
    mvarAppExe = vData
End Property
Public Property Get AppExe() As String
    AppExe = mvarAppExe
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
Public Property Let MICRO(ByVal vData As String)
    mvarMICRO = vData
End Property
Public Property Get MICRO() As String
    MICRO = mvarMICRO
End Property
'Public Property Let dbDrive_Orig(ByVal vData As String)
'    mvardbDrive_Orig = vData
'End Property
'Public Property Get dbDrive_Orig() As String
'    dbDrive_Orig = mvardbDrive_Orig
'End Property
'Public Property Let DrvRpt(ByVal vData As String)
'    mvarDrvRpt = vData
'End Property
'Public Property Get DrvRpt() As String
'    DrvRpt = mvarDrvRpt
'End Property
'Public Property Let DrvDrive(ByVal vData As String)
'    mvarDrvDrive = vData
'End Property
'Public Property Get DrvDrive() As String
'    DrvDrive = mvarDrvDrive
'End Property
Public Property Let DrvErro(ByVal vData As String)
    mvarDrvError = vData
End Property
Public Property Get DrvErro() As String
    DrvErro = mvarDrvError
End Property
'Public Property Let dbVersion(ByVal vData As String)
'    mvardbVersion = vData
'End Property
'Public Property Get dbVersion() As String
'    dbVersion = mvardbVersion
'End Property
'Public Property Let dbVersao(ByVal vData As String)
'    mvardbVersao = vData
'End Property
'Public Property Get dbVersao() As String
'    dbVersao = mvardbVersao
'End Property
'Public Property Let dbODBC(ByVal vData As Boolean)
'    mvardbODBC = vData
'End Property
'Public Property Get dbODBC() As Boolean
'   dbODBC = mvardbODBC
'End Property
'Public Property Let dbName(ByVal vData As String)
'   mvardbName = vData
'End Property
'Public Property Get dbName() As String
'    dbName = mvardbName
'End Property
'Public Property Let dbDrive(ByVal vData As String)
'    mvardbDrive = vData
'End Property
'Public Property Get dbDrive() As String
'    dbDrive = mvardbDrive
'End Property
'* Objetos Pre - definidos
'Private mvarDB As New DS_BANCO
'Private mvarBANCO As New BANCO_SCC
'Private mvarClsUser As New USER
'Private SysMdi As New MdiPrincipal
'Public Property Set DB(ByVal vData As DS_BANCO)
'    Set mvarDB = vData
'End Property
'Public Property Let DB(ByVal vData As DS_BANCO)
'    Set mvarDB = vData
'End Property
'Public Property Get DB() As DS_BANCO
'    DB = mvarDB
'End Property
'Public Property Set BANCO(ByVal vData As BANCO_SCC)
'    Set mvarBANCO = vData
'End Property
'Public Property Let BANCO(ByVal vData As BANCO_SCC)
'    Set mvarBANCO = vData
'End Property
'Public Property Get BANCO() As BANCO_SCC
'    BANCO = mvarBANCO
'End Property
'Public Property Set ClsUser(ByVal vData As USER)
'    Set mvarClsUser = vData
'End Property
'Public Property Let ClsUser(ByVal vData As USER)
'    Set mvarClsUser = vData
'End Property
'Public Property Get ClsUser() As USER
'    ClsUser = mvarClsUser
'End Property
'Public Property Set SysMdi(ByVal vData As MdiPrincipal)
'    Set mvarSysMdi = vData
'End Property
'Public Property Let SysMdi(ByVal vData As MdiPrincipal)
'    Set mvarSysMdi = vData
'End Property
'Public Property Get SysMdi() As MdiPrincipal
'    SysMdi = mvarSysMdi
'End Property
Private Sub Class_Initialize()
   With App
      mvarAppExe = .EXEName
      mvarAppName = .ProductName
      mvarAppTitle = .Title
      mvarAppVer = LoadMsg(49) + Trim(CStr(.Major)) + "." + _
                                Trim(CStr(.Minor))
      mvarNomeEmpresa = .CompanyName
   End With
   
   mvarIdioma = 5000
   mvarFundoTela = "FUNDO"
End Sub
