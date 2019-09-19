VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarPathUpdate  As String
Private mvarIDUSU       As String
Private mvarNMPESSOA    As String
Private mvarAPELIDO     As String
Private mvarIDPESSOA    As Double
Private mvarAppExeName  As String
Private mvarDrvRpt      As String
Private mvarDSVM        As Boolean
Private mvarIdioma      As Double
Private mvarCODSIS      As String
Private mvarPortaSis    As Double
Private mvarPortaPadrao As Double
Public Property Let PathUpdate(ByVal vData As String)
   mvarPathUpdate = vData
End Property
Public Property Get PathUpdate() As String
   PathUpdate = mvarPathUpdate
End Property
Public Property Let PortaPadrao(ByVal vData As Double)
   mvarPortaPadrao = vData
End Property
Public Property Get PortaPadrao() As Double
   PortaPadrao = mvarPortaPadrao
End Property
Public Property Let PortaSis(ByVal vData As Double)
   mvarPortaSis = vData
End Property
Public Property Get PortaSis() As Double
   PortaSis = mvarPortaSis
End Property
Public Property Get NMPESSOA() As String
   NMPESSOA = mvarNMPESSOA
End Property
Public Property Get APELIDO() As String
   APELIDO = mvarAPELIDO
End Property
Public Property Get IDPESSOA() As String
   IDPESSOA = mvarIDPESSOA
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarAppExeName
   CODSIS = "MCIM"
End Property
Public Property Let IDUSU(ByVal vData As String)
   Dim Sql As String
   If mvarIDUSU <> vData Then
      Sql = "Select * from PESSOA Where CODPESSOA = " & SqlStr(vData)
      If XDb.AbreTabela(Sql) Then
         mvarNMPESSOA = XDb.RSAux("NMPESSOA") & ""
         mvarAPELIDO = XDb.RSAux("APELIDO") & ""
         mvarIDPESSOA = XDb.RSAux("IDPESSOA") & ""
         If mvarAPELIDO = "" Then mvarAPELIDO = mvarNMPESSOA
      End If
   End If
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let AppExeName(ByVal vData As String)
   mvarAppExeName = vData
End Property
Public Property Get AppExeName() As String
   If mvarAppExeName = "" Then mvarAppExeName = App.EXEName
   mvarAppExeName = UCase(mvarAppExeName)
   AppExeName = mvarAppExeName
End Property
Public Property Let DrvRpt(ByVal vData As String)
   mvarDrvRpt = vData
End Property
Public Property Get DrvRpt() As String
   DrvRpt = mvarDrvRpt
End Property
Public Property Let DSVM(ByVal vData As Boolean)
   mvarDSVM = vData
End Property
Public Property Get DSVM() As Boolean
   DSVM = mvarDSVM
End Property
Public Property Let Idioma(ByVal vData As Double)
   mvarIdioma = vData
End Property
Public Property Get Idioma() As Double
   Idioma = mvarIdioma
End Property
Private Sub Class_Initialize()
   mvarCODSIS = "PMINFO"
   mvarPortaSis = 100
   mvarPortaPadrao = 200
End Sub
