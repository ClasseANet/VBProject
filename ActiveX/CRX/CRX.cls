VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmView
Attribute mvarMe.VB_VarHelpID = -1

Private mvarReport   As CRAXDDRT.Report
Private mvarRptName  As String
Private mvarRptPath  As String

Private mvarServidor As String
Private mvarBanco    As String
Private mvarUsuario  As String
Private mvarSenha    As String
Public Property Let RptPath(ByVal vData As String)
   mvarRptPath = vData
End Property
Public Property Get RptPath() As String
   RptPath = mvarRptPath
End Property
Public Property Let RptName(ByVal vData As String)
   mvarRptName = vData
End Property
Public Property Get RptName() As String
   RptName = mvarRptName
End Property
Public Property Set Report(ByVal vData As CRAXDDRT.Report)
   Set mvarReport = vData
End Property
Public Property Get Report() As CRAXDDRT.Report
   Set Report = mvarReport
End Property
Public Property Let Servidor(ByVal vData As String)
   mvarServidor = vData
End Property
Public Property Get Servidor() As String
   Servidor = mvarServidor
End Property
Public Property Let Banco(ByVal vData As String)
   mvarBanco = vData
End Property
Public Property Get Banco() As String
   Banco = mvarBanco
End Property
Public Property Let Usuario(ByVal vData As String)
   mvarUsuario = vData
End Property
Public Property Get Usuario() As String
   Usuario = mvarUsuario
End Property
Public Property Let Senha(ByVal vData As String)
   mvarSenha = vData
End Property
Public Property Get Senha() As String
   Senha = mvarSenha
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmView
End Sub
Private Sub mvarMe_Activate()
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   With mvarMe.CRV
      .ReportSource = Report
      .ViewReport
      .Zoom 1
   End With
   Screen.MousePointer = vbDefault
Exit Sub

TrataErro:
   MsgBox "Error : " & CStr(Err.Number) & " - " & Err.Description, vbCritical, Err.Source
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CRVDrillOnGroup(GroupNameList As Variant, ByVal DrillType As CRVIEWERLibCtl.CRDrillType, UseDefault As Boolean)
   mvarMe.CRV.DisplayTabs = True
End Sub
Private Sub mvarMe_Load()
   With mvarMe.CRV
      .DisplayBackgroundEdge = True
      .DisplayBorder = True
      .DisplayGroupTree = False
      .DisplayTabs = False
      .DisplayToolbar = True
            
      .EnableAnimationCtrl = False
      .EnableCloseButton = False
      .EnableDrillDown = True
      .EnableExportButton = True
      .EnableGroupTree = True
      .EnableHelpButton = False
      .EnableNavigationControls = True
      .EnablePopupMenu = True
      .EnablePrintButton = True
      .EnableProgressControl = True
      .EnableRefreshButton = False
      .EnableSearchControl = True
      .EnableSearchExpertButton = False
      .EnableSelectExpertButton = False
      .EnableStopButton = True
      .EnableToolbar = True
      .EnableZoomControl = True
   End With
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   With mvarMe.CRV
      .Top = 60
      .Width = mvarMe.Width - 120
      .Height = mvarMe.Height - 400
   End With
End Sub
Property Let SetFormula(Nome As String, ByVal vData As Variant)
   If TypeName(vData) = "String" Then vData = "'" & vData & "'"
   
   mvarReport.FormulaFields.GetItemByName(Nome).Text = vData
End Property
Public Sub Show(Optional pAbertura As Integer = 0)
   mvarMe.Show pAbertura
End Sub
