VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ExportRPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmExportRPS
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmExportRPS
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call MontaTela
      'Call PopulaTela
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   Dim NgFin As NG_Financeiro
   Dim sDTINI As String
   Dim sDTFIM As String
   
   sDTINI = Format(mvarMe.CmbDTINI.Value, "dd/mm/yyyy")
   sDTFIM = Format(mvarMe.CmbDTFIM.Value, "dd/mm/yyyy")
   If Trim(sDTINI) <> "" And Trim(sDTFIM) <> "" Then
      Set NgFin = New NG_Financeiro
      Set NgFin.Sys = mvarSys
      Call NgFin.ExportarRPS(pDTINI:=sDTINI, pDTFIM:=sDTFIM)
      Set NgFin = Nothing
      Unload mvarMe
   Else
      Call ExibirInformacao("Data inválida")
   End If
End Sub
Private Sub mvarMe_Load()
   DoEvents
   Call MontaTela
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   Dim sDIni   As String
   Dim sDFim   As String
   Dim sDia    As String
   
   sDia = mvarSys.xDb.SysDate(1)
   If Day(sDia) <= 15 Then
      sDIni = "16" & Format(CDate(sDia) - 20, "/mm/yyyy")
      sDFim = CDate("01" & Format(sDia, "/mm/yyyy")) - 1
   Else
      sDIni = "01" & Format(sDia, "/mm/yyyy")
      sDFim = "15" & Format(sDia, "/mm/yyyy")
   End If
   'sDTINI = InputBox("Informe a Data Inicio.", "Exportar RPS", CDate(sDIni))
   'sDTFIM = InputBox("Informe a Data Fim.", "Exportar RPS", CDate(sDFim))

   mvarMe.CmbDTINI.Value = sDIni
   mvarMe.CmbDTFIM.Value = sDFim
End Sub

