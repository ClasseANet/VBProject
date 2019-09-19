VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContaRel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Registra LOG de Impressão
Public Function RegistraLog(Optional ByRef pQueries As Collection, Optional ByRef pNMREL As String, Optional pIDUSU = "") As String
                            
   Dim MyMaq                     As New DSACTIVE.DSR
   Dim MyLOG                     As New TB_REGEVTPRINT
      
   If pQueries Is Nothing Then Set pQueries = New Collection
   
   If pIDUSU = "" Then
       pIDUSU = Sys.IDUSU
   End If
   Set MyLOG.XDb = mvarXDb
   With MyLOG
      .CODSIS = Sys.CODSIS
      .NMREL = pNMREL
      .QTPRINT = xVal(.QTPRINT) + 1
      .IDUSU = pIDUSU
      .MAQ = MyMaq.ComputerName
      .DTLOG = XDb.Sysdate()
      .HORALOG = XDb.Sysdate(eSysDate.HORA)
      pQueries.Add .QryInsert
      
   End With
   
   RegistraLog = ""
   
End Function


