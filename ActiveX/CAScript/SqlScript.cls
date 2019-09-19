VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mvarRUN As RunSqlScript
Attribute mvarRUN.VB_VarHelpID = -1

Private mvarScrFile  As String
Private mvarScrColl  As Collection
Private bTransaction As Boolean
Private mvarxDb      As Object
Private mvarSys      As Object
Private mvarisSystem As Boolean
Public Sub ShowRun(Optional pMode As FormShowConstants = vbModal)
   mvarRUN.Show vbModal
End Sub
Private Sub Class_Initialize()
    Set mvarRUN = New RunSqlScript
End Sub
Private Sub Class_Terminate()
   Set mvarRUN = Nothing
End Sub
Private Sub mvarRun_CmdOkClick()
   Dim nVez       As Integer
   Dim bConectado As Boolean
   Dim i          As Integer
   
   Screen.MousePointer = vbHourglass
   
   mvarRUN.LblStatus.Caption = "Atualizando..."
   
   mvarRUN.PrgBar.Scrolling = 2

   If Not mvarSys Is Nothing Then
      If Not mvarSys.xDb Is Nothing Then
         bConectado = mvarSys.xDb.Conectado
      End If
   End If
   nVez = 3
   If Not bConectado Then
      For i = 1 To nVez
         bConectado = ConectarBanco
         If bConectado Then
            i = nVez
         End If
      Next
   End If
   
   If bConectado Then
      Call RunSqlScript
   End If
         
   mvarRUN.PrgBar.Scrolling = 1
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   'On Error Resume Next
   mvarRUN.LblStatusBD.Caption = ""
   
   
'   Set PrgBar = mvarRUN.Controls.Add("Codejock.ProgressBar.11.2.2", "PrgBar")
'   With PrgBar
'      .Visible = True
'      .Enabled = True
'      .Top = 1180
'      .Height = 255
'      .Left = 240
'      .Width = 6100
'      .UseVisualStyle = True
'      .Scrolling = 0
'   End With
End Sub
Private Function ConectarBanco() As Boolean
   Dim sLocalReg  As String
   Dim sPathReg   As String
   
   If mvarSys Is Nothing Then
      Set mvarSys = CreateObject("SysA.SetA")
   End If
   If Not mvarSys Is Nothing Then
      With mvarSys
         sLocalReg = .LocalReg
         
         If Right(UCase(sLocalReg), 4) <> ".REG" Then
            sPathReg = App.Path & "\"
            sLocalReg = sPathReg & Dir(sPathReg & "*.reg")
         End If
         If Right(UCase(sLocalReg), 4) <> ".REG" Then
            sPathReg = Environ("PROGRAMFILES") & "\ClasseA\Suprimento\"
            sLocalReg = sPathReg & Dir(sPathReg & "*.reg")
         End If
         If Right(UCase(sLocalReg), 4) <> ".REG" Then
            sPathReg = Environ("PROGRAMFILES") & "\ClasseA\Producao\"
            sLocalReg = sPathReg & Dir(sPathReg & "*.reg")
         End If
         If Right(UCase(sLocalReg), 4) <> ".REG" Then
            sPathReg = Environ("PROGRAMFILES") & "\ClasseA\Manutencao\"
            sLocalReg = sPathReg & Dir(sPathReg & "*.reg")
         End If
         
         .LocalReg = IIf(Right(UCase(sLocalReg), 4) = ".REG", sLocalReg, .LocalReg)
         
         If .IDUSU = "" Then .IDUSU = "DIO"
   
         If .xDb.Server = "" Then .xDb.Server = .GetParam("SERVER", "S2", "Conection 0")
         If .xDb.dbName = "" Then .xDb.dbName = .GetParam("DBNAME", "S2", "Conection 0")
         
         .xDb.UID = "CADBA"
         .xDb.PWD = Decrypt2("787A7766756B7F79")
         If Not .xDb.Conectado Then
            .xDb.SrvConecta
            If Not .xDb.Conectado Then
               .xDb.UID = "DBA"
               .xDb.PWD = Decrypt2("787A7766756B7F79")
               .xDb.SrvConecta
               If Not .xDb.Conectado Then
                  .xDb.UID = .GetParam("UID", "S2", "Conection 0")
                  .xDb.PWD = Decrypt2(.GetParam("PWD", "S2", "Conection 0"))
                  .xDb.SrvConecta
               End If
            End If
         End If
            
         If .xDb.Conectado Then
            mvarRUN.LblStatusBD.Caption = "Conectado"
            mvarRUN.LblStatusBD.ForeColor = vbBlue
            mvarisSystem = .USER.isSystem
         Else
            mvarRUN.LblStatusBD.Caption = "Desconectado"
            mvarRUN.LblStatusBD.ForeColor = vbRed
         End If
         'mvarLocalPath = ResolvePathName(IIf(.PathSetup = "", mvarLocalPath, .PathSetup))
         'mvarLocalPathRPT = ResolvePathName(IIf(.PathRPT = "", mvarLocalPathRPT, .PathRPT))
      End With
   End If
   ConectarBanco = mvarSys.xDb.Conectado
End Function
Private Sub RunSqlScript()
   Dim MyCodigo As SqlCodigo
   Dim Trans    As Collection
   Dim Queries  As Collection
   Dim n        As Collection
   Dim QtdSql   As Long
   Dim QtdTrans As Long
   Dim sAux     As String
   
   Set MyCodigo = New SqlCodigo
   Set Trans = MyCodigo.GetTransactions
   
   QtdTrans = 0
   For Each n In Trans
      Set Queries = n
      If Not Queries Is Nothing Then
         If mvarSys.xDb.Executa(Queries, True) Then
            QtdTrans = QtdTrans + 1
            QtdSql = QtdSql + Queries.Count
         End If
      End If
   Next
   If QtdSql > 0 Then
      sAux = "(" & QtdSql & ") " & IIf(QtdSql <= 1, "instrução executada", "instruções executadas")
      sAux = sAux & " com sucesso"
      sAux = sAux & ", em " & QtdTrans & IIf(QtdTrans <= 1, " transação!!", " transações!!")
      mvarRUN.LblStatus.Caption = sAux
   Else
      mvarRUN.LblStatus.Caption = "Transação não executada!"
   End If
End Sub

Private Sub mvarRUN_LblStatusBDDblClick()
   Dim sMsg As String
   
   If Not mvarSys Is Nothing Then
      If Not mvarSys.xDb Is Nothing Then
         sMsg = ""
         sMsg = sMsg & "[" & mvarSys.xDb.Server & "]"
         sMsg = sMsg & ".[" & mvarSys.xDb.dbName & "]"
         sMsg = sMsg & ".[" & mvarSys.xDb.UID & "]"
   
         MsgBox sMsg, vbOKOnly, "Conexão"
      Else
         MsgBox "Objeto Banco não definido", vbOKOnly, "Conexão"
      End If
   Else
      MsgBox "Objeto de configuração não definido", vbOKOnly, "Conexão"
   End If
End Sub

Private Sub mvarRUN_LblTitStatusDblClick()
   Call ConectarBanco
End Sub
Private Sub mvarRUN_Load()
   mvarRUN.LblStatusBD.Caption = ""
   mvarRUN.LblStatus.Caption = ""
End Sub
Private Sub mvarRUN_MnuConnClick()
   Call ConectarBanco
End Sub
