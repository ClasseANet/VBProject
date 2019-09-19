VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SqlPlus"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents mvarMe As FrmSql
Attribute mvarMe.VB_VarHelpID = -1
Private mvarDBase As Object
Private mvarxDb As DS_BANCO ' Object
Private mvarOwners As New Collection
Private mvarOwner As String
Private mvarFundoTela As String
Private mvarIdioma As Integer
Private mvarDefaultOwner As String

Const Msg5001 = " Status : Concluída!!"
Const Msg5002 = " Status : Não Realizada!!"
Const Msg5003 = " Linhas : "
Public Property Let DefaultOwner(ByVal vData As String)
    mvarDefaultOwner = vData
End Property
Public Property Get DefaultOwner() As String
    DefaultOwner = mvarDefaultOwner
End Property
Public Property Let Idioma(ByVal vData As Integer)
    mvarIdioma = vData
End Property
Public Property Get Idioma() As Integer
    Idioma = mvarIdioma
End Property
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let Owner(ByVal vData As String)
    mvarOwner = vData
End Property
Public Property Get Owner() As String
    Owner = mvarOwner
End Property
Public Property Set Owners(ByVal vData As Object)
    Set mvarOwners = vData
End Property
Public Property Get Owners() As Collection
    Set Owners = mvarOwners
End Property
Public Sub Show()
Attribute Show.VB_Description = "Exibe Tela de Consulta ao Banco de Dados."
   mvarMe.Show vbModal
End Sub
Public Property Set DBase(ByVal vData As Object)
Attribute DBase.VB_Description = "Banco de Dados que será consultado."
    Set mvarDBase = vData
End Property
Public Property Get DBase() As Object
    Set DBase = mvarDBase
End Property
Public Property Set xDb(ByVal vData As DS_BANCO)
    Set mvarxDb = vData
End Property
Public Property Get xDb() As DS_BANCO
    Set xDb = mvarxDb
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmSql
   mvarIdioma = 5000
   mvarFundoTela = "FUNDO"
End Sub
Private Sub Class_Terminate()
   Set mvarOwners = Nothing
   Set mvarMe = Nothing
End Sub
'**************************************
'**************************************
'**************************************

Private Sub mvarMe_CmbOwnerClick()
   mvarOwner = mvarMe.CmbOwner
End Sub
Private Sub mvarMe_CmdOperClick(index As Integer)
   Select Case index
      Case 0: Call ExecSql
      Case 1: UnLoad mvarMe
   End Select
End Sub
Private Sub mvarMe_FormActivate()
   Call ClsDsr.SetHourglass(mvarMe.hWnd)
'   Set Sys.MDIFilho = mvarme
   If Not mvarDBase Is Nothing Then
      mvarMe.lbl(2).Caption = " [" & UCase(mvarDBase.DSN) & "." & UCase(mvarDBase.StrDATABASE) & "]"
   ElseIf Not mvarxDb Is Nothing Then
      mvarMe.lbl(2).Caption = " [" & UCase(mvarxDb.Server) & "." & UCase(mvarxDb.dbName) & "]"
   End If
   mvarMe.TxtSql.SetFocus
   Call ClsDsr.SetDefault(mvarMe.hWnd)
'   If Not VerificaAcesso(mvarme.Acesso, LEITURA) Then
'      Unload mvarme
'   End If
End Sub

Private Sub mvarMe_FormLoad()
   Dim i%, Pos%, Txt$, n
   Call ClsDsr.SetHourglass(mvarMe.hWnd)
   mvarMe.GrdSql.Rows = 0
   mvarMe.GrdSql.Cols = 0
   
   If Not mvarDBase Is Nothing Then
      'mvarme.DataEVT.DatabaseName = mvardbase.dBase.Name
      With mvarMe.DataSql
         .DataSourceName = mvarDBase.DSN
         .UserName = mvarDBase.UID
         .Password = mvarDBase.PWD
         If mvarDBase.isODBC Then .CursorDriver = rdUseOdbc
      End With
      
      Call mvarDBase.GetOwners(mvarDBase, mvarOwners)
   ElseIf Not mvarxDb Is Nothing Then
      Call mvarxDb.GetOwners(mvarxDb, mvarOwners)
   End If
   
   Call ClsCtrl.MontarCombo(mvarMe.CmbOwner, mvarOwners)
   If mvarDefaultOwner <> "" Then
      Call ClsCtrl.LocalizarCombo(mvarMe.CmbOwner, mvarDefaultOwner)
   End If
   
   Call ClsCtrl.ConfigForm(mvarMe, "", mvarFundoTela)
   Call ClsDsr.SetDefault(mvarMe.hWnd)
End Sub
Public Sub ExecSql()
   Dim MontouGrid  As Boolean
   Dim Sql As String
   Dim Cab, Pos As Integer
   
   
   Screen.MousePointer = vbHourglass
   mvarMe.GrdSql.Rows = 2
   mvarMe.GrdSql.Cols = 2
   mvarMe.GrdSql.FixedRows = 1
   Sql = ClsDsr.StrReplace(mvarMe.TxtSql, Chr(13) & Chr(10), " ")
   Sql = UCase(Trim(ClsDsr.StrReplace(Sql, """", "'")))
   Pos = InStr(Sql, ";")
   While Pos <> 0
      If Pos = Len(Sql) Then
         Sql = Trim(Mid(Sql, 1, Pos - 1))
         Pos = 0
      Else
         Sql = Trim(Mid(Sql, Pos + 1))
         Pos = InStr(Sql, ";")
      End If
   Wend
   Pos = InStr(Trim(Sql), " ")
   If Pos = 0 Then GoTo Saida
   Select Case Mid(Sql, 1, Pos - 1)
      Case "SELECT"
         If Not mvarDBase Is Nothing Then
            MontouGrid = ClsMSGrid.MontarMSGrid(mvarMe.DataSql, mvarMe.GrdSql, Cab, Sql)
         ElseIf Not mvarxDb Is Nothing Then
            MontouGrid = ClsMSGrid.MontarMSGrid(xDb.RSAux, mvarMe.GrdSql, Cab, Sql)
         End If
         If MontouGrid Then
            mvarMe.lbl(0).ForeColor = vbBlue
            mvarMe.lbl(0).Caption = Msg5001
            If Not mvarDBase Is Nothing Then
               mvarMe.lbl(1).Caption = Msg5003 & CStr(mvarMe.DataSql.Resultset.RowCount)
            ElseIf Not mvarxDb Is Nothing Then
               mvarMe.lbl(1).Caption = Msg5003 & CStr(xDb.RSAux.RecordCount)
            End If
            mvarMe.GrdSql.SetFocus
         Else
            mvarMe.lbl(0).ForeColor = vbRed
            mvarMe.lbl(0).Caption = Msg5002
            mvarMe.lbl(1).Caption = Msg5003 & "0"
            mvarMe.TxtSql.SetFocus
         End If
      Case "INSERT", "UPDATE", "DELETE"
         If mvarDBase.Executa(Sql) Then
            mvarMe.lbl(0).ForeColor = vbBlue
            mvarMe.lbl(0).Caption = Msg5001
            mvarMe.lbl(1).Caption = Msg5003 & CStr(mvarMe.DataSql.Resultset.RowCount)
         Else
            mvarMe.lbl(0).ForeColor = vbRed
            mvarMe.lbl(0).Caption = Msg5002
            mvarMe.lbl(1).Caption = Msg5003 & "0"
         End If
         mvarMe.TxtSql.SetFocus
      Case "DESC"
         Tabela = Mid(Sql, Pos + 1)
         On Error Resume Next
         Aux = mvarDBase.Tables(mvarOwner & IIf(mvarOwner = "", "", ".") & Tabela)
         If Err = 3265 Then
            On Error GoTo 0
            Screen.MousePointer = vbDefault
            GoTo Saida
         Else
            On Error GoTo Fim
         End If
         With mvarDBase.Tables(mvarOwner & IIf(mvarOwner = "", "", ".") & Tabela)
            mvarMe.GrdSql.Rows = .Fields.Count + 1
            mvarMe.GrdSql.Cols = 3
            mvarMe.GrdSql.TextMatrix(0, 0) = "CAMPO"
            mvarMe.GrdSql.TextMatrix(0, 1) = "TIPO"
            mvarMe.GrdSql.TextMatrix(0, 2) = "TAMANHO"

            For i = 1 To .Fields.Count
               mvarMe.GrdSql.TextMatrix(i, 0) = .Fields(i).Nome
               Tam = IIf(Tam > Len(.Fields(i).Nome), Tam, Len(.Fields(i).Nome))
               Select Case GrpTipoCampo(.Fields(i).Tipo)
                  Case 1: mvarMe.GrdSql.TextMatrix(i, 1) = "NUMÉRICO"
                  Case 2: mvarMe.GrdSql.TextMatrix(i, 1) = "DATA"
                  Case 3: mvarMe.GrdSql.TextMatrix(i, 1) = "CARACTER"
               End Select
               mvarMe.GrdSql.TextMatrix(i, 2) = .Fields(i).Tam
            Next
            mvarMe.GrdSql.ColWidth(0) = Tam * 120
            mvarMe.GrdSql.ColWidth(1) = 9 * 120
            mvarMe.GrdSql.ColWidth(2) = 8 * 120
            mvarMe.lbl(1).Caption = " Linhas : " & CStr(.Fields.Count)
         End With
         mvarMe.GrdSql.SetFocus
      Case "?"
        
   End Select
   mvarMe.GrdSql.FixedCols = 0
   Screen.MousePointer = vbDefault
Saida:
   Screen.MousePointer = vbDefault
   Exit Sub
Fim:
   ClsMsg.ShowError
   GoTo Saida
End Sub

Private Sub mvarmvarMe_FormResize()
   Call ClsCtrl.PintarFundo(mvarMe, mvarFundoTela)
   mvarMe.CmdOper(0).Move 120, mvarMe.Height - mvarMe.CmdOper(0).Height - 480
   mvarMe.CmdOper(1).Move mvarMe.Width - mvarMe.CmdOper(1).Width - 240, mvarMe.Height - mvarMe.CmdOper(1).Height - 480
End Sub

Private Sub mvarMe_GrdSqlClick()
   If mvarMe.GrdSql.MouseRow = 0 Then
      Call ClsMSGrid.OrdenarMSGrid(mvarMe.DataSql, mvarMe.GrdSql, mvarMe.GrdSql.MouseCol)
   End If
End Sub

Private Sub mvarMe_GrdSqlDblClick()
   On Error Resume Next
   mvarMe.TxtSql = Trim(mvarMe.TxtSql)
   mvarMe.TxtSql = mvarMe.TxtSql & " " & mvarMe.GrdSql.TextMatrix(mvarMe.GrdSql.MouseRow, mvarMe.GrdSql.MouseCol)
End Sub

Private Sub mvarMe_GrdSqlKeyPress(KeyAscii As Integer)
   mvarMe.TxtSql.SetFocus
   SendKeys "^{END}" & Chr(KeyAscii)
End Sub
Private Sub mvarMe_TxtSqlKeyPress(KeyAscii As Integer)
   Dim Sql$
   If KeyAscii = vbKeyReturn Then
      Sql = Trim(mvarMe.TxtSql)
      If Mid(Sql, 1, 1) = ">" Then
         Sql = Mid(mvarMe.TxtSql, 2)
      End If
      Sql = ClsDsr.StrReplace(Sql, Chr(13) & Chr(10), " ")
      If Right(Trim(Sql), 1) = ";" Then Call ExecSql
   End If
End Sub
