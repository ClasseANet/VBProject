VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ExportMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mvarMe    As FrmExportMov
Attribute mvarMe.VB_VarHelpID = -1

Dim nIDCONTA        As Integer

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDCONTA        As Integer
Public Property Let IDCONTA(ByVal vData As Integer)
    mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Integer
   IDCONTA = mvarIDCONTA
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   If mvarIDCONTA = 0 Then mvarIDCONTA = mvarSys.Propriedades("IDCONTA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   mvarMe.Show pMode
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmExportMov
End Sub

Private Sub mvarMe_Activate()
'   MontaTela
End Sub

Private Sub mvarMe_CmdExportarClick()
   Dim MyMoney As NG_Money
   Dim MyRs As Object
   Dim Sql As String
   Dim sPath As String
   Dim bExportou As Boolean
   Dim bOk As Boolean
   Dim bNDoc As Boolean
   Dim sArqName As String
   Dim sCtrl As String
     
   Sql = "Select D.DSCDESP, D2.DSCDESP [DSCSUBDESP], C2.DSCCONTA [DSCTRANSF], L.*" & vbNewLine
   Sql = Sql & " From FLAN L" & vbNewLine
   Sql = Sql & " Left Join FDESPESA D     On  L.IDLOJA= D.IDLOJA And L.IDDESP=D.IDDESP" & vbNewLine
   Sql = Sql & " Left Join FDESPESA D2    On  L.IDLOJA=D2.IDLOJA And L.IDSUBDESP=D2.IDDESP" & vbNewLine
   Sql = Sql & " Left Join FLAN L2        On  L.IDLOJA=L2.IDLOJA And L.IDTRANSF=L2.IDLAN" & vbNewLine
   Sql = Sql & " Left Join FCCORRENTE C2  On L2.IDLOJA=C2.IDLOJA And L2.IDCONTA=C2.IDCONTA" & vbNewLine
   Sql = Sql & " Where L.IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And L.IDCONTA=" & nIDCONTA & vbNewLine
   Sql = Sql & " And L.DTVENCIMENTO>=" & SqlDate(Format(mvarMe.CmbDTINI.Value, "dd/mm/yyyy") & " 00:00:00") & vbNewLine
   Sql = Sql & " And L.DTVENCIMENTO<=" & SqlDate(Format(mvarMe.CmbDTFIM.Value, "dd/mm/yyyy") & " 23:59:59")
   With mvarSys.xdb
      bOk = False
      If .AbreTabela(Sql, MyRs) Then
         If MyRs.RecordCount > 0 Then
            bOk = True
         End If
      End If
      If bOk Then
         sCtrl = Sys.Propriedades("RZABREV") & "_" & Format(mvarMe.CmbDTINI.Value, "yy") & StrZero(DateDiff("d", "01/01/" & Year(mvarMe.CmbDTINI.Value), mvarMe.CmbDTINI.Value), 3)
         sArqName = mvarSys.CODSIS & "_" & sCtrl & "_" & Format(mvarMe.CmbDTINI.Value, "ddmmm") & "_" & Format(mvarMe.CmbDTFIM.Value, "ddmmm")
         Select Case True
            Case mvarMe.OptFormato(0).Value
               sArqName = sArqName & ".qif"
               mvarMe.CommonDialog1.Filter = "(*.qif)|*.qif"
            Case mvarMe.OptFormato(1).Value
               sArqName = sArqName & ".ofc"
               mvarMe.CommonDialog1.Filter = "(*.ofc)|*.ofc"
            Case mvarMe.OptFormato(2).Value
               'sArqName = sArqName & ".ofc"
               'mvarMe.CommonDialog1.Filter = "(*.ofc)|*.ofc"
         End Select
         bNDoc = (mvarMe.ChkNDOC.Value = xtpChecked)
         On Error Resume Next
         mvarMe.CommonDialog1.DialogTitle = "Arquivos para Money/Quicken"
         mvarMe.CommonDialog1.FileName = sArqName
         mvarMe.CommonDialog1.InitDir = mvarSys.PathTmp
         Call mvarMe.CommonDialog1.ShowSave
         If mvarMe.CommonDialog1.FileTitle <> "" Then
            sArqName = GetNameFromPath(mvarMe.CommonDialog1.FileName, sPath)
         End If
         If sArqName <> "" Then
            sArqName = GetNameFromPath(mvarMe.CommonDialog1.FileName, sPath)
            If sPath = "" Then
               sPath = mvarSys.PathTmp
            End If
         
            Set MyMoney = New NG_Money
            With MyMoney
               Set .Sys = mvarSys
               Select Case True
                  Case mvarMe.OptFormato(0).Value: bExportou = .GerarQIFFromRs(MyRs, sPath, sArqName, bNDoc:=bNDoc, bMsg:=True)
                  Case mvarMe.OptFormato(1).Value: bExportou = .GerarOFCFromRs(MyRs, sPath, sArqName, bNDoc:=bNDoc, bMsg:=True)
               End Select
               If bExportou Then
                  Sql = "Update FLAN"
                  Sql = Sql & " Set FLGEXPORT=1"
                  Sql = Sql & " , ALTERSTAMP=1"
                  Sql = Sql & " , TIMESTAMP=GETDATE()"
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And IDCONTA=" & nIDCONTA
                  Sql = Sql & " And DTVENCIMENTO>=" & SqlDate(Format(mvarMe.CmbDTINI.Value, "dd/mm/yyyy") & " 00:00:00")
                  Sql = Sql & " And DTVENCIMENTO<=" & SqlDate(Format(mvarMe.CmbDTFIM.Value, "dd/mm/yyyy") & " 23:59:59")
                  Call mvarSys.xdb.Executa(Sql)
               End If
            End With
            Unload mvarMe
         End If
      Else
         Call ExibirInformacao("Não existe movimento a ser exportado neste período.", "Exportar Movimento")
      End If
   End With
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_Load()
   mvarMe.Top = 1500
   mvarMe.Left = mvarMe.Top
   Call MontaTela
End Sub
Private Sub MontaTela()
   Dim Sql As String
   Dim MyRs As Object
   
   mvarMe.CmdExportar.Enabled = False
   mvarMe.CmbDTFIM.Enabled = False
   mvarMe.CmbDTINI.Enabled = False

   
   Sql = "Select * " & vbNewLine
   Sql = Sql & " From FCCORRENTE" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And (ATIVO=1 Or ATIVO is Null)"
   mvarMe.LstContas.Clear
   With mvarSys.xdb
      If .AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            mvarMe.LstContas.AddItem MyRs("DSCCONTA")
            mvarMe.LstContas.ItemData(mvarMe.LstContas.NewIndex) = MyRs("IDCONTA")
            If mvarIDCONTA = MyRs("IDCONTA") Then
               mvarMe.LstContas.Checked(mvarMe.LstContas.NewIndex) = True
            End If
            MyRs.MoveNext
         Wend
         
      End If
   End With
   
   Call DefineDatas(mvarIDCONTA)
End Sub
Private Function VerificaCampos() As Boolean
   Dim i As Integer
   
   If mvarMe.CmbDTINI.Value > mvarMe.CmbDTFIM.Value Then
      Call ExibirAviso("Data início maior que a fim.", "Exportar")
      Exit Function
   End If
   For i = 0 To mvarMe.LstContas.ListCount
      If mvarMe.LstContas.Checked(i) Then
         Exit For
      End If
   Next
   VerificaCampos
End Function

Private Sub mvarMe_LstContasItemCheck(ByVal Item As Long)
   Dim i As Integer
   If mvarMe.LstContas.Checked(Item) Then
      For i = 0 To mvarMe.LstContas.ListCount - 1
         If i <> Item Then
            mvarMe.LstContas.Checked(i) = False
         End If
      Next
      nIDCONTA = mvarMe.LstContas.ItemData(Item)
      
      Call DefineDatas(nIDCONTA)
      mvarMe.CmdExportar.Enabled = False
      
      
      mvarMe.CmdExportar.Enabled = True
      mvarMe.CmbDTFIM.Enabled = True
      mvarMe.CmbDTINI.Enabled = True
      
   Else
      mvarMe.CmdExportar.Enabled = False
      mvarMe.CmbDTFIM.Enabled = False
      mvarMe.CmbDTINI.Enabled = False
   End If
End Sub
Private Sub DefineDatas(Optional pIDCONTA As Integer)
   Dim Sql As String
   Dim SqlDt As String
   Dim Rs  As Object
   Dim RsDt As Object
   Dim bOk As Boolean
   
   SqlDt = "Select isNull(Max(F2.DTVENCIMENTO), " & SqlDate("01/01/" & Year(Now())) & ")" & vbNewLine
   SqlDt = SqlDt & " From FLAN F2" & vbNewLine
   SqlDt = SqlDt & " Where F2.IDLOJA=" & mvarIDLOJA & vbNewLine
   SqlDt = SqlDt & " And F2.FLGEXPORT=1" & vbNewLine
   If pIDCONTA > 0 Then SqlDt = SqlDt & " And F2.IDCONTA=" & nIDCONTA & vbNewLine
   
   Sql = "Select Min(DTVENCIMENTO) [DTINI], Max(DTVENCIMENTO) [DTFIM]" & vbNewLine
   Sql = Sql & " From FLAN" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And (FLGEXPORT=0 Or FLGEXPORT is Null)" & vbNewLine
   If pIDCONTA > 0 Then Sql = Sql & " And IDCONTA=" & nIDCONTA & vbNewLine
   Sql = Sql & " And DTVENCIMENTO >= (" & SqlDt & ")"
   
   With mvarSys.xdb
      If .AbreTabela(Sql, Rs) Then
         If IsNull(Rs("DTINI")) Then
            bOk = False
            If .AbreTabela(SqlDt, RsDt) Then If Not IsNull(RsDt(0)) Then bOk = True
            If bOk Then
               mvarMe.CmbDTINI.Value = Format(RsDt(0), "dd/mm/yyyy")
            Else
               mvarMe.CmbDTINI.Value = "01/" & Format(mvarSys.xdb.SysDate(), "mm/yyyy")
            End If
         Else
            mvarMe.CmbDTINI.Value = Format(Rs("DTINI"), "dd/mm/yyyy")
         End If
         If IsNull(Rs("DTFIM")) Then
            mvarMe.CmbDTFIM.Value = Format(mvarSys.xdb.SysDate(), "dd/mm/yyyy")
         Else
            mvarMe.CmbDTFIM.Value = Format(Rs("DTFIM"), "dd/mm/yyyy")
         End If
      Else
         mvarMe.CmbDTINI.Value = "01/" & Format(mvarSys.xdb.SysDate(), "mm/yyyy")
         mvarMe.CmbDTFIM.Value = Format(mvarSys.xdb.SysDate(), "dd/mm/yyyy")
      End If
   End With
End Sub
