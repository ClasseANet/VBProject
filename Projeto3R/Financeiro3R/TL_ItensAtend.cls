VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ITENSATEND"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents mvarMe     As FrmItensAtend
Attribute mvarMe.VB_VarHelpID = -1

Private mvarsys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDATEND  As Integer
Private mvarDeFaultIDVENDA  As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarsys = vData
   mvarIDLOJA = mvarsys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarsys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDATEND(ByVal vData As Integer)
   mvarIDATEND = vData
End Property
Public Property Get IDATEND() As Integer
   IDATEND = mvarIDATEND
End Property
Public Property Let DeFaultIDVENDA(ByVal vData As Integer)
   mvarDeFaultIDVENDA = vData
End Property
Public Property Get DeFaultIDVENDA() As Integer
   DeFaultIDVENDA = mvarDeFaultIDVENDA
End Property
Public Sub Show()
   'If mvarIDATEND = 0 Then
   '   Call ExibirInformacao("Atendimento inválido.")
   'Else
      mvarMe.Show vbModal
   'End If
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmItensAtend
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub PopulaTela()
   Dim Sql As String
   Dim MySel As Collection
   Dim i As Integer
      
   Sql = "Select I.IDSESSAO, S.DSCSERVICO [Serviço], T.DSCTRATAMENTO [Tratamento], A.DSCAREA [Área]" & vbNewLine
   'Sql = Sql + ", Right('000000' + Cast(I.IDVENDA as varchar), 6) [Venda]" & VBNEWLINE
   Sql = Sql + ", I.IDVENDA [Venda]" & vbNewLine
   Sql = Sql & ", I.IDTPSERVICO, I.IDTPTRATAMENTO, I.IDAREA" & vbNewLine
   Sql = Sql & " From OSESSAO I " & vbNewLine
   Sql = Sql & " Join OTPSERVICO S     On I.IDLOJA=S.IDLOJA And I.IDTPSERVICO=S.IDTPSERVICO" & vbNewLine
   Sql = Sql & " Left Join OTPTRATAMENTO T  On I.IDLOJA=T.IDLOJA And I.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
   Sql = Sql & " Left Join OAREA A          On I.IDLOJA=A.IDLOJA And I.IDAREA=A.IDAREA" & vbNewLine
   Sql = Sql & " Where I.IDLOJA = " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And I.IDATENDIMENTO= " & mvarIDATEND
   If mvarIDATEND = 0 Then
      Sql = Sql & " And (I.IDVENDA= " & mvarDeFaultIDVENDA
      Sql = Sql & " or I.IDVENDA= 0"
      Sql = Sql & " or I.IDVENDA is Null)"
   End If
   'Sql = Sql & " And I.IDTPSERVICO<>1"
      
   bExisteItemVenda = False
   If mvarsys.xDb.AbreTabela(Sql, MyRs) Then
      mvarMe.GrdItens.FillFromRS MyRs
      
      For i = 1 To mvarMe.GrdItens.ColCount
         Call mvarMe.GrdItens.AutoWidthCol(i)
         mvarMe.GrdItens.ColWidth(i) = mvarMe.GrdItens.ColWidth(i) + 5
      Next
      mvarMe.GrdItens.ColWidth("Venda") = 80
      
      For i = 1 To mvarMe.GrdItens.RowCount
         If mvarMe.GrdItens.CellValue(1, "IDTPSERVICO") <> 1 Then
            bExisteItemVenda = True
            Exit For
         End If
      Next
   End If
   
   If bExisteItemVenda Then
      mvarMe.CmdOk.Enabled = True
   Else
      mvarMe.CmdOk.Enabled = False
      Call ExibirInformacao("Atendimento não possui itens de venda.")
      Unload mvarMe
   End If
End Sub

Private Sub mvarMe_Active()
   Call MontaTela
   Call PopulaTela
   
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      '* Definir foco
      With mvarMe.GrdItens
         If .Visible And .Enabled Then
            .SetFocus
         End If
      End With
      Call SetTag(mvarMe, "1VEZ", 0)
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   Call SalvarItens
   Unload mvarMe
End Sub
Private Sub SalvarItens()
   Dim Sql As String
   Dim i As Integer
   Dim Queries As Collection
   
   Set Queries = New Collection
'   Sql = "Update OSESSAO "
'   Sql = Sql & " Set IDVENDA= Null"
'   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'   Sql = Sql & " And IDATENDIMENTO=" & SqlNum(nIDATEND)
'   Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
   'Call mvarSys.xDb.Executa(Sql)
   For i = 1 To mvarMe.GrdItens.RowCount
      Sql = "Update OSESSAO "
      Sql = Sql & " Set IDVENDA= " & SqlNum(xVal(mvarMe.GrdItens.CellValue(i, "VENDA")))
      Sql = Sql & ", ALTERSTAMP=1"
      Sql = Sql & ", TIMESTAMP = GetDate()"
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And IDATENDIMENTO=" & SqlNum(mvarIDATEND)
      Sql = Sql & " And IDSESSAO=" & SqlNum(mvarMe.GrdItens.CellValue(i, "IDSESSAO"))
      Queries.Add Sql
   Next
   Call mvarsys.xDb.Executa(Queries)
End Sub

Private Sub mvarMe_GrdItensDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
'   DoEvents
   If mvarDeFaultIDVENDA = 0 Then
      mvarMe.GrdItens.CellValue(lRow, "Venda") = 0
   Else
      mvarMe.GrdItens.CellValue(lRow, "Venda") = mvarDeFaultIDVENDA
   End If
End Sub

Private Sub mvarMe_GrdItensMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   If Button = 2 Then
      mvarMe.GrdItens.CellValue(lRow, "Venda") = 0
   End If
End Sub

Private Sub mvarMe_Load()
   mvarMe.CmdOk.Enabled = False
   'Call MotaTela
End Sub
Private Sub MontaTela()
   With mvarMe.GrdItens
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      .BorderStyle = igBorderThin
      .AutoWidthColWithHeader = True
      .RowMode = True
      '.FocusRect = False
       
     
      With .Combos
         With .Add("Venda")
            Sql = "Select AV.IDVENDA"
            Sql = Sql & " From OATENDIMENTO_VENDA AV"
            Sql = Sql & " Where AV.IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And AV.IDATENDIMENTO=" & mvarIDATEND
            If mvarIDATEND = 0 Then
               Sql = Sql & " And AV.IDVENDA=" & mvarDeFaultIDVENDA
            End If
            If mvarsys.xDb.AbreTabela(Sql) Then
               .AddItem Space(6), "0"
               While Not mvarsys.xDb.RsAux.EOF
                  sAux = mvarsys.xDb.RsAux("IDVENDA")
                  .AddItem Right("000000" & sAux, 6), sAux
                  
                  mvarsys.xDb.RsAux.MoveNext
               Wend
            End If
            Set TbProd = Nothing
            .AutoAdjustWidth
         End With
      End With
      With .AddCol(sKey:="IDSESSAO", sHeader:="#", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "IDSESSAO"
      End With
      With .AddCol(sKey:="Serviço", sHeader:="Serviço", lWidth:=80, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "Serviço"
          
      End With
      With .AddCol(sKey:="Tratamento", sHeader:="Tratamento", lWidth:=80, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "Tratamento"
      End With
      With .AddCol(sKey:="Área", sHeader:="Área", lWidth:=80, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "Área"
      End With
      With .AddCol(sKey:="Venda", sHeader:="Venda", lWidth:=40, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .eTextFlags = igTextLeft 'Or igTextNoClip
         .sCtrlKey = "Venda"
      End With
      With .AddCol(sKey:="IDTPSERVICO", sHeader:="IDTPSERVICO", lWidth:=40, bVisible:=False)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft 'Or igTextNoClip
         .sCtrlKey = "IDTPSERVICO"
      End With
      With .AddCol(sKey:="IDTPTRATAMENTO", sHeader:="IDTPTRATAMENTO", lWidth:=40, bVisible:=False)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft 'Or igTextNoClip
         .sCtrlKey = "IDTPTRATAMENTO"
      End With
      With .AddCol(sKey:="IDAREA", sHeader:="IDAREA", lWidth:=40, bVisible:=False)
         .eType = igCellCombo
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft 'Or igTextNoClip
         .sCtrlKey = "IDAREA"
      End With
      .ColHeaderTextFlags("IDSESSAO") = igTextCenter
      .ColHeaderTextFlags("Venda") = igTextCenter
   End With
End Sub
