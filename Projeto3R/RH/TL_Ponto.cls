VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Ponto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mvarMe  As FrmPonto
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys            As Object
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   'mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmPonto
End Sub
Private Sub PopulaTela()
   Dim Sql As String
   Dim Rs  As Object
   
   Sql = ""

   Sql = "Select distinct B.IDLOJA, b.idbatida [ID1], b2.idbatida [ID2], F.NOME, B.DTBATIDA [Batida1], B2.DTBATIDA [Batida2]"
   Sql = Sql & ", (Select Max(T.HHFIM) From OATENDIMENTO T Where day(b.dtbatida)=day(t.dtatend) and month(b.dtbatida)=month(t.dtatend) and year(b.dtbatida)=year(t.dtatend))  as [ULTCLI]"
   Sql = Sql & " from rbatida b"
   Sql = Sql & " Left join rbatida b2 on b.idloja=b2.idloja and b.idfuncionario=b2.idfuncionario"
   Sql = Sql & " Left join rfuncionario f on b.idloja=f.idloja and b.idfuncionario=f.idfuncionario"
   Sql = Sql & " where f.nome like '%Rob%'"
   Sql = Sql & " and month(b.dtbatida) =8"
   Sql = Sql & " and year(b.dtbatida)= 2012"
   Sql = Sql & " and b.dtbatida=(Select min(a.dtbatida) from rbatida a where a.idfuncionario=b.idfuncionario"
   Sql = Sql & "                and day(a.dtbatida)=day(b.dtbatida)"
   Sql = Sql & "                and month(a.dtbatida)=month(b.dtbatida)"
   Sql = Sql & "                and year(a.dtbatida)=year(b.dtbatida)"
   Sql = Sql & "             )"
   Sql = Sql & " and b2.dtbatida=(Select max(a.dtbatida) from rbatida a where a.idfuncionario=b.idfuncionario"
   Sql = Sql & "                and day(a.dtbatida)=day(b.dtbatida)"
   Sql = Sql & "                and month(a.dtbatida)=month(b.dtbatida)"
   Sql = Sql & "                and year(a.dtbatida)=year(b.dtbatida)"
   Sql = Sql & "             )"
   Sql = Sql & " Order by B.DTBATIDA , B2.DTBATIDA "

   If mvarSys.XDb.AbreTabela(Sql, Rs) Then
      Call MontaTela
      Call PopulaMes(7, 2012)
      With mvarMe
         With .GrdPonto
         End With
      End With
   End If

End Sub
Private Sub PopulaMes(pMes As Integer, pAno As Integer)
   Dim Dia As Date
   Dia = CDate("01/" & pMes & "/" & pAno)
   
   With mvarMe.GrdPonto
      .Clear
      While Month(Dia) = pMes
         .AddRow Day(Dia)
         .CellValue(Day(Dia), .ColIndex("Dia")) = StrZero(Day(Dia), 2)
         .CellValue(Day(Dia), .ColIndex("Dia Semana")) = StrConv(WeekdayName(Weekday(Dia)), vbProperCase)
         
         Dia = DateAdd("d", 1, Dia)
      Wend
      .ColWidth("Dia") = 45
      .ColDefaultCell("Dia").eTextFlags = igTextCenter
   End With
End Sub
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
'   If Trim(mvarMe.TxtCampo01.Text) = "" Then
'      ExibirAviso "Descrição inválida."
'      bResult = False
'      mvarMe.TxtCampo01.SetFocus
'      GoTo Fim
'   End If
Fim:
   ValidaCampos = bResult
End Function
Private Sub LimpaTela()
   Dim n As Object
   On Error Resume Next
   With mvarMe
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
   End With
End Sub
Private Sub MontaTela()
   With mvarMe.GrdPonto
      .Clear True
      'IDLOJA , NOME, ENT, SAI, ULTCLI
      With .AddCol(sKey:="Batida1", sHeader:="Batida 1", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Batida1"
         .bSelected = False
      End With
      With .AddCol(sKey:="Batida2", sHeader:="Batida 2", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Batida2"
         .bSelected = False
      End With
      With .AddCol(sKey:="Dia", sHeader:="Dia", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Dia"
         .bSelected = False
      End With
      With .AddCol(sKey:="Dia Semana", sHeader:="Dia Semana", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Dia Semana"
         .bSelected = False
      End With
      With .AddCol(sKey:="Entrada", sHeader:="Entrada", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Entrada"
         .bSelected = False
      End With
      With .AddCol(sKey:="Saida", sHeader:="Saida", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Entrada"
         .bSelected = False
      End With
      With .AddCol(sKey:="HHTRAB", sHeader:="HH Trab.", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "HHTRAB"
         .bSelected = False
      End With
      With .AddCol(sKey:="HHESP", sHeader:="HH Esp.", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "HHESP"
         .bSelected = False
      End With
      With .AddCol(sKey:="Saldo", sHeader:="Saldo ", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Saldo"
         .bSelected = False
      End With
      With .AddCol(sKey:="Banco", sHeader:="Banco ", lWidth:=55, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "Banco"
         .bSelected = False
      End With
   End With
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   
End Sub
Private Function Salvar() As Boolean

End Function

Private Sub mvarMe_CmdRefreshClick()
   Call PopulaTela
End Sub

Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   If Salvar Then
   End If
End Sub
Private Sub mvarMe_Load()
   DoEvents
   'Call MontaTela
   'Call PopulaTela
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      Call MontaTela
      Call PopulaTela
   End If
End Sub


