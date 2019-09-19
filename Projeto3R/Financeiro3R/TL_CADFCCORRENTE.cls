VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADFCCORRENTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents mvarMe     As FrmCADFCCORRENTE
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDCONTA  As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   mvarIDCONTA = mvarSys.Propriedades("IDCONTA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCONTA(ByVal vData As Integer)
   mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Integer
   IDCONTA = mvarIDCONTA
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADFCCORRENTE
   mvarIDLOJA = 1
End Sub
Public Sub Show()
   If mvarIDCONTA = 0 Then Call ExibirAviso("Não existe conta corrente definida.")
   If Reload Then
      Set Sys = Sys
      Unload mvarMe
   End If
   
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
      
   '* Definir foco
   With mvarMe.TxtDSCCONTA
      If .Visible And .Enabled Then
         .SetFocus
      End If
   End With

End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
'   Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdSalvarClick()
   Call Salvar
End Sub
Private Sub mvarMe_Load()
   Call MontaTela
   Call PopulaTela
End Sub
Private Sub mvarMe_Rezise()
  With mvarMe
     .SccConta.Move 0, 50, mvarMe.ScaleWidth - 50, 420
  End With
End Sub
Private Sub MontaTela()
   Call LimpaTela
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
      .SccConta.VisualTheme = xtpShortcutThemeOffice2003
      .SccConta.GradientHorizontal = False
   End With
End Sub
Private Sub LimpaTela()
   On Error Resume Next
   For Each n In mvarMe.Controls
      n.Text = ""
   Next
End Sub
Public Sub PopulaTela()
   Dim TbConta As Object
   Set TbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   With TbConta
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA) Then
         mvarMe.SccConta.Caption = "Conta: " & .DSCCONTA
         mvarMe.TxtIDCONTA.Text = .IDCONTA
         mvarMe.TxtDSCCONTA.Text = .DSCCONTA
         mvarMe.TxtNUMBANCO.Text = .NUMBANCO
         mvarMe.TxtNUMAGENCIA.Text = .NUMAGENCIA
         mvarMe.TxtNUMCONTA.Text = .NUMCONTA
         mvarMe.TxtDVCONTA.Text = .DVCONTA
         mvarMe.ChkEVENDA.Value = (.EVENDA = 1)
         mvarMe.ChkATIVO.Value = (.ATIVO = 1)
         If .TPCONTA = "D" Then
            mvarMe.OptTPCONTA(1).Value = True
         ElseIf .TPCONTA = "B" Then
            mvarMe.OptTPCONTA(0).Value = True
         End If
      End If
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   Dim Queries As Collection
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      'bIsDirt = TbCad.isDirt
      'If bIsDirt Then
      '   If TbCad.Salvar() Then
      '      bOk = True
      '   End If
      'Else
      '   bOk = True
      'End If
      bOk = False
      If Queries.Count = 0 Then
         bOk = True
      Else
         If mvarSys.xDb.Executa(Queries) Then
            bOk = True
         End If
      End If
      
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtDSCCONTA.Text) = "" Then
      ExibirAviso "Descrição inválida."
      bResult = False
      mvarMe.TxtDSCCONTA.SetFocus
      GoTo Fim
   End If
   If Not (mvarMe.OptTPCONTA(0).Value Or mvarMe.OptTPCONTA(1).Value) Then
      ExibirAviso "Escolha o tipo da conta."
      bResult = False
      mvarMe.OptTPCONTA(0).SetFocus
      GoTo Fim
   End If
   
Fim:
   ValidaCampos = bResult
End Function
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim TbConta As Object
   Set TbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")

   With TbConta
      Set .xDb = mvarSys.xDb
      
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCONTA:=mvarIDCONTA)
      .DSCCONTA = mvarMe.TxtDSCCONTA.Text
      .NUMBANCO = mvarMe.TxtNUMBANCO.Text
      .NUMAGENCIA = mvarMe.TxtNUMAGENCIA.Text
      .NUMCONTA = mvarMe.TxtNUMCONTA.Text
      .DVCONTA = mvarMe.TxtDVCONTA.Text
      .EVENDA = IIf(mvarMe.ChkEVENDA.Value, 1, 0)
      If mvarMe.OptTPCONTA(0).Value Then
         .TPCONTA = "B"
      ElseIf mvarMe.OptTPCONTA(1).Value Then
         .TPCONTA = "D"
      End If
      .ATIVO = IIf(mvarMe.ChkATIVO.Value, 1, 0)
      
      If .isDirt Then
         If .Salvar Then
            mvarIDCONTA = .IDCONTA
            Call PopulaTela
            Call ExibirResultado(mvarSys, True, 1)
         End If
      End If
   End With
End Sub

