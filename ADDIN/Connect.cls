VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = " VB EditorUtil"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ADDINS"
Option Explicit


Const guidMYTOOL$ = "{B7AFC8D0-EBE5-11cf-A497-00A0C911E8B0}"

Implements IDTExtensibility
Public WithEvents Menu1 As CommandBarEvents          'command bar event handler
Attribute Menu1.VB_VarHelpID = -1
Public WithEvents Menu2 As CommandBarEvents          'command bar event handler
Attribute Menu2.VB_VarHelpID = -1
Public WithEvents Menu3 As CommandBarEvents          'command bar event handler
Attribute Menu3.VB_VarHelpID = -1
Public WithEvents Menu4 As CommandBarEvents          'command bar event handler
Attribute Menu4.VB_VarHelpID = -1
Public WithEvents Menu5 As CommandBarEvents          'command bar event handler
Attribute Menu5.VB_VarHelpID = -1
Public WithEvents Menu6 As CommandBarEvents          'command bar event handler
Attribute Menu6.VB_VarHelpID = -1
Public WithEvents Menu7 As CommandBarEvents          'command bar event handler
Attribute Menu7.VB_VarHelpID = -1
Public WithEvents Menu8 As CommandBarEvents          'command bar event handler
Attribute Menu8.VB_VarHelpID = -1
Public WithEvents Menu9 As CommandBarEvents          'command bar event handler
Attribute Menu9.VB_VarHelpID = -1
Public WithEvents Menu10 As CommandBarEvents         'command bar event handler
Attribute Menu10.VB_VarHelpID = -1
Public WithEvents Menu11 As CommandBarEvents         'command bar event handler
Attribute Menu11.VB_VarHelpID = -1
Public WithEvents Menu12 As CommandBarEvents         'command bar event handler
Attribute Menu12.VB_VarHelpID = -1
Public WithEvents Menu13 As CommandBarEvents         'command bar event handler
Attribute Menu13.VB_VarHelpID = -1

Public WithEvents Mn1 As CommandBarEvents          'command bar event handler
Attribute Mn1.VB_VarHelpID = -1
Public WithEvents Mn2 As CommandBarEvents          'command bar event handler
Attribute Mn2.VB_VarHelpID = -1
Public WithEvents Mn3 As CommandBarEvents          'command bar event handler
Attribute Mn3.VB_VarHelpID = -1
Public WithEvents Mn4 As CommandBarEvents          'command bar event handler
Attribute Mn4.VB_VarHelpID = -1
Public WithEvents Mn5 As CommandBarEvents          'command bar event handler
Attribute Mn5.VB_VarHelpID = -1
Public WithEvents Mn6 As CommandBarEvents          'command bar event handler
Attribute Mn6.VB_VarHelpID = -1
Public WithEvents Mn7 As CommandBarEvents          'command bar event handler
Attribute Mn7.VB_VarHelpID = -1
Public WithEvents Mn8 As CommandBarEvents          'command bar event handler
Attribute Mn8.VB_VarHelpID = -1
Public WithEvents Mn9 As CommandBarEvents          'command bar event handler
Attribute Mn9.VB_VarHelpID = -1
Public WithEvents Mn10 As CommandBarEvents         'command bar event handler
Attribute Mn10.VB_VarHelpID = -1
Public WithEvents Mn11 As CommandBarEvents         'command bar event handler
Attribute Mn11.VB_VarHelpID = -1
Public WithEvents Mn12 As CommandBarEvents         'command bar event handler
Attribute Mn12.VB_VarHelpID = -1
Public WithEvents Mn13 As CommandBarEvents         'command bar event handler
Attribute Mn13.VB_VarHelpID = -1

Public WithEvents PrjHandler  As VBProjectsEvents          'projects event handler
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler  As VBComponentsEvents        'components event handler
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler  As VBControlsEvents          'controls event handler
Attribute CtlHandler.VB_VarHelpID = -1

Public WithEvents FileCtlHnd As FileControlEvents
Attribute FileCtlHnd.VB_VarHelpID = -1
Public WithEvents RefHnd As ReferencesEvents
Attribute RefHnd.VB_VarHelpID = -1
Public WithEvents SelCtlHnd As SelectedVBControlsEvents
Attribute SelCtlHnd.VB_VarHelpID = -1

'Public VBInstance             As VBIDE.VBE
Public CmdBarControl          As Office.CommandBarControl
Public FormDisplayed          As Boolean

Private mFrmAddIn             As New FrmAddIn
'Private mFrmMsgBox            As New FrmMsgBox
Private mFrmProp              As New FrmPropAddin
Public NonModalApp            As Boolean
Private IndBar                As Integer

Private ActiveProject         As VBProject
Private ActiveComponent       As VBComponent
Private ActiveMember          As MEMBER
Private ActiveCode            As CodeModule

Private mvarConstru            As New CONSTRUTOR
Private mvarNameBar           As String

Public Property Get NameBar() As String
   NameBar = mvarNameBar
End Property
Public Property Let NameBar(vData As String)
   mvarNameBar = vData
End Property
Public Property Get Constru() As CONSTRUTOR
   If mvarConstru Is Nothing Then
      Set mvarConstru = New CONSTRUTOR
   End If
   Set Constru = mvarConstru
End Property
Public Property Set Constru(vData As CONSTRUTOR)
   Set mvarConstru = vData
End Property



Private Sub Class_Initialize()
   'Set ActiveComponent = VBInstance.ActiveVBProject.VBComponents
   NonModalApp = True
End Sub

Private Sub Class_Terminate()
   Set mvarConstru = Nothing
End Sub
Sub Hide()
   On Error Resume Next
   'mvarConstru.Hide
   FormDisplayed = False
   '   mFrmAddIn.Hide
   Set mvarConstru = Nothing
   Unload mFrmAddIn
End Sub
Sub Show(Optional Index = 1)

   On Error GoTo Fim
   Screen.MousePointer = vbHourglass

   Set glbProj = VBInstance.ActiveVBProject

   Select Case Index
      Case 1 '* Construtor
         'Set mvarConstru.ClsConnect = Me
         Set Constru = Nothing
         Set Constru.VbProj = VBInstance.ActiveVBProject
         Set mvarConstru.ClsConnect = Me
         Me.FormDisplayed = True
         mvarConstru.Show
      Case 2 '* Propriedades
         mFrmProp.Show vbModal
         
      Case 3 '* TabOrder
           gWinTabOrder.Visible = True
           gDocTabOrder.RefreshList 3
   End Select
   Screen.MousePointer = vbDefault
   Exit Sub
Fim:
   Screen.MousePointer = vbDefault
   If InArray(Err, Array(457)) Then
      '* This key is already associated with an element of this collection
      Resume Next
   End If
End Sub
Private Sub CmpHandler_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
   On Error GoTo CmpHandler_ItemActivatedErr
   
   Set ActiveComponent = VBComponent
   If gWinTabOrder.Visible Then
      gDocTabOrder.RefreshList 0
   End If

CmpHandler_ItemActivatedErr:
   
End Sub

Private Sub CmpHandler_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
   CmpHandler_ItemActivated VBComponent
End Sub
Private Sub CtlHandler_ItemAdded(ByVal VBControl As VBIDE.VBControl)
  If gWinTabOrder.Visible Then
    gDocTabOrder.ControlAdded VBControl
  End If
End Sub
Private Sub CtlHandler_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
  If gWinTabOrder.Visible Then
    gDocTabOrder.ControlRemoved VBControl
  End If
End Sub
Private Sub CtlHandler_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
  If gWinTabOrder.Visible Then
    gDocTabOrder.ControlRenamed VBControl, OldName, OldIndex
  End If
End Sub
Public Sub ConnectAddIn(CurrentApp As VBIDE.Application)
    Set VbApplication = CurrentApp
End Sub

Private Sub FileCtlHnd_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
'   MsgBox "FileCtlHnd_AfterAddFile"
End Sub

Private Sub FileCtlHnd_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)
'   MsgBox "FileCtlHnd_BeforeLoadFile"
End Sub

Private Sub FileCtlHnd_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
'   MsgBox "FileCtlHnd_DoGetNewFileName"
End Sub

'************************************************
'* Este Método Adiciona o "Add-In" para o VB    *
'************************************************
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
'   On Error GoTo Fim

   Dim VBI As VBIDE.AddIn


   '* Salvar instância do VB
   Set VBInstance = VBInst
 
   Call GetConfig
   
   If Not gWinTabOrder Is Nothing Then
     '* Apenas mostrar Tela pois aplicação já está rodando
     Call Show(3)
     If ConnectMode = vbext_cm_AfterStartup Then
        '* Iniciado do "addin manager"
         Call AddCommandBar
      End If
      Exit Sub
   End If

   '* Criar Janela da aplicação
   If ConnectMode = vbext_cm_External Then
      '* Verificar se Aplicação já estar rodando
      On Error Resume Next
      Set VBI = VBInstance.Addins("VbEditorUtil.Connect")
      On Error GoTo Fim
      If VBI Is Nothing Then
         '* App is not in the VBADDIN.INI file so it is not in the collection
         '* so lets attempt to use the 1st addin in the collection just
         '* to get this app running and if there are none, an error
         '* will occur and this app will not run
         Set gWinTabOrder = VBInstance.Windows.CreateToolWindow(VBInstance.Addins(1), "VbEditorUtil.docTabOrder", LoadResString(200), guidMYTOOL$, gDocTabOrder)
         'Set gWindow = FrmAddIn
      Else
         If VBI.Connect = False Then
            Set gWinTabOrder = VBInstance.Windows.CreateToolWindow(VBI, "VbEditorUtil.docTabOrder", LoadResString(200), guidMYTOOL$, gDocTabOrder)
            'Set gWindow = FrmAddIn
         End If
      End If
   Else
      '* Deve estar sendo executado do "addin mgr"
      Set gWinTabOrder = VBInstance.Windows.CreateToolWindow(AddInInst, "VbEditorUtil.docTabOrder", LoadResString(200), guidMYTOOL$, gDocTabOrder)
      'Set gWindow = FrmAddIn
   End If

   'sink the project, components and controls event handler
   Set PrjHandler = VBInstance.Events.VBProjectsEvents
   Set CmpHandler = VBInstance.Events.VBComponentsEvents(Nothing)
   Set CtlHandler = VBInstance.Events.VBControlsEvents(Nothing, Nothing)

   If ConnectMode = vbext_cm_External Then
      '* Executado "addin toolbar"
      Call Show(3)
   ElseIf ConnectMode = vbext_cm_AfterStartup Then
      '* Executado de "addin manager"
      Call AddCommandBar
   ElseIf ConnectMode = vbext_cm_Startup Then
      '* Executado ao abrir uma instância do VB
      Call AddCommandBar
   End If
   Exit Sub
Fim:
   Call ShowError
End Sub
'**************************************
'* Este método remove o Add-In do VB  *
'**************************************
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
   On Error Resume Next
   Dim CmdBarPPEdit As CommandBarPopup
   Dim CmdBarPP As CommandBarPopup
   
   Set CmdBarPPEdit = VBInstance.CommandBars("Menu Bar").Controls("Edit")
   Set CmdBarPP = CmdBarPPEdit.Controls("VB EditorUtil")
   CmdBarPP.Delete
   Set CmdBarPP = Nothing

   
   '* Exclui a Barra de Tarefa.
   If RemoveMode = vbext_dm_UserClosed Then
      Call SaveSetting(App.Title, "Settings", "ShowCmdBar", True)
   ElseIf RemoveMode = vbext_dm_HostShutdown Then
      Call SaveSetting(App.Title, "Settings", "ShowCmdBar", VBInstance.CommandBars(mvarNameBar).Visible)
   End If
   
   If VBInstance.CommandBars(mvarNameBar).Visible Then
      VBInstance.CommandBars(mvarNameBar).Delete
   End If
   '* Exclui o botão da barra de comandos
   '* CmdBarControl.Delete

   '* Fecha a aplicação
   If FormDisplayed Then
      SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
      FormDisplayed = False
   Else
      SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
   End If
   If gWinTabOrder.Visible Then
'      SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "1"
   Else
'      SaveSetting APP_CATEGORY, App.Title, "DisplayOnConnect", "0"
   End If


   Set mvarConstru = Nothing
   
   Set gWinTabOrder = Nothing
   
   Unload mFrmAddIn
   Set mFrmAddIn = Nothing
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
   If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
      '* set this to display the form on connect
      Me.Show
   End If
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
   Dim x
   x = x
End Sub
Public Sub AddCommandBar()

On Error GoTo Fim:

   '* Adiconar se não existir os butões do Add-ins.
  ' mvarNameBar = "Edit"
'   Set CmdBarControl = AddControlBar("Edit", "Indentar Função", "BMP_INDENTAR")
   Dim CmdBarCtl          As Office.CommandBarControl
   Dim CmdBarPPEdit As CommandBarPopup
   Dim CmdBarPP As CommandBarPopup
   Dim i%, n As Variant


   mvarNameBar = "VbEditorUtil"
   IndBar = 2
   IndBar = VBInstance.CommandBars.Count + 1
   
   VBInstance.CommandBars(mvarNameBar).Delete
   VBInstance.CommandBars.Add mvarNameBar, 1, , True
'   VBInstance.CommandBars(IndBar).Visible = True '* Tornar o a brra visível.
   
   Set CmdBarControl = AddControlBar(mvarNameBar, "Indentar Função", "BMP_INDENTAR")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Indentar Módulo", "BMP_INDENTAR_MODULO")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Indentar Projeto", "BMP_INDENTAR_PROJETO")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Cabeçalho da Função", "BMP_CAB_FUNCAO", True)
   Set CmdBarControl = AddControlBar(mvarNameBar, "Cabeçalho do Módulo", "BMP_CAB_MODULO")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Cabeçalho em todas Funções", "BMP_CAB_TUDO")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Caixa de Mesagem", "BMP_MSGBOX", True)
   Set CmdBarControl = AddControlBar(mvarNameBar, "Tratamento de Erro na Função", "BMP_ERR_FUNCAO", True)
   Set CmdBarControl = AddControlBar(mvarNameBar, "Tratamento de Erro em todas as Funções", "BMP_ERR_TUDO")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Ordenar Controles", "BMP_TABORDER", True)
   Set CmdBarControl = AddControlBar(mvarNameBar, "Construtor", "BMP_CONSTRUTOR", True)
   Set CmdBarControl = AddControlBar(mvarNameBar, "Biblioteca de Funções", "BMP_CD")
   Set CmdBarControl = AddControlBar(mvarNameBar, "Propriedades", "BMP_PROPRIEDADE", True)
      
   VBInstance.CommandBars(mvarNameBar).Visible = GetSetting(App.Title, "Settings", "ShowCmdBar", True)
   
   On Error GoTo LimpaEdit

   '* Exclui menu existente
   Set CmdBarPPEdit = VBInstance.CommandBars("Menu Bar").Controls("Edit")
   Set CmdBarPP = CmdBarPPEdit.Controls("VB EditorUtil")
   CmdBarPP.Delete
   Set CmdBarPP = Nothing
   
   
   '* Cria novo Menu
   Set CmdBarPP = CmdBarPPEdit.Controls.Add(msoControlPopup, , , 14)
   CmdBarPP.Caption = "VB EditorUtil"
   CmdBarPP.Visible = True
   CmdBarPP.BeginGroup = True
   'CmdBarPP.CommandBar.ShowPopup
   '* Monta os controles do novo menu
   For Each n In VBInstance.CommandBars(mvarNameBar).Controls
      Set CmdBarControl = AddControlBar("", n.Caption, n.Tag, n.BeginGroup, CmdBarPP)
   Next

   'restore the last state
   If GetSetting(APP_CATEGORY, App.Title, "DisplayOnConnect", "0") = "1" Then
      'set this to display the form on connect
      Me.Show
   End If
'RefreshWindows
'Call SynchCodePaneScroll

   Exit Sub

LimpaEdit:
   'VBInstance.CommandBars("Menu Bar").Controls("Edit").Reset
   Resume Next
Fim:
   If Err = 5 Then '* Invalid procedure call or argument
'      IndBar = VBInstance.CommandBars(mvarNameBar).Index
      Resume Next
   ElseIf Err = 13 Then
      Resume Next
   Else
      Call ShowError
   End If
End Sub
Public Function AddControlBar(NameBar$, Caption$, Optional Icone, Optional Separador = False, Optional CmdBar) As Office.CommandBarControl
   
   Dim lBarControl As Office.CommandBarControl
   Dim EvtMenu As CommandBarEvents
   Dim MenuBool As Boolean
   
   If IsMissing(CmdBar) Then
      Set CmdBar = VBInstance.CommandBars(NameBar$)
      MenuBool = True
   End If
   Set lBarControl = CmdBar.Controls.Add(msoControlButton, , 3)
   

   lBarControl.BeginGroup = Separador
   lBarControl.Caption = Caption$

   If Not IsMissing(Icone) Then
      lBarControl.Tag = Icone
      Clipboard.SetData LoadResPicture(Icone, vbResBitmap)
      lBarControl.PasteFace
      Clipboard.Clear
   End If
   If MenuBool Then
      Select Case lBarControl.Index
         Case 1:  Set Me.Menu1 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 2:  Set Me.Menu2 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 3:  Set Me.Menu3 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 4:  Set Me.Menu4 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 5:  Set Me.Menu5 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 6:  Set Me.Menu6 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 7:  Set Me.Menu7 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 8:  Set Me.Menu8 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 9:  Set Me.Menu9 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 10: Set Me.Menu10 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 11: Set Me.Menu11 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 12: Set Me.Menu12 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 13: Set Me.Menu13 = VBInstance.Events.CommandBarEvents(lBarControl)
      End Select
   Else
      Select Case lBarControl.Index
         Case 1:  Set Me.Mn1 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 2:  Set Me.Mn2 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 3:  Set Me.Mn3 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 4:  Set Me.Mn4 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 5:  Set Me.Mn5 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 6:  Set Me.Mn6 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 7:  Set Me.Mn7 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 8:  Set Me.Mn8 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 9:  Set Me.Mn9 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 10: Set Me.Mn10 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 11: Set Me.Mn11 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 12: Set Me.Mn12 = VBInstance.Events.CommandBarEvents(lBarControl)
         Case 13: Set Me.Mn13 = VBInstance.Events.CommandBarEvents(lBarControl)
      End Select
   End If
   Set EvtMenu = VBInstance.Events.CommandBarEvents(lBarControl)
   Set AddControlBar = lBarControl
End Function

Private Sub Mn1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   Call Menu1_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu2_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu3_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn4_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu4_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn5_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu5_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn6_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu6_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn7_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu7_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu8_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn9_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu9_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn10_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu10_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn11_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu11_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn12_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu12_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub Mn13_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Call Menu13_Click(CommandBarControl, handled, CancelDefault)
End Sub
Private Sub PrjHandler_ItemActivated(ByVal VBProject As VBIDE.VBProject)
'   MsgBox "PrjHandler_ItemActivated"
End Sub

Private Sub PrjHandler_ItemAdded(ByVal VBProject As VBIDE.VBProject)
   Set mvarConstru.ClsConnect = Me
   Set mvarConstru.VbProj = VBProject
   'mvarConstru.Show
End Sub

Private Sub PrjHandler_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
  'this takes care of the user removing the only project
'  MsgBox "PrjHandler_ItemRemoved"
'  If gWinTabOrder.Visible Then
'    gDocTabOrder.RefreshList 0
'  End If
End Sub
Private Sub PrjHandler_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
'   MsgBox "PrjHandler_ItemRenamed"
End Sub
Private Sub Menu1_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'* Indentar Função
   Dim Atu_TopLine As Integer
   Dim StartLine&, StartColumn&, EndLine&, EndColumn&
'   On Error Resume Next
   Atu_TopLine = VBInstance.ActiveCodePane.TopLine
   Call VBInstance.ActiveCodePane.GetSelection(StartLine&, StartColumn&, EndLine&, EndColumn&)
   
   Call IndentarFuncao
   
   VBInstance.ActiveCodePane.TopLine = Atu_TopLine
   Call VBInstance.ActiveCodePane.SetSelection(StartLine&, StartColumn&, EndLine&, EndColumn&)
   On Error GoTo 0
End Sub
Private Sub Menu2_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'* Indentar Módulo
   Dim MyMember As MEMBER, n As Variant
   Dim Total%, Vl%, Str$

   Dim Atu_TopLine As Integer
   Dim StartLine&, StartColumn&, EndLine&, EndColumn&
   Dim LinhaStart As Long
   On Error Resume Next
   If VBInstance.ActiveCodePane Is Nothing Then
      Atu_TopLine = VBInstance.ActiveCodePane.TopLine
   End If
   Atu_TopLine = VBInstance.ActiveCodePane.TopLine
   Call VBInstance.ActiveCodePane.GetSelection(StartLine&, StartColumn&, EndLine&, EndColumn&)
   
   
   Vl = 0
   Total = VBInstance.ActiveCodePane.CodeModule.Members.Count
   For Each n In VBInstance.ActiveCodePane.CodeModule.Members
      Set MyMember = n
      If MyMember.Type <> vbext_mt_Const And MyMember.Type <> vbext_mt_Variable Then
         Str = "Processando... " & MyMember.Name
      End If
      Vl = Vl + 1
      If Not AtuFlood(Vl, Total, Str) Then Exit For

      If MyMember.Type <> vbext_mt_Const And MyMember.Type <> vbext_mt_Variable Then
         LinhaStart = VBInstance.ActiveCodePane.CodeModule.ProcStartLine(MyMember.Name, vbext_ProcKind.vbext_pk_Proc)
         If LinhaStart > VBInstance.ActiveCodePane.CodeModule.CountOfDeclarationLines Then
            Call IndentarFuncao(, MyMember)
         End If
      End If
      Set MyMember = Nothing
   Next
   If LinhaStart > 0 Then
      Call VBInstance.ActiveCodePane.SetSelection(1, 1, 1, 1)
      Call IndentarFuncao(, , , , , , , , False)
   End If
   
   
   VBInstance.ActiveCodePane.TopLine = Atu_TopLine
   Call VBInstance.ActiveCodePane.SetSelection(StartLine&, StartColumn&, EndLine&, EndColumn&)
   
   Call FimFlood
Fim:
End Sub
Private Sub Menu3_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'* Indentar Projeto
   Dim MyComp As VBComponent, MyMember As MEMBER
   Dim n As Variant, n1 As Variant
   Dim Total%, Vl%
   Dim Total2%, Vl2%
   Dim Atu_TopLine As Integer
   Dim StartLine&, StartColumn&, EndLine&, EndColumn&
   Dim LinhaStart As Long
   Dim i%, StrWin$
   
   
On Error Resume Next
   Atu_TopLine = VBInstance.ActiveCodePane.TopLine
   Call VBInstance.ActiveCodePane.GetSelection(StartLine&, StartColumn&, EndLine&, EndColumn&)
   StrWin$ = "|"
   For i = 1 To VBInstance.Windows.Count
      If VBInstance.Windows(i).Type = 0 Then
         StrWin$ = StrWin$ & VBInstance.Windows(i).Caption & "|"
      End If
   Next
   Total% = VBInstance.ActiveVBProject.VBComponents.Count
   For Each n In VBInstance.ActiveVBProject.VBComponents
      Set MyComp = n
      Vl% = Vl% + 1
      If MyComp.Type <> 4 Then   '* ResFile
         
         Total2 = MyComp.CodeModule.Members.Count
         Vl2 = 0
         For Each n1 In MyComp.CodeModule.Members
            Set MyMember = n1
            Vl2 = Vl2 + 1
            If Not AtuFlood(Vl2, Total2, "Componente " & Trim(CStr(Vl)) & " / " & Trim(CStr(Total)) & " (" & MyComp.Name & ")") Then
               GoTo Fim
            End If
            If MyMember.Type <> vbext_mt_Const And MyMember.Type <> vbext_mt_Variable Then
               LinhaStart = MyComp.CodeModule.CodePane.CodeModule.ProcStartLine(MyMember.Name, vbext_ProcKind.vbext_pk_Proc)
               If LinhaStart > MyComp.CodeModule.CountOfDeclarationLines Then
                  If IndentarFuncao(MyComp, MyMember) = "End" Then GoTo Fim
               End If
            End If
            Set MyMember = Nothing
         Next
         If LinhaStart > 0 Then
            Call MyComp.CodeModule.CodePane.SetSelection(1, 1, 1, 1)
            Call IndentarFuncao(, , , , , , , , False)
         End If
         
      End If
      For i = 1 To VBInstance.Windows.Count - 1
         If VBInstance.Windows(i).Type = 0 Then
            If InStr(StrWin$, "|" & VBInstance.Windows(i).Caption & "|") = 0 Then
               VBInstance.Windows(i).Close
            End If
         End If
      Next
      Set MyComp = Nothing
   Next
   

   VBInstance.ActiveCodePane.TopLine = Atu_TopLine
   Call VBInstance.ActiveCodePane.SetSelection(StartLine&, StartColumn&, EndLine&, EndColumn&)
   
   Call FimFlood
Exit Sub
Fim:
   Call FimFlood
   If Err <> 0 Then MsgBox CStr(Err) & "-" & Error
End Sub

Private Sub Menu4_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Inserir Cabeçalho na Função
   Call InserirCabFuncao
End Sub
Private Sub Menu5_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Inserir Cabeçalho no Módulo
   Call InserirCabModulo
End Sub
Private Sub Menu6_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   On Error GoTo Fim
   '* Inserir Cabeçalho no Módulo e em todas as suas funções
   Dim n As Variant
   With VBInstance.ActiveCodePane.CodeModule
      For Each n In .Members
         If n.Type = vbext_mt_Method Then
            Call InserirCabFuncao(n.Name)
         End If
      Next
   End With
   Call InserirCabModulo
Fim:
End Sub
Private Sub Menu7_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Montador Auxiliar de MsgBox
   MsgBox "Auxuliar de Caixa de Mensagem (MsgBox) não implementado.", vbInformation
   'Me.Show
End Sub
Private Sub Menu8_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Inserir Tratamento de Erro na Função
   Call InserirErrFuncao
End Sub
Private Sub Menu9_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Inserir Tratamento de Erro em todas as Funções do módulo
   Dim n As Variant
   Dim i
   With VBInstance.ActiveCodePane.CodeModule
      For Each n In .Members
         i = i + 1
         If n.Type = vbext_mt_Method Then
            Call InserirErrFuncao(n.Name)
         End If
      Next
   End With
End Sub
Private Sub Menu10_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* TabOrder
   Call Show(3)
End Sub
Private Sub Menu11_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Construtor
   Me.Show
End Sub
Private Sub Menu12_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   'Dim x As VBIDE.Application
   Dim lAux As Long
   '* Biblioteca de Funções
   If Sys.Edit.ExeAuxiliar <> "" Then
      lAux = Shell(Sys.Edit.ExeAuxiliar, vbNormalFocus)
   End If
   'Call CreateMenu(VbApplication)
End Sub
Private Sub Menu13_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   '* Propriedades
   Screen.MousePointer = vbHourglass
   Call Show(2)
End Sub
Public Sub InserirCabFuncao(Optional NmFunc = "")
   Dim Str$, Linha%, Pos%, Divisao As Boolean
   Dim StrTemp$, IndLin%
   Dim StartLine&, StarColumn&, EndLine&, EndColumn&
   Dim MyCodePane As CodePane

   On Error GoTo Fim
   Set MyCodePane = VBInstance.ActiveCodePane
   
   With MyCodePane
      If Trim(NmFunc) = "" Then
         .GetSelection StartLine&, StarColumn&, EndLine&, EndColumn&
         NmFunc = .CodeModule.ProcOfLine(StartLine&, vbext_ProcKind.vbext_pk_Proc)
      End If
      Linha% = .CodeModule.ProcBodyLine(NmFunc, vbext_ProcKind.vbext_pk_Proc) + 1
   End With
   With Sys.Edit
      StrTemp = .Templates(.Template)
      StrTemp = Mid(StrTemp, InStr(StrTemp, " ") + 1)
      Pos = InStr(StrTemp, "|")
      
      Str$ = Str$ & "'**********************************************************************" & vbCrLf
      While Pos > 0
         IndLin = CInt(Mid(StrTemp, 1, Pos - 1))
         If IndLin <> 6 And IndLin <> 8 Then '* Projeto e Nome do Arquivo
            If IndLin > 11 And Not Divisao Then
               Divisao = True
               Str$ = Str$ & "'**********************************************************************" & vbCrLf
            End If
            Str$ = Str$ & LinhaCab(IndLin, NmFunc) & vbCrLf
         End If
         StrTemp = Mid(StrTemp, Pos + 1)
         Pos = InStr(StrTemp, "|")
      Wend
      Str$ = Str$ & "'**********************************************************************" & vbCrLf
      MyCodePane.CodeModule.InsertLines Linha%, Str$
      Str$ = Mid(Str$, 1, InStr(Str$, LoadRes(gwComentario)))
   End With
   
   While InStr(Str, vbCrLf) <> 0
      Str = Mid(Str, 1, Len(Str) - 1)
      If Right(Str, 1) = vbCr Then
         Linha = Linha + 1
      End If
   Wend

   Pos = Len(MyCodePane.CodeModule.Lines(Linha, 1)) + 1
   Pos = IIf(Pos <= 0, 1, Pos)
   MyCodePane.CodeModule.CodePane.SetSelection Linha, Pos, Linha, Pos
Fim:
End Sub
Public Sub InserirCabModulo()
   Dim Str$, Linha%, Pos%, Divisao As Boolean
   Dim StrTemp$, IndLin%
   Dim MyCodePane As CodePane

'   On Error GoTo Fim
   Set MyCodePane = VBInstance.ActiveCodePane
   
   Linha = 1
   
   With Sys.Edit
      StrTemp = .Templates(.Template)
      StrTemp = Mid(StrTemp, InStr(StrTemp, " ") + 1)
      Pos = InStr(StrTemp, "|")
      
      Str$ = Str$ & "'**********************************************************************" & vbCrLf
      While Pos > 0
         IndLin = CInt(Mid(StrTemp, 1, Pos - 1))
         If IndLin <> 9 And IndLin <> 10 Then '* Função e Parâmetros
            If IndLin > 10 And Not Divisao Then
               Divisao = True
               Str$ = Str$ & "'**********************************************************************" & vbCrLf
            End If
            Str$ = Str$ & LinhaCab(IndLin) & vbCrLf
         End If
         StrTemp = Mid(StrTemp, Pos + 1)
         Pos = InStr(StrTemp, "|")
      Wend
      Str$ = Str$ & "'**********************************************************************" & vbCrLf
      MyCodePane.CodeModule.InsertLines Linha, Str$
      Str$ = Mid(Str$, 1, InStr(Str$, LoadRes(gwComentario)))
   End With
  
   While InStr(Str, vbCrLf) <> 0
      Str = Mid(Str, 1, Len(Str) - 1)
      If Right(Str, 1) = vbCr Then
         Linha = Linha + 1
      End If
   Wend

   Pos = Len(MyCodePane.CodeModule.Lines(Linha, 1)) + 1
   Pos = IIf(Pos <= 0, 1, Pos)
   MyCodePane.CodeModule.CodePane.SetSelection Linha, Pos, Linha, Pos
Fim:
End Sub
Public Sub InserirErrFuncao(Optional NmFunc)
   Dim Str$, Linha%, Pos%
   Dim NmModulo$, NmArq$, Param$
   Dim StartLine&, StarColumn&, EndLine&, EndColumn&
   Dim MyCodePane As CodePane

   On Error GoTo Fim
   Set MyCodePane = VBInstance.ActiveCodePane
   
   With MyCodePane
      NmArq$ = GetNameFromPath(.CodeModule.Parent.FileNames(1))
      NmModulo$ = .CodeModule.Parent.Name
      If IsMissing(NmFunc) Then
         .GetSelection StartLine&, StarColumn&, EndLine&, EndColumn&
         NmFunc = .CodeModule.ProcOfLine(StartLine&, vbext_ProcKind.vbext_pk_Proc)
      End If
      Linha% = .CodeModule.ProcBodyLine(NmFunc, vbext_ProcKind.vbext_pk_Proc) + 1
      If Linha < .CodeModule.CountOfDeclarationLines Then
         GoTo Fim
      End If
      Str$ = "   On Error Goto Trata_Erro" & vbCrLf
      .CodeModule.InsertLines Linha%, Str$
      
      Linha% = .CodeModule.ProcStartLine(NmFunc, vbext_ProcKind.vbext_pk_Proc) + .CodeModule.ProcCountLines(NmFunc, vbext_ProcKind.vbext_pk_Proc) - 1
      
      Str$ = ""
      Str$ = Str$ & "Saida:" & vbCrLf
      Str$ = Str$ & "   Exit Sub" & vbCrLf
      Str$ = Str$ & "Trata_Erro:" & vbCrLf
      Str$ = Str$ & "   ShowError"
      .CodeModule.InsertLines Linha%, Str$

   End With
   
Fim:
End Sub
Private Sub RefHnd_ItemAdded(ByVal Reference As VBIDE.Reference)
'   MsgBox "RefHnd_ItemAdded"
End Sub
Private Sub RefHnd_ItemRemoved(ByVal Reference As VBIDE.Reference)
'   MsgBox "RefHnd_ItemAdded"
End Sub
Private Sub SelCtlHnd_ItemAdded(ByVal VBControl As VBIDE.VBControl)
'   MsgBox "SelCtlHnd_ItemAdded"
End Sub
Private Sub SelCtlHnd_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
'   MsgBox "SelCtlHnd_ItemRemoved"
End Sub
