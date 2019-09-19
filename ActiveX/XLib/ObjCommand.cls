VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjCommand"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function CriarToolbar(pSys As Object, pNmToolBar As String) As Object
   Dim oToolBar As Object 'CommandBars
   Dim oBar     As Object 'CommandBar
   Dim n        As Object 'CommandBar
   
   Set oToolBar = pSys.MDI.CommandBars
   With pSys
      '* Verificar se Toolbar Existe
      For Each n In oToolBar
         If n.Title = pNmToolBar Then
            Set oBar = n
            Exit For
         End If
      Next
      
      '* Se Toolbar não Existe então cria
      If oBar Is Nothing Then
         Set oBar = oToolBar.Add(pNmToolBar, 4)  ' 0=xtpBarTop, 4=xtpBarFloating
         oBar.Visible = False
      End If
   End With
   Set CriarToolbar = oBar
End Function
'Public Function CriarButtonToolbar(pToolbar As Object, pType As XTPControlType, pId As Long, _
         Optional pCaption As String, Optional pCategory As String, Optional pStyle As Integer = 2, _
         Optional pBeginGroup As Boolean, Optional pIconId As Long, Optional pChecked As Boolean, _
         Optional pParameter) As Object
Public Function CriarButtonToolbar(pToolbar As Object, pType As Integer, pId As Long, _
         Optional pCaption As String, Optional pCategory As String, Optional pStyle As Integer = 2, _
         Optional pBeginGroup As Boolean, Optional pIconId As Long, Optional pChecked As Boolean, _
         Optional pParameter) As Object
   
   Dim oControl As Object 'CommandBarControl
      
   With pToolbar
      Set oControl = .Controls.Find(pType, pId)
      If oControl Is Nothing Then
         Set oControl = .Controls.Add(pType, pId, pCaption)
         With oControl
            .Category = pCategory
            .IconId = pIconId
            .Checked = pChecked
            .BeginGroup = pBeginGroup
            .Style = pStyle
            If Not IsMissing(pParameter) Then
               .Parameter = pParameter
            End If
         End With
      End If
   End With
   
   Set CriarButtonToolbar = oControl
End Function
Public Function AddButtonBar(Controls As Object, _
                              Id As Long, Caption As String, _
                              Optional BeginGroup As Boolean = False, _
                              Optional ControlType As Integer = 1, _
                              Optional Category As String = "") As Object
'Public Function AddButtonBar(Controls As CommandBarControls, _
                              Id As Long, Caption As String, _
                              Optional BeginGroup As Boolean = False, _
                              Optional ControlType As XTPControlType = xtpControlButton, _
                              Optional Category As String = "") As CommandBarControl
   
   'Dim oMenuItem As CommandBarControl
   Dim oMenuItem As Object
    
   Set oMenuItem = Controls.Add(ControlType, Id, Caption)
   With oMenuItem
      .BeginGroup = BeginGroup
    
      .Category = Category
      .Parameter = ClsGeneral.SetTag(.Parameter, "CARREGADO", 0)
      .Parameter = ClsGeneral.SetTag(.Parameter, "MENUCHILD", "S")
   End With
   
   Set AddButtonBar = oMenuItem
End Function
Public Sub DockBarRightOf(pBarToDock As Variant, pBarOnLeft As Variant, Optional pSys As Object)
   Dim ToolBar As Object 'CommandBars
   Dim nBar    As Object
   Dim Left    As Long
   Dim Top     As Long
   Dim Right   As Long
   Dim Bottom  As Long
   
   Dim BarToDock As Object 'CommandBar
   Dim BarOnLeft As Object 'CommandBar
    
   If Not IsEmpty(pSys) And Not pSys Is Nothing Then
      Set ToolBar = pSys.MDI.CommandBars
   End If
      
   If ClsGeneral.InArray(TypeName(pBarToDock), Array("IMenuBar", "ICommandBar")) Then
      Set BarToDock = pBarToDock
   ElseIf TypeName(pBarOnLeft) = "String" Then
      For Each nBar In ToolBar
         If nBar.Title = pBarToDock Then
            Set BarToDock = nBar
            Exit For
         End If
      Next
   ElseIf TypeName(pBarOnLeft) = "Integer" Then
      Set BarToDock = ToolBar(pBarToDock)
   End If
   
   Set ToolBar = BarToDock.CommandBars
   
   If ClsGeneral.InArray(TypeName(pBarOnLeft), Array("IMenuBar", "ICommandBar")) Then
      Set BarOnLeft = pBarOnLeft
   ElseIf TypeName(pBarOnLeft) = "String" Then
      For Each nBar In ToolBar
         If nBar.Title = pBarOnLeft Then
            Set BarOnLeft = nBar
            Exit For
         End If
      Next
   ElseIf TypeName(pBarOnLeft) = "Integer" Then
      Set BarOnLeft = ToolBar(pBarOnLeft)
   End If
    
   If BarToDock Is Nothing Then Exit Sub
   If BarOnLeft Is Nothing Then Exit Sub
       
    ToolBar.RecalcLayout
        
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    ToolBar.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub


