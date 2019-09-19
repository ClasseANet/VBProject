VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmEditorH
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarFile           As String
Public Property Let File(ByVal vData As String)
    mvarFile = vData
End Property
Public Property Get File() As String
   File = mvarFile
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
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmEditorH
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Load()
    On Error Resume Next
    
    Dim x As Long
    '---------------------------------------
    mvarMe.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    mvarMe.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    mvarMe.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    mvarMe.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    '---------------------------------------
    sDocFileName = "Untitled.html"
    SetCaption
    
    mvarMe.HtmlEditor1.Visible = True
    CurrentMode = 0
    'Do While Not mvarme.HtmlEditor1.ReadyState: DoEvents: Loop
    'Debug.Print mvarme.HtmlEditor1.ReadyState
    '----------------------------------------------------------------
    mvarMe.HtmlEditor1.designMode = True
    
    mvarMe.HtmlEditor1.LiveResize = True
    mvarMe.HtmlEditor1.MultipleSelection = True
    mvarMe.HtmlEditor1.SelectTables = False
    
    'mvarme.HtmlEditor1.ShowBorders = True
    'mvarme.HtmlEditor1.Content = "hello <table><tr><td>ahmed amin</td></tr></table>"
    mvarMe.HtmlEditor1.DocumentHtml = "" '"hello <p>I am here</p><br><table width=450 cellspacing=0 cellpadding=2><tr><td width=350>ahmed amin</td><td>cell  2</td></tr><tr><td>elsheshtawy</td><td>cell 4</td></tr><tr><td>elsheshtawy</td><td>cell 4</td></tr><tr><td>elsheshtawy</td><td>cell 4</td></tr></table> <!--comment here--><br><script>var x;</script> <style>Shesh</style>"
        
    Dim Document As HTMLDocument
    
    Set Document = mvarMe.HtmlEditor1.Document
    'Document.body.innerHTML = "ahmed"
    'Document.execCommand "bold"
    '----------------------------------------------------------------
    Dim Formats() As String
    
    Formats() = mvarMe.HtmlEditor1.GetBlockFormats
    For x = LBound(Formats) To UBound(Formats) - 1
        mvarMe.StyleCombo.AddItem Formats(x)
    Next x
    '----------------------------------------------------------------
    Dim fontNames() As String
    
    fontNames = mvarMe.HtmlEditor1.GetFontNames
    For x = LBound(fontNames) To UBound(fontNames)
        mvarMe.FontCombo.AddItem fontNames(x)
    Next x

    mvarMe.FontCombo.ListIndex = 0
    '----------------------------------------------------------------
    ' Font size menu
    For x = 1 To 7
        mvarMe.FontSizeCombo.AddItem CStr(x)
    Next x
    
    mvarMe.FontSizeCombo.ListIndex = 0
    '----------------------------------------------------------------
    'mvarme.HtmlEditor1.ShowBreakGlyph True
    'mvarme.HtmlEditor1.ShowAllGlyph True
    'mvarme.HtmlEditor1.ShowAreaGlyph True
    'mvarme.HtmlEditor1.ShowCommentGlyph True
    'mvarme.HtmlEditor1.ShowStyleGlyph True
    'mvarme.HtmlEditor1.DisableEditFocus True
    'mvarme.HtmlEditor1.KeepSelection True
    'mvarme.HtmlEditor1.OverrideCursor False
    'mvarme.HtmlEditor1.InsertObject
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    If mvarFile <> "" Then Call AbrirArquivo(mvarFile)
End Sub
Private Sub SetCaption()
    mvarMe.Caption = AppName & " [" & sDocFileName & "]"
End Sub

Private Sub mvarMe_mnuFileOpenClick()
    Call AbrirArquivo
End Sub
Public Sub AbrirArquivo(Optional ByRef pFile As String)
   Dim sFile As String
   Dim FileNumber As Long
    
   If ExisteArquivo(pFile) Then
      sFile = pFile
   Else
      With mvarMe.CommonDialog1
          .DialogTitle = "Open"
          .CancelError = False
          'ToDo: set the flags and attributes of the common dialog control
          .Filter = "Html Files (*.html)|*.htm;*.html"
          .ShowOpen
          If Len(.FileName) = 0 Then
              Exit Sub
          End If
          sFile = .FileName
      End With
   End If
   
   'ActiveForm.Caption = sFile
   'ActiveForm.rtfText.SaveFile sFile
   'ActiveForm.SourceEditor1.SetFocus
   If sFile = "" Then Exit Sub
   If Dir(sFile) = "" Then Exit Sub
   
   Dim DocText As String
   
   DocText = mvarMe.HtmlEditor1.DocumentHtml
    
   FileNumber = FreeFile
   Open sFile For Input As FileNumber
   DocText = Input(LOF(FileNumber), FileNumber)
   Close FileNumber
   
   mvarMe.HtmlEditor1.DocumentHtml = DocText
   DocText = GetHTMLCode(sFile)
   If Trim(DocText) <> "" Then
     mvarMe.HtmlEditor1.DocumentHtml = DocText
   End If
   
   sDocFileName = sFile
   mvarFile = sFile
   SetCaption
   Exit Sub
Errors:
    MsgBox "Error loading file: " & Err.Description
End Sub

Private Sub mvarMe_mnuFileSaveAsClick()
    Dim sFile As String
    Dim FileNumber As Long
        
    With mvarMe.CommonDialog1
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Html Files (*.html)|*.html"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    
    If sFile = "" Then Exit Sub
    sDocFileName = sFile
    
    Dim DocText As String
    
    DocText = mvarMe.HtmlEditor1.DocumentHtml
    DocText = SetHTMLCode(sFile, DocText)
    If Trim(DocText) = "" Then
      DocText = mvarMe.HtmlEditor1.DocumentHtml
    End If
    
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, DocText
    Close #FileNumber
    
    mvarFile = sFile
    sDocFileName = sFile
    Call SetCaption
    Exit Sub
Errors:
    MsgBox "Error saving file: " & Err.Description
End Sub

Private Sub mvarMe_mnuFileSaveClick()
   Dim FileNumber As Long
   Dim DocText As String
    
   If sDocFileName = "Untitled.html" Then
       Call mvarMe_mnuFileSaveAsClick
       Exit Sub
   End If
   
   On Error GoTo Errors
   
   DocText = mvarMe.HtmlEditor1.DocumentHtml
   DocText = SetHTMLCode(sDocFileName, DocText)
   If Trim(DocText) = "" Then
     DocText = mvarMe.HtmlEditor1.DocumentHtml
   End If
   
   FileNumber = FreeFile
   Open sDocFileName For Output As #FileNumber
   Print #FileNumber, DocText
   Close #FileNumber
   
   mvarFile = sDocFileName
   Call SetCaption
   Exit Sub
Errors:
    MsgBox "Error saving file: " & Err.Description
End Sub
