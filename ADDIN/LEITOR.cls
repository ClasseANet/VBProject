VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LEITOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarfFile As Long
Private mvarLineCount As Integer
Private mvarPathFileName As String
Public Property Let PathFileName(ByVal vData As Long)
    mvarPathFileName = vData
End Property
Public Property Get PathFileName() As Long
    PathFileName = mvarPathFileName
End Property
Public Property Let fFile(ByVal vData As Long)
    mvarfFile = vData
End Property
'Public Property Get fFile() As Long
'    fFile = mvarfFile
'End Property
Public Property Let LineCount(ByVal vData As Long)
    mvarLineCount = vData
End Property
Public Property Get LineCount() As Long
    LineCount = mvarLineCount
End Property
Public Sub Ler_Classe(Optional pPathFileName)
   If IsMissing(pPathFileName) Then
      pPathFileName = mvarPathFileName
   End If
   mvarfArq = FreeFile
   Open pPathFileName For Input As #nArq
   While Not EOF(nArq)
      Line Input #nArq, Textline
      'Call AnalisarLinha_Bas(TextLine, pBAS)
      If Textline <> "" Then pBAS.LINHAS = CStr(Val(pBAS.LINHAS) + 1)
   Wend
End Sub
