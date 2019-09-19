VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl PctAzulEscuro 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox PctOlivaEscuro 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3720
      Picture         =   "MCIGrid.ctx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PctCinzaEscuro 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3720
      Picture         =   "MCIGrid.ctx":0938
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PctOlivaClaro 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3360
      Picture         =   "MCIGrid.ctx":1470
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PctCinzaClaro 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3360
      Picture         =   "MCIGrid.ctx":1C69
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PctAzulEscuro 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3720
      Picture         =   "MCIGrid.ctx":2556
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CellTextBox 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PctAzulClaro 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3360
      Picture         =   "MCIGrid.ctx":2EA3
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ObjGrid 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3836
      _Version        =   393216
      BackColorSel    =   16777152
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   8421504
      GridColorUnpopulated=   -2147483632
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "PctAzulEscuro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Property Variables:
Dim m_ActiveControl As Control
'Event Declarations:
Event Click() 'MappingInfo=ObjGrid,ObjGrid,-1,Click
Attribute Click.VB_Description = "Fired when the user presses and releases the mouse button over the control."
Event DblClick() 'MappingInfo=ObjGrid,ObjGrid,-1,DblClick
Attribute DblClick.VB_Description = "Fired when the user double-clicks the mouse over the control."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=ObjGrid,ObjGrid,-1,KeyDown
Attribute KeyDown.VB_Description = "Fired when the user pushes a key."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=ObjGrid,ObjGrid,-1,KeyPress
Attribute KeyPress.VB_Description = "Fired when the user presses a key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=ObjGrid,ObjGrid,-1,KeyUp
Attribute KeyUp.VB_Description = "Fired when the user releases a key."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=ObjGrid,ObjGrid,-1,MouseDown
Attribute MouseDown.VB_Description = "Fired when the user presses a mouse button over the control."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=ObjGrid,ObjGrid,-1,MouseMove
Attribute MouseMove.VB_Description = "Fired when the user moves the mouse over the control."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=ObjGrid,ObjGrid,-1,MouseUp
Attribute MouseUp.VB_Description = "Fired when the user releases a mouse button over the control."
Event WriteProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Attribute WriteProperties.VB_Description = "Occurs when a user control or user document is asked to write its data to a file."
Event Validate(Cancel As Boolean) 'MappingInfo=ObjGrid,ObjGrid,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Event SelChange() 'MappingInfo=ObjGrid,ObjGrid,-1,SelChange
Attribute SelChange.VB_Description = "Fired when the selected range of cells changes."
Event Scroll() 'MappingInfo=ObjGrid,ObjGrid,-1,Scroll
Attribute Scroll.VB_Description = "Fired when the TopRow or LeftCol properties change."
Event RowColChange() 'MappingInfo=ObjGrid,ObjGrid,-1,RowColChange
Attribute RowColChange.VB_Description = "Fired when the current cell changes."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Attribute ReadProperties.VB_Description = "Occurs when a user control or user document is asked to read its data from a file."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=ObjGrid,ObjGrid,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "OLEStartDrag event."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=ObjGrid,ObjGrid,-1,OLESetData
Attribute OLESetData.VB_Description = "OLESetData event."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=ObjGrid,ObjGrid,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "OLEGiveFeedback event."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer) 'MappingInfo=ObjGrid,ObjGrid,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "OLEDragOver event."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=ObjGrid,ObjGrid,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "OLEDragDrop event."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=ObjGrid,ObjGrid,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "OLECompleteDrag event."
Event LeaveCell() 'MappingInfo=ObjGrid,ObjGrid,-1,LeaveCell
Attribute LeaveCell.VB_Description = "Fired after the cursor leaves a cell."
Event InitProperties() 'MappingInfo=UserControl,UserControl,-1,InitProperties
Attribute InitProperties.VB_Description = "Occurs the first time a user control or user document is created."
Event HitTest(x As Single, Y As Single, HitResult As Integer) 'MappingInfo=UserControl,UserControl,-1,HitTest
Attribute HitTest.VB_Description = "Occurs in a windowless user control in response to mouse activity."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event GetDataMember(DataMember As String, Data As Object) 'MappingInfo=UserControl,UserControl,-1,GetDataMember
Attribute GetDataMember.VB_Description = "Occurs when a data consumer is asking this data source for one of it's data members."
Event Expand(Cancel As Boolean) 'MappingInfo=ObjGrid,ObjGrid,-1,Expand
Attribute Expand.VB_Description = "Fired when user clicks on expand graphic to expand collapsed data."
Event EnterCell() 'MappingInfo=ObjGrid,ObjGrid,-1,EnterCell
Attribute EnterCell.VB_Description = "Fired before the cursor enters a cell."
Event Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer) 'MappingInfo=ObjGrid,ObjGrid,-1,Compare
Attribute Compare.VB_Description = "Fired during custom sorts to compare two rows."
Event Collapse(Cancel As Boolean) 'MappingInfo=ObjGrid,ObjGrid,-1,Collapse
Attribute Collapse.VB_Description = "Fired when user clicks on collapse graphic to collapse expanded data."
Event AsyncReadProgress(AsyncProp As AsyncProperty) 'MappingInfo=UserControl,UserControl,-1,AsyncReadProgress
Attribute AsyncReadProgress.VB_Description = "Occurs when more data is available as a result of the AsyncReadProgress method."
Event AsyncReadComplete(AsyncProp As AsyncProperty) 'MappingInfo=UserControl,UserControl,-1,AsyncReadComplete
Attribute AsyncReadComplete.VB_Description = "Occurs when all of the data is available as a result of the AsyncRead method."
Private mvarSort As Integer
Private Sub UserControl_Initialize()
   Call DefineImgCab
   Call CarregaControles
End Sub
Private Sub UserControl_Resize()
   RaiseEvent Resize
   ObjGrid.Move 0, 0, UserControl.Width, UserControl.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   UserControl.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
   ForeColor = ObjGrid.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
   ObjGrid.ForeColor() = New_ForeColor
   PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
   Enabled = ObjGrid.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   ObjGrid.Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns or sets the default font or the font for individual cells."
Attribute Font.VB_UserMemId = -512
   Set Font = ObjGrid.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set ObjGrid.Font = New_Font
   PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
   BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
   UserControl.BackStyle() = New_BackStyle
   PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleSettings
Attribute BorderStyle.VB_Description = "Returns or sets the border style for an object."
   BorderStyle = ObjGrid.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
   ObjGrid.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
   ObjGrid.Refresh
End Sub

Private Sub ObjGrid_Click()
   RaiseEvent Click
End Sub

Private Sub ObjGrid_DblClick()
   RaiseEvent DblClick
End Sub

Private Sub ObjGrid_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub ObjGrid_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub ObjGrid_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ObjGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub ObjGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub ObjGrid_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Dim Index         As Integer
   Dim BandNumber    As Long
   Dim BandColIndex  As Long
   Dim BandData      As Long
   
   RaiseEvent WriteProperties(PropBag)
   Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
   Call PropBag.WriteProperty("ForeColor", ObjGrid.ForeColor, &H80000008)
   Call PropBag.WriteProperty("Enabled", ObjGrid.Enabled, True)
   Call PropBag.WriteProperty("Font", ObjGrid.Font, Ambient.Font)
   Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
   Call PropBag.WriteProperty("BorderStyle", ObjGrid.BorderStyle, 0)
   Call PropBag.WriteProperty("WordWrap", ObjGrid.WordWrap, False)
   Call PropBag.WriteProperty("WhatsThisHelpID", ObjGrid.WhatsThisHelpID, 0)
   Call PropBag.WriteProperty("ToolTipText", ObjGrid.ToolTipText, "")
   Call PropBag.WriteProperty("TextStyleFixed", ObjGrid.TextStyleFixed, 0)
   Call PropBag.WriteProperty("TextStyle", ObjGrid.TextStyle, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   Call PropBag.WriteProperty("TextMatrix" & Index, ObjGrid.TextMatrix(Row, Col), "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("TextArray" & Index, ObjGrid.TextArray(Index), "")
   'Call PropBag.WriteProperty("Sort", ObjGrid.Sort, 0)
   mvarSort = 0
   Call PropBag.WriteProperty("SelectionMode", ObjGrid.SelectionMode, 1)
   Call PropBag.WriteProperty("ScrollTrack", ObjGrid.ScrollTrack, False)
   Call PropBag.WriteProperty("ScrollBars", ObjGrid.ScrollBars, 3)
   Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 4800)
   Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
   Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
   Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
   Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3600)
   Call PropBag.WriteProperty("RowSizingMode", ObjGrid.RowSizingMode, 0)
   Call PropBag.WriteProperty("Rows", ObjGrid.Rows, 2)
   Call PropBag.WriteProperty("Columns", ObjGrid.Cols(0), 3)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   Call PropBag.WriteProperty("RowPosition" & Index, ObjGrid.RowPosition(Index), 0)
   Call PropBag.WriteProperty("RowHeightMin", ObjGrid.RowHeightMin, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("RowHeight" & Index, ObjGrid.RowHeight(Index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("RowData" & Index, ObjGrid.RowData(Index), 0)
   Call PropBag.WriteProperty("RightToLeft", ObjGrid.RightToLeft, 0)
   Call PropBag.WriteProperty("Redraw", ObjGrid.Redraw, True)
   Call PropBag.WriteProperty("Recordset", Recordset, Nothing)
   Call PropBag.WriteProperty("PictureType", ObjGrid.PictureType, 0)
   Call PropBag.WriteProperty("PaletteMode", UserControl.PaletteMode, 3)
   Call PropBag.WriteProperty("Palette", Palette, Nothing)
   Call PropBag.WriteProperty("OLEDropMode", ObjGrid.OLEDropMode, 0)
   Call PropBag.WriteProperty("MousePointer", ObjGrid.MousePointer, 0)
   Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
   Call PropBag.WriteProperty("MergeCells", ObjGrid.MergeCells, 0)
   Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
   Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
   Call PropBag.WriteProperty("HitBehavior", UserControl.HitBehavior, 1)
   Call PropBag.WriteProperty("HighLight", ObjGrid.HighLight, 1)
   Call PropBag.WriteProperty("GridLineWidthUnpopulated", ObjGrid.GridLineWidthUnpopulated, 1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("GridLineWidthIndent" & Index, ObjGrid.GridLineWidthIndent(BandNumber), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("GridLineWidthHeader" & Index, ObjGrid.GridLineWidthHeader(BandNumber), 0)
   Call PropBag.WriteProperty("GridLineWidthFixed", ObjGrid.GridLineWidthFixed, 1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("GridLineWidthBand" & Index, ObjGrid.GridLineWidthBand(BandNumber), 0)
   Call PropBag.WriteProperty("GridLineWidth", ObjGrid.GridLineWidth, 1)
   Call PropBag.WriteProperty("GridLinesUnpopulated", ObjGrid.GridLinesUnpopulated, 1)
   Call PropBag.WriteProperty("GridLinesFixed", ObjGrid.GridLinesFixed, 1)
   Call PropBag.WriteProperty("GridLines", ObjGrid.GridLines, 1)
   Call PropBag.WriteProperty("GridColorUnpopulated", ObjGrid.GridColorUnpopulated, &H80000010)
   Call PropBag.WriteProperty("GridColorFixed", ObjGrid.GridColorFixed, &H80000010)
   Call PropBag.WriteProperty("GridColor", ObjGrid.GridColor, &H80000010)
   Call PropBag.WriteProperty("FormatString", ObjGrid.FormatString, "")
   Call PropBag.WriteProperty("ForeColorSel", ObjGrid.ForeColorSel, 2147483662#)
   Call PropBag.WriteProperty("ForeColorFixed", ObjGrid.ForeColorFixed, 2147483666#)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("FontWidthHeader" & Index, ObjGrid.FontWidthHeader(BandNumber), 0)
   Call PropBag.WriteProperty("FontWidthFixed", ObjGrid.FontWidthFixed, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("FontWidthBand" & Index, ObjGrid.FontWidthBand(BandNumber), 0)
   Call PropBag.WriteProperty("FontWidth", ObjGrid.FontWidth, 0)
   Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
   Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)
   Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
   Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
   Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
   Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
'   Call PropBag.WriteProperty("FontHeader", ObjGrid.FontHeader, Ambient.Font)
   Call PropBag.WriteProperty("FontFixed", ObjGrid.FontFixed, Ambient.Font)
   Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
'   Call PropBag.WriteProperty("FontBand", ObjGrid.FontBand, Ambient.Font)
   Call PropBag.WriteProperty("FocusRect", ObjGrid.FocusRect, 1)
   Call PropBag.WriteProperty("FixedRows", ObjGrid.FixedRows, 1)
   Call PropBag.WriteProperty("FixedCols", ObjGrid.FixedCols, 1)
   Call PropBag.WriteProperty("FillStyle", ObjGrid.FillStyle, 0)
   Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
   Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
   Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
   Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
   Call PropBag.WriteProperty("DataSource", DataSource, Nothing)
   Call PropBag.WriteProperty("DataMember", ObjGrid.DataMember, "")
   Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
   Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColWordWrapOptionHeader" & Index, ObjGrid.ColWordWrapOptionHeader(BandNumber, BandColIndex), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColWordWrapOptionFixed" & Index, ObjGrid.ColWordWrapOptionFixed(Index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColWordWrapOptionBand" & Index, ObjGrid.ColWordWrapOptionBand(BandNumber, BandColIndex), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColWordWrapOption" & Index, ObjGrid.ColWordWrapOption(Index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColWidth" & Index, ObjGrid.ColWidth(Index, BandNumber), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("Cols" & Index, ObjGrid.Cols(BandNumber), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   Call PropBag.WriteProperty("ColPosition" & Index, ObjGrid.ColPosition(Index, BandNumber), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColHeaderCaption" & Index, ObjGrid.ColHeaderCaption(BandNumber, BandColIndex), "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColData" & Index, ObjGrid.ColData(Index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColAlignmentHeader" & Index, ObjGrid.ColAlignmentHeader(BandNumber, BandColIndex), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColAlignmentFixed" & Index, ObjGrid.ColAlignmentFixed(Index), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColAlignmentBand" & Index, ObjGrid.ColAlignmentBand(BandNumber, BandColIndex), 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("ColAlignment" & Index, ObjGrid.ColAlignment(Index), 0)
   Call PropBag.WriteProperty("ClipControls", UserControl.ClipControls, True)
   Call PropBag.WriteProperty("ClipBehavior", UserControl.ClipBehavior, 1)
   Call PropBag.WriteProperty("CausesValidation", ObjGrid.CausesValidation, True)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("BandIndent" & Index, ObjGrid.BandIndent(BandNumber), 0)
   Call PropBag.WriteProperty("BandDisplay", ObjGrid.BandDisplay, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   Call PropBag.WriteProperty("BandData" & Index, ObjGrid.BandData(BandData), 0)
   Call PropBag.WriteProperty("BackColorUnpopulated", ObjGrid.BackColorUnpopulated, 2147483663#)
   Call PropBag.WriteProperty("BackColorSel", ObjGrid.BackColorSel, 2147483661#)
   Call PropBag.WriteProperty("BackColorFixed", ObjGrid.BackColorFixed, 2147483663#)
   Call PropBag.WriteProperty("BackColorBkg", ObjGrid.BackColorBkg, &H80000005)
   Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
   Call PropBag.WriteProperty("Appearance", ObjGrid.Appearance, 1)
   Call PropBag.WriteProperty("AllowUserResizing", ObjGrid.AllowUserResizing, 0)
   Call PropBag.WriteProperty("AllowBigSelection", ObjGrid.AllowBigSelection, True)
   Call PropBag.WriteProperty("ActiveControl", m_ActiveControl, Nothing)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,WordWrap
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns or sets whether text within a cell should be allowed to wrap to multiple lines."
   WordWrap = ObjGrid.WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
   ObjGrid.WordWrap() = New_WordWrap
   PropertyChanged "WordWrap"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,WhatsThisHelpID
Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
   WhatsThisHelpID = ObjGrid.WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
   ObjGrid.WhatsThisHelpID() = New_WhatsThisHelpID
   PropertyChanged "WhatsThisHelpID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Version
Public Property Get Version() As Integer
Attribute Version.VB_Description = "Returns the version of the currently loaded Hierarchical FlexGrid control."
   Version = ObjGrid.Version
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ValidateControls
Public Sub ValidateControls()
Attribute ValidateControls.VB_Description = "Validate contents of the last control on the form before exiting the form"
   UserControl.ValidateControls
End Sub

Private Sub ObjGrid_Validate(Cancel As Boolean)
   RaiseEvent Validate(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
   ToolTipText = ObjGrid.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
   ObjGrid.ToolTipText() = New_ToolTipText
   PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
   TextWidth = UserControl.TextWidth(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,TextStyleFixed
Public Property Get TextStyleFixed() As TextStyleSettings
Attribute TextStyleFixed.VB_Description = "Returns or sets 3-D effects for displaying text."
   TextStyleFixed = ObjGrid.TextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal New_TextStyleFixed As TextStyleSettings)
   ObjGrid.TextStyleFixed() = New_TextStyleFixed
   PropertyChanged "TextStyleFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,TextStyle
Public Property Get TextStyle() As TextStyleSettings
Attribute TextStyle.VB_Description = "Returns or sets 3-D effects for displaying text."
   TextStyle = ObjGrid.TextStyle
End Property

Public Property Let TextStyle(ByVal New_TextStyle As TextStyleSettings)
   ObjGrid.TextStyle() = New_TextStyle
   PropertyChanged "TextStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,TextMatrix
Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
Attribute TextMatrix.VB_Description = "Returns or sets the text content of an arbitrary cell (row/column subscripts)."
   TextMatrix = ObjGrid.TextMatrix(Row, Col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
   ObjGrid.TextMatrix(Row, Col) = New_TextMatrix
   PropertyChanged "TextMatrix"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
   TextHeight = UserControl.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,TextArray
Public Property Get TextArray(ByVal Index As Long) As String
Attribute TextArray.VB_Description = "Returns or sets the text content of an arbitrary cell (single subscript)."
   TextArray = ObjGrid.TextArray(Index)
End Property

Public Property Let TextArray(ByVal Index As Long, ByVal New_TextArray As String)
   ObjGrid.TextArray(Index) = New_TextArray
   PropertyChanged "TextArray"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Sort
Public Property Get Sort() As Integer
Attribute Sort.VB_Description = "An action-type property that sorts selected rows according to specified criteria. Not available at design time; write-only at run time."
   Sort = mvarSort
   'Sort = ObjGrid.Sort
End Property

Public Property Let Sort(ByVal New_Sort As Integer)
   ObjGrid.Sort() = New_Sort
   mvarSort = New_Sort
   PropertyChanged "Sort"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Size
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
   UserControl.Size Width, Height
End Sub

Private Sub UserControl_Show()
   RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,SelectionMode
Public Property Get SelectionMode() As SelectionModeSettings
Attribute SelectionMode.VB_Description = "Returns or sets whether a Hierarchical FlexGrid should allow regular cell selection, selection by rows, or selection by columns."
   SelectionMode = ObjGrid.SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
   ObjGrid.SelectionMode() = New_SelectionMode
   PropertyChanged "SelectionMode"
End Property

Private Sub ObjGrid_SelChange()
   RaiseEvent SelChange
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ScrollTrack
Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_Description = "Returns or sets whether the Hierarchical FlexGrid should scroll its contents while the user moves the scroll box along the scroll bars."
   ScrollTrack = ObjGrid.ScrollTrack
End Property

Public Property Let ScrollTrack(ByVal New_ScrollTrack As Boolean)
   ObjGrid.ScrollTrack() = New_ScrollTrack
   PropertyChanged "ScrollTrack"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ScrollBars
Public Property Get ScrollBars() As ScrollBarsSettings
Attribute ScrollBars.VB_Description = "Returns or sets whether the Hierarchical FlexGrid has horizontal or vertical scroll bars."
   ScrollBars = ObjGrid.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
   ObjGrid.ScrollBars() = New_ScrollBars
   PropertyChanged "ScrollBars"
End Property

Private Sub ObjGrid_Scroll()
   RaiseEvent Scroll
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleY
Public Function ScaleY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Converts the value for the height of a Form, PictureBox, or Printer from one unit of measure to another."
   ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleX
Public Function ScaleX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."
   ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
   ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
   UserControl.ScaleWidth() = New_ScaleWidth
   PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
   ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
   UserControl.ScaleTop() = New_ScaleTop
   PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
   ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
   UserControl.ScaleMode() = New_ScaleMode
   PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
   ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
   UserControl.ScaleLeft() = New_ScaleLeft
   PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
   ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
   UserControl.ScaleHeight() = New_ScaleHeight
   PropertyChanged "ScaleHeight"
End Property

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)
   UserControl.Scale (X1, Y1)-(X2, Y2)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowSizingMode
Public Property Get RowSizingMode() As RowSizingSettings
Attribute RowSizingMode.VB_Description = "Returns or sets the row sizing mode."
   RowSizingMode = ObjGrid.RowSizingMode
End Property

Public Property Let RowSizingMode(ByVal New_RowSizingMode As RowSizingSettings)
   ObjGrid.RowSizingMode() = New_RowSizingMode
   PropertyChanged "RowSizingMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Rows
Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Determines the total number of columns or rows in the Hierarchical FlexGrid."
   Rows = ObjGrid.Rows
End Property
Public Property Let Rows(ByVal New_Rows As Long)
   ObjGrid.Rows() = New_Rows
   PropertyChanged "Rows"
   Call DefineImgCab
   Call CarregaControles
End Property
Public Property Get Columns() As Long
   Columns = ObjGrid.Cols
End Property

Public Property Let Columns(ByVal New_Columns As Long)
   ObjGrid.Cols(0) = New_Columns
   PropertyChanged "Columns"
   Call DefineImgCab
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowPosition
'Public Property Get RowPosition(ByVal Index As Long) As Long
'   RowPosition = ObjGrid.RowPosition(Index)
'End Property

Public Property Let RowPosition(ByVal Index As Long, ByVal New_RowPosition As Long)
Attribute RowPosition.VB_Description = "Returns the distance, in Twips, between the upper left corner of the control and the upper left corner of a specified row."
   ObjGrid.RowPosition(Index) = New_RowPosition
   PropertyChanged "RowPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowPos
Public Property Get RowPos(ByVal Index As Long) As Long
Attribute RowPos.VB_Description = "Returns the distance, in Twips, between the upper left corner of the control and the upper left corner of a specified row."
   RowPos = ObjGrid.RowPos(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowIsVisible
Public Property Get RowIsVisible(ByVal Index As Long) As Boolean
Attribute RowIsVisible.VB_Description = "Returns True if the specified row is visible."
   RowIsVisible = ObjGrid.RowIsVisible(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowHeightMin
Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Returns or sets a minimum row height for the entire control, in Twips."
   RowHeightMin = ObjGrid.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
   ObjGrid.RowHeightMin() = New_RowHeightMin
   PropertyChanged "RowHeightMin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowHeight
Public Property Get RowHeight(ByVal Index As Long) As Long
Attribute RowHeight.VB_Description = "Returns or sets the height of the specified row, in Twips. Not available at design time."
   RowHeight = ObjGrid.RowHeight(Index)
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal New_RowHeight As Long)
   ObjGrid.RowHeight(Index) = New_RowHeight
   PropertyChanged "RowHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowExpandable
Public Property Get RowExpandable() As Boolean
Attribute RowExpandable.VB_Description = "Returns the expand and collapse state of the current row in the current band."
   RowExpandable = ObjGrid.RowExpandable
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RowData
Public Property Get RowData(ByVal Index As Long) As Long
Attribute RowData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time."
   RowData = ObjGrid.RowData(Index)
End Property

Public Property Let RowData(ByVal Index As Long, ByVal New_RowData As Long)
   ObjGrid.RowData(Index) = New_RowData
   PropertyChanged "RowData"
End Property

Private Sub ObjGrid_RowColChange()
   RaiseEvent RowColChange
   
   Dim i As Integer
   Dim LinAntes As Integer
   Dim ColAntes As Integer
   
   ObjGrid.Redraw = False
   LinAntes = ObjGrid.Row
   ColAntes = ObjGrid.Col
   For i = 1 To ObjGrid.Cols - 1
      ObjGrid.Col = i
      CellTextBox(i - 1).Left = ObjGrid.CellLeft
      CellTextBox(i - 1).Top = ObjGrid.CellTop
      CellTextBox(i - 1).Height = ObjGrid.CellHeight
      CellTextBox(i - 1).Width = ObjGrid.CellWidth
      CellTextBox(i - 1).BackColor = ObjGrid.BackColorSel
      CellTextBox(i - 1).Text = ObjGrid.Text
      CellTextBox(i - 1).Visible = True
      CellTextBox(i - 1).ZOrder 0
      If ObjGrid.CellAlignment <= 2 Then
         CellTextBox(i - 1).Alignment = 0
      ElseIf ObjGrid.CellAlignment >= 3 And ObjGrid.CellAlignment <= 5 Then
         CellTextBox(i - 1).Alignment = 2
      Else
         CellTextBox(i - 1).Alignment = 1
      End If
      
   Next
   ObjGrid.Row = LinAntes
   ObjGrid.Col = ColAntes
   ObjGrid.Redraw = True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and controls visual appearance on a bidirectional system."
   RightToLeft = ObjGrid.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
   ObjGrid.RightToLeft() = New_RightToLeft
   PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,RemoveItem
Public Sub RemoveItem(ByVal Index As Long)
Attribute RemoveItem.VB_Description = "Removes a row from a Hierarchical FlexGrid control at run time"
   ObjGrid.RemoveItem Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Redraw
Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Enables or disables redrawing of the Hierarchical FlexGrid control."
   Redraw = ObjGrid.Redraw
End Property

Public Property Let Redraw(ByVal New_Redraw As Boolean)
   ObjGrid.Redraw() = New_Redraw
   PropertyChanged "Redraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Recordset
Public Property Get Recordset() As IRecordset
Attribute Recordset.VB_Description = "Binds the Hierarchical FlexGrid to an ADO Recordset. Not available at design time."
   Set Recordset = ObjGrid.Recordset
End Property

Public Property Set Recordset(ByVal New_Recordset As IRecordset)
   Set ObjGrid.Recordset = New_Recordset
   PropertyChanged "Recordset"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Dim Index         As Integer
   Dim Row           As Long
   Dim Col           As Long
   Dim BandNumber    As Long
   Dim BandData      As Long
   Dim BandColIndex  As Long
   
   RaiseEvent ReadProperties(PropBag)
   UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
   ObjGrid.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
   ObjGrid.Enabled = PropBag.ReadProperty("Enabled", True)
   Set ObjGrid.Font = PropBag.ReadProperty("Font", Ambient.Font)
   UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
   ObjGrid.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
   ObjGrid.WordWrap = PropBag.ReadProperty("WordWrap", False)
   ObjGrid.WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
   ObjGrid.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
   ObjGrid.TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", 0)
   ObjGrid.TextStyle = PropBag.ReadProperty("TextStyle", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.TextMatrix(Row, Col) = PropBag.ReadProperty("TextMatrix" & Index, "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.TextArray(Index) = PropBag.ReadProperty("TextArray" & Index, "")
   ObjGrid.Sort = PropBag.ReadProperty("Sort", 0)
   ObjGrid.SelectionMode = PropBag.ReadProperty("SelectionMode", 1)
   ObjGrid.ScrollTrack = PropBag.ReadProperty("ScrollTrack", False)
   ObjGrid.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)
   UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 4800)
   UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
   UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
   UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
   UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3600)
   ObjGrid.RowSizingMode = PropBag.ReadProperty("RowSizingMode", 0)
   ObjGrid.Rows = PropBag.ReadProperty("Rows", 2)
   ObjGrid.Cols(0) = PropBag.ReadProperty("Columns", 3)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.RowPosition(Index) = PropBag.ReadProperty("RowPosition" & Index, 0)
   ObjGrid.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.RowHeight(Index) = PropBag.ReadProperty("RowHeight" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.RowData(Index) = PropBag.ReadProperty("RowData" & Index, 0)
   ObjGrid.RightToLeft = PropBag.ReadProperty("RightToLeft", 0)
   ObjGrid.Redraw = PropBag.ReadProperty("Redraw", True)
   Set Recordset = PropBag.ReadProperty("Recordset", Nothing)
   ObjGrid.PictureType = PropBag.ReadProperty("PictureType", 0)
   UserControl.PaletteMode = PropBag.ReadProperty("PaletteMode", 3)
   Set Palette = PropBag.ReadProperty("Palette", Nothing)
   ObjGrid.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
   ObjGrid.MousePointer = PropBag.ReadProperty("MousePointer", 0)
   Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
   ObjGrid.MergeCells = PropBag.ReadProperty("MergeCells", 0)
   Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
   UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
   UserControl.HitBehavior = PropBag.ReadProperty("HitBehavior", 1)
   ObjGrid.HighLight = PropBag.ReadProperty("HighLight", 1)
   ObjGrid.GridLineWidthUnpopulated = PropBag.ReadProperty("GridLineWidthUnpopulated", 1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.GridLineWidthIndent(BandNumber) = PropBag.ReadProperty("GridLineWidthIndent" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.GridLineWidthHeader(BandNumber) = PropBag.ReadProperty("GridLineWidthHeader" & Index, 0)
   ObjGrid.GridLineWidthFixed = PropBag.ReadProperty("GridLineWidthFixed", 1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.GridLineWidthBand(BandNumber) = PropBag.ReadProperty("GridLineWidthBand" & Index, 0)
   ObjGrid.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
   ObjGrid.GridLinesUnpopulated = PropBag.ReadProperty("GridLinesUnpopulated", 1)
   ObjGrid.GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", 1)
   ObjGrid.GridLines = PropBag.ReadProperty("GridLines", 1)
   ObjGrid.GridColorUnpopulated = PropBag.ReadProperty("GridColorUnpopulated", &H80000010)
   ObjGrid.GridColorFixed = PropBag.ReadProperty("GridColorFixed", &H80000010)
   ObjGrid.GridColor = PropBag.ReadProperty("GridColor", &H80000010)
   ObjGrid.FormatString = PropBag.ReadProperty("FormatString", "")
   ObjGrid.ForeColorSel = PropBag.ReadProperty("ForeColorSel", vbBlack)  '&HFFFFC0)
   ObjGrid.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", 0)  '2147483666#)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.FontWidthHeader(BandNumber) = PropBag.ReadProperty("FontWidthHeader" & Index, 0)
   ObjGrid.FontWidthFixed = PropBag.ReadProperty("FontWidthFixed", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.FontWidthBand(BandNumber) = PropBag.ReadProperty("FontWidthBand" & Index, 0)
   ObjGrid.FontWidth = PropBag.ReadProperty("FontWidth", 0)
   UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
   UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
   UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
   UserControl.FontSize = PropBag.ReadProperty("FontSize", 8.25)
   UserControl.FontName = PropBag.ReadProperty("FontName", "Ms Sans Serif")
   UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
'   Set ObjGrid.FontHeader = PropBag.ReadProperty("FontHeader", Ambient.Font)
   Set ObjGrid.FontFixed = PropBag.ReadProperty("FontFixed", Ambient.Font)
   UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
'   Set ObjGrid.FontBand = PropBag.ReadProperty("FontBand", Ambient.Font)
   ObjGrid.FocusRect = PropBag.ReadProperty("FocusRect", 1)
   ObjGrid.FixedRows = PropBag.ReadProperty("FixedRows", 1)
   ObjGrid.FixedCols = PropBag.ReadProperty("FixedCols", 1)
   ObjGrid.FillStyle = PropBag.ReadProperty("FillStyle", 0)
   UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
   UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
   UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
   UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
   Set DataSource = PropBag.ReadProperty("DataSource", Nothing)
   ObjGrid.DataMember = PropBag.ReadProperty("DataMember", "")
   UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
   UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.ColWordWrapOptionHeader(BandNumber, BandColIndex) = PropBag.ReadProperty("ColWordWrapOptionHeader" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.ColWordWrapOptionFixed(Index) = PropBag.ReadProperty("ColWordWrapOptionFixed" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.ColWordWrapOptionBand(BandNumber, BandColIndex) = PropBag.ReadProperty("ColWordWrapOptionBand" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.ColWordWrapOption(Index) = PropBag.ReadProperty("ColWordWrapOption" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.ColWidth(Index, BandNumber) = PropBag.ReadProperty("ColWidth" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.Cols(BandNumber) = PropBag.ReadProperty("Cols" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.ColPosition(Index, BandNumber) = PropBag.ReadProperty("ColPosition" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.ColHeaderCaption(BandNumber, BandColIndex) = PropBag.ReadProperty("ColHeaderCaption" & Index, "")
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.ColData(Index) = PropBag.ReadProperty("ColData" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.ColAlignmentHeader(BandNumber, BandColIndex) = PropBag.ReadProperty("ColAlignmentHeader" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.ColAlignmentFixed(Index) = PropBag.ReadProperty("ColAlignmentFixed" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.ColAlignmentBand(BandNumber, BandColIndex) = PropBag.ReadProperty("ColAlignmentBand" & Index, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'   ObjGrid.ColAlignment(Index) = PropBag.ReadProperty("ColAlignment" & Index, 0)
   UserControl.ClipControls = PropBag.ReadProperty("ClipControls", True)
   UserControl.ClipBehavior = PropBag.ReadProperty("ClipBehavior", 1)
   ObjGrid.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.BandIndent(BandNumber) = PropBag.ReadProperty("BandIndent" & Index, 0)
   ObjGrid.BandDisplay = PropBag.ReadProperty("BandDisplay", 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
   ObjGrid.BandData(BandData) = PropBag.ReadProperty("BandData" & Index, 0)
'   ObjGrid.BackColorUnpopulated = PropBag.ReadProperty("BackColorUnpopulated", 2147483663#)
   ObjGrid.BackColorSel = PropBag.ReadProperty("BackColorSel", &HFFFFC0)    '2147483661#)
   ObjGrid.BackColorFixed = PropBag.ReadProperty("BackColorFixed", 0) '2147483663#)
'   ObjGrid.BackColorBkg = PropBag.ReadProperty("BackColorBkg", 2147483663#)
   UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
   ObjGrid.Appearance = PropBag.ReadProperty("Appearance", 1)
   ObjGrid.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 0)
   ObjGrid.AllowBigSelection = PropBag.ReadProperty("AllowBigSelection", True)
   Set m_ActiveControl = PropBag.ReadProperty("ActiveControl", Nothing)
   
'   Call DefineImgCab
End Sub

'The Underscore following "PSet" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PSet
Public Sub PSet_(x As Single, Y As Single, Color As Long)
   UserControl.PSet Step(x, Y), Color
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal x As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
   UserControl.PopupMenu Menu, Flags, x, Y, DefaultMenu
End Sub

'The Underscore following "Point" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Point
Public Function Point(x As Single, Y As Single) As Long
Attribute Point.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
   Point = UserControl.Point(x, Y)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,PictureType
Public Property Get PictureType() As PictureTypeSettings
Attribute PictureType.VB_Description = "Returns or sets the type of picture that is generated by the Picture property."
   PictureType = ObjGrid.PictureType
End Property

Public Property Let PictureType(ByVal New_PictureType As PictureTypeSettings)
   ObjGrid.PictureType() = New_PictureType
   PropertyChanged "PictureType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns a picture of the Hierarchical FlexGrid, suitable for printing, saving to disk, copying to the clipboard, or assigning to a different control."
   Set Picture = ObjGrid.Picture
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaletteMode
Public Property Get PaletteMode() As Integer
Attribute PaletteMode.VB_Description = "Returns/sets a value that determines which palette to use for the controls on a object."
   PaletteMode = UserControl.PaletteMode
End Property

Public Property Let PaletteMode(ByVal New_PaletteMode As Integer)
   UserControl.PaletteMode() = New_PaletteMode
   PropertyChanged "PaletteMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Palette
Public Property Get Palette() As Picture
Attribute Palette.VB_Description = "Returns/sets an image that contains the palette to use on an object when PaletteMode is set to Custom"
   Set Palette = UserControl.Palette
End Property

Public Property Set Palette(ByVal New_Palette As Picture)
   Set UserControl.Palette = New_Palette
   PropertyChanged "Palette"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
   UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

Private Sub UserControl_Paint()
   RaiseEvent Paint
End Sub
Private Sub ObjGrid_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns or sets whether this control acts as an OLE drop target."
   OLEDropMode = ObjGrid.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
   ObjGrid.OLEDropMode() = New_OLEDropMode
   PropertyChanged "OLEDropMode"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag and drop event with the given control as the source."
   ObjGrid.OLEDrag
End Sub

Private Sub ObjGrid_OLECompleteDrag(Effect As Long)
   RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,MouseRow
Public Property Get MouseRow() As Long
Attribute MouseRow.VB_Description = "Returns or sets the row or column over which the mouse pointer is positioned. Not available at design time."
   MouseRow = ObjGrid.MouseRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,MousePointer
Public Property Get MousePointer() As MousePointerSettings
Attribute MousePointer.VB_Description = "Returns or sets the type of mouse pointer displayed when over part of an object."
   MousePointer = ObjGrid.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerSettings)
   ObjGrid.MousePointer() = New_MousePointer
   PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns or sets a custom mouse icon."
   Set MouseIcon = ObjGrid.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
   Set ObjGrid.MouseIcon = New_MouseIcon
   PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,MouseCol
Public Property Get MouseCol() As Long
Attribute MouseCol.VB_Description = "Returns or sets the row or column over which the mouse pointer is positioned. Not available at design time."
   MouseCol = ObjGrid.MouseCol
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,MergeCells
Public Property Get MergeCells() As MergeCellsSettings
Attribute MergeCells.VB_Description = "Returns or sets whether cells with the same contents are grouped in a single cell spanning multiple rows or columns."
   MergeCells = ObjGrid.MergeCells
End Property

Public Property Let MergeCells(ByVal New_MergeCells As MergeCellsSettings)
   ObjGrid.MergeCells() = New_MergeCells
   PropertyChanged "MergeCells"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskPicture
Public Property Get MaskPicture() As Picture
Attribute MaskPicture.VB_Description = "Returns/sets the picture that specifies the clickable/drawable area of the control when BackStyle is 0 (transparent)."
   Set MaskPicture = UserControl.MaskPicture
End Property

Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
   Set UserControl.MaskPicture = New_MaskPicture
   PropertyChanged "MaskPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As Long
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
   MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As Long)
   UserControl.MaskColor() = New_MaskColor
   PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Line
'Public Sub Line(ByVal Flags As Integer, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long)
'   UserControl.Line Flags,X1,Y1,X2,Y2,Color
'End Sub

Private Sub ObjGrid_LeaveCell()
   RaiseEvent LeaveCell
   If ObjGrid.SelectionMode <> flexSelectionFree Then
      ObjGrid.CellBackColor = ObjGrid.BackColor
      ObjGrid.CellForeColor = ObjGrid.ForeColor
   End If
End Sub

Private Sub UserControl_InitProperties()
   RaiseEvent InitProperties
'   Call DefineImgCab
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
   Set Image = UserControl.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HyperLink
Public Property Get HyperLink() As HyperLink
Attribute HyperLink.VB_Description = "Returns a Hyperlink object used for browser style navigation."
   Set HyperLink = UserControl.HyperLink
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle to a form or control."
   hWnd = ObjGrid.hWnd
End Property

Private Sub UserControl_HitTest(x As Single, Y As Single, HitResult As Integer)
   RaiseEvent HitTest(x, Y, HitResult)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HitBehavior
Public Property Get HitBehavior() As Integer
Attribute HitBehavior.VB_Description = "Indicates which mode of automatic hit testing a windowless UserControl employs."
   HitBehavior = UserControl.HitBehavior
End Property

Public Property Let HitBehavior(ByVal New_HitBehavior As Integer)
   UserControl.HitBehavior() = New_HitBehavior
   PropertyChanged "HitBehavior"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,HighLight
Public Property Get HighLight() As HighLightSettings
Attribute HighLight.VB_Description = "Returns or sets whether selected cells appear highlighted."
   HighLight = ObjGrid.HighLight
End Property

Public Property Let HighLight(ByVal New_HighLight As HighLightSettings)
   ObjGrid.HighLight() = New_HighLight
   PropertyChanged "HighLight"
End Property

Private Sub UserControl_Hide()
   RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
   hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
   HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLineWidthUnpopulated
Public Property Get GridLineWidthUnpopulated() As Integer
Attribute GridLineWidthUnpopulated.VB_Description = "Returns or sets the width, in Pixels, of the gridlines."
   GridLineWidthUnpopulated = ObjGrid.GridLineWidthUnpopulated
End Property

Public Property Let GridLineWidthUnpopulated(ByVal New_GridLineWidthUnpopulated As Integer)
   ObjGrid.GridLineWidthUnpopulated() = New_GridLineWidthUnpopulated
   PropertyChanged "GridLineWidthUnpopulated"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLineWidthIndent
Public Property Get GridLineWidthIndent(ByVal BandNumber As Long) As Integer
Attribute GridLineWidthIndent.VB_Description = "Returns or sets the width, in Pixels, of the gridlines."
   GridLineWidthIndent = ObjGrid.GridLineWidthIndent(BandNumber)
End Property

Public Property Let GridLineWidthIndent(ByVal BandNumber As Long, ByVal New_GridLineWidthIndent As Integer)
   ObjGrid.GridLineWidthIndent(BandNumber) = New_GridLineWidthIndent
   PropertyChanged "GridLineWidthIndent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLineWidthHeader
Public Property Get GridLineWidthHeader(ByVal BandNumber As Long) As Integer
Attribute GridLineWidthHeader.VB_Description = "Returns or sets the width, in Pixels, of the gridlines."
   GridLineWidthHeader = ObjGrid.GridLineWidthHeader(BandNumber)
End Property

Public Property Let GridLineWidthHeader(ByVal BandNumber As Long, ByVal New_GridLineWidthHeader As Integer)
   ObjGrid.GridLineWidthHeader(BandNumber) = New_GridLineWidthHeader
   PropertyChanged "GridLineWidthHeader"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLineWidthFixed
Public Property Get GridLineWidthFixed() As Integer
Attribute GridLineWidthFixed.VB_Description = "Returns or sets the width, in Pixels, of the gridlines."
   GridLineWidthFixed = ObjGrid.GridLineWidthFixed
End Property

Public Property Let GridLineWidthFixed(ByVal New_GridLineWidthFixed As Integer)
   ObjGrid.GridLineWidthFixed() = New_GridLineWidthFixed
   PropertyChanged "GridLineWidthFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLineWidthBand
Public Property Get GridLineWidthBand(ByVal BandNumber As Long) As Integer
Attribute GridLineWidthBand.VB_Description = "Returns or sets the width, in Pixels, of the gridlines."
   GridLineWidthBand = ObjGrid.GridLineWidthBand(BandNumber)
End Property

Public Property Let GridLineWidthBand(ByVal BandNumber As Long, ByVal New_GridLineWidthBand As Integer)
   ObjGrid.GridLineWidthBand(BandNumber) = New_GridLineWidthBand
   PropertyChanged "GridLineWidthBand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLineWidth
Public Property Get GridLineWidth() As Integer
Attribute GridLineWidth.VB_Description = "Returns or sets the width, in Pixels, of the gridlines."
   GridLineWidth = ObjGrid.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Integer)
   ObjGrid.GridLineWidth() = New_GridLineWidth
   PropertyChanged "GridLineWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLinesUnpopulated
Public Property Get GridLinesUnpopulated() As GridLineSettings
Attribute GridLinesUnpopulated.VB_Description = "Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells."
   GridLinesUnpopulated = ObjGrid.GridLinesUnpopulated
End Property

Public Property Let GridLinesUnpopulated(ByVal New_GridLinesUnpopulated As GridLineSettings)
   ObjGrid.GridLinesUnpopulated() = New_GridLinesUnpopulated
   PropertyChanged "GridLinesUnpopulated"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLinesFixed
Public Property Get GridLinesFixed() As GridLineSettings
Attribute GridLinesFixed.VB_Description = "Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells."
   GridLinesFixed = ObjGrid.GridLinesFixed
End Property

Public Property Let GridLinesFixed(ByVal New_GridLinesFixed As GridLineSettings)
   ObjGrid.GridLinesFixed() = New_GridLinesFixed
   PropertyChanged "GridLinesFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridLines
Public Property Get GridLines() As GridLineSettings
Attribute GridLines.VB_Description = "Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells."
   GridLines = ObjGrid.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As GridLineSettings)
   ObjGrid.GridLines() = New_GridLines
   PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridColorUnpopulated
Public Property Get GridColorUnpopulated() As OLE_COLOR
Attribute GridColorUnpopulated.VB_Description = "Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells."
   GridColorUnpopulated = ObjGrid.GridColorUnpopulated
End Property

Public Property Let GridColorUnpopulated(ByVal New_GridColorUnpopulated As OLE_COLOR)
   ObjGrid.GridColorUnpopulated() = New_GridColorUnpopulated
   PropertyChanged "GridColorUnpopulated"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridColorFixed
Public Property Get GridColorFixed() As OLE_COLOR
Attribute GridColorFixed.VB_Description = "Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells."
   GridColorFixed = ObjGrid.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
   ObjGrid.GridColorFixed() = New_GridColorFixed
   PropertyChanged "GridColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,GridColor
Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells."
   GridColor = ObjGrid.GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
   ObjGrid.GridColor() = New_GridColor
   PropertyChanged "GridColor"
End Property

Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
   RaiseEvent GetDataMember(DataMember, Data)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FormatString
Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "Allows you to set up column widths, alignments, and fixed row and column text for a Hierarchical FlexGrid at design time. See Help for more information."
   FormatString = ObjGrid.FormatString
End Property

Public Property Let FormatString(ByVal New_FormatString As String)
   ObjGrid.FormatString() = New_FormatString
   PropertyChanged "FormatString"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ForeColorSel
Public Property Get ForeColorSel() As OLE_COLOR
Attribute ForeColorSel.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
   ForeColorSel = ObjGrid.ForeColorSel
End Property

Public Property Let ForeColorSel(ByVal New_ForeColorSel As OLE_COLOR)
   ObjGrid.ForeColorSel() = New_ForeColorSel
   PropertyChanged "ForeColorSel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ForeColorFixed
Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "Determines the color used to draw text on each part of the Hierarchical FlexGrid."
   ForeColorFixed = ObjGrid.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
   ObjGrid.ForeColorFixed() = New_ForeColorFixed
   PropertyChanged "ForeColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontWidthHeader
Public Property Get FontWidthHeader(ByVal BandNumber As Long) As Single
Attribute FontWidthHeader.VB_Description = "Returns or sets the default font or the font for individual cells."
   FontWidthHeader = ObjGrid.FontWidthHeader(BandNumber)
End Property

Public Property Let FontWidthHeader(ByVal BandNumber As Long, ByVal New_FontWidthHeader As Single)
   ObjGrid.FontWidthHeader(BandNumber) = New_FontWidthHeader
   PropertyChanged "FontWidthHeader"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontWidthFixed
Public Property Get FontWidthFixed() As Single
Attribute FontWidthFixed.VB_Description = "Returns or sets the width, in points, for the current cell text."
   FontWidthFixed = ObjGrid.FontWidthFixed
End Property

Public Property Let FontWidthFixed(ByVal New_FontWidthFixed As Single)
   ObjGrid.FontWidthFixed() = New_FontWidthFixed
   PropertyChanged "FontWidthFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontWidthBand
Public Property Get FontWidthBand(ByVal BandNumber As Long) As Single
Attribute FontWidthBand.VB_Description = "Returns or sets the default font or the font for individual cells."
   FontWidthBand = ObjGrid.FontWidthBand(BandNumber)
End Property

Public Property Let FontWidthBand(ByVal BandNumber As Long, ByVal New_FontWidthBand As Single)
   ObjGrid.FontWidthBand(BandNumber) = New_FontWidthBand
   PropertyChanged "FontWidthBand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontWidth
Public Property Get FontWidth() As Single
Attribute FontWidth.VB_Description = "Returns or sets the width, in points, for the current cell text."
   FontWidth = ObjGrid.FontWidth
End Property

Public Property Let FontWidth(ByVal New_FontWidth As Single)
   ObjGrid.FontWidth() = New_FontWidth
   PropertyChanged "FontWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
   FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
   UserControl.FontUnderline() = New_FontUnderline
   PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
   FontTransparent = UserControl.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
   UserControl.FontTransparent() = New_FontTransparent
   PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
   FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
   UserControl.FontStrikethru() = New_FontStrikethru
   PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
   FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
   UserControl.FontSize() = New_FontSize
   PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
   FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
   UserControl.FontName() = New_FontName
   PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
   FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
   UserControl.FontItalic() = New_FontItalic
   PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontHeader
Public Property Get FontHeader(ByVal BandNumber As Long) As Font
Attribute FontHeader.VB_Description = "Returns or sets the default font or the font for individual cells."
   Set FontHeader = ObjGrid.FontHeader(BandNumber)
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontFixed
Public Property Get FontFixed() As Font
Attribute FontFixed.VB_Description = "Returns or sets the default font or the font for individual cells."
   Set FontFixed = ObjGrid.FontFixed
End Property

Public Property Set FontFixed(ByVal New_FontFixed As Font)
   Set ObjGrid.FontFixed = New_FontFixed
   PropertyChanged "FontFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
   FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
   UserControl.FontBold() = New_FontBold
   PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FontBand
Public Property Get FontBand(ByVal BandNumber As Long) As Font
Attribute FontBand.VB_Description = "Returns or sets the default font or the font for individual cells."
   Set FontBand = ObjGrid.FontBand(BandNumber)
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FocusRect
Public Property Get FocusRect() As FocusRectSettings
Attribute FocusRect.VB_Description = "Determines whether the Hierarchical FlexGrid control should draw a focus rectangle around the current cell."
   FocusRect = ObjGrid.FocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As FocusRectSettings)
   ObjGrid.FocusRect() = New_FocusRect
   PropertyChanged "FocusRect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FixedRows
Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid."
   FixedRows = ObjGrid.FixedRows
End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)
   ObjGrid.FixedRows() = New_FixedRows
   PropertyChanged "FixedRows"
   Call DefineImgCab
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FixedCols
Public Property Get FixedCols() As Long
Attribute FixedCols.VB_Description = "Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid."
   FixedCols = ObjGrid.FixedCols
End Property

Public Property Let FixedCols(ByVal New_FixedCols As Long)
   ObjGrid.FixedCols() = New_FixedCols
   PropertyChanged "FixedCols"
   Call DefineImgCab
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,FillStyle
Public Property Get FillStyle() As FillStyleSettings
Attribute FillStyle.VB_Description = "Determines whether setting the Text property or one of the cell formatting properties of a Hierarchical FlexGrid applies the change to all selected cells."
   FillStyle = ObjGrid.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleSettings)
   ObjGrid.FillStyle() = New_FillStyle
   PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
   FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
   UserControl.FillColor() = New_FillColor
   PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ExpandAll
Public Sub ExpandAll(ByVal BandNumber As Long)
Attribute ExpandAll.VB_Description = "Expands all rows in specified band or all bands."
   ObjGrid.ExpandAll BandNumber
End Sub

Private Sub ObjGrid_Expand(Cancel As Boolean)
   RaiseEvent Expand(Cancel)
End Sub

Private Sub ObjGrid_EnterCell()
   RaiseEvent EnterCell
   If ObjGrid.SelectionMode <> flexSelectionFree Then
      ObjGrid.CellBackColor = ObjGrid.BackColorSel
      ObjGrid.CellForeColor = ObjGrid.ForeColorSel
   End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns/sets the line width for output from graphics methods."
   DrawWidth = UserControl.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
   UserControl.DrawWidth() = New_DrawWidth
   PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
   DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
   UserControl.DrawStyle() = New_DrawStyle
   PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Sets the appearance of output from graphics methods or of a Shape or Line control."
   DrawMode = UserControl.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
   UserControl.DrawMode() = New_DrawMode
   PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,DataSource
Public Property Get DataSource() As Object
Attribute DataSource.VB_Description = "Returns or sets the data source for the control."
   Set DataSource = ObjGrid.DataSource
End Property

Public Property Set DataSource(ByVal New_DataSource As Object)
   Set ObjGrid.DataSource = New_DataSource
   PropertyChanged "DataSource"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DataMembers
Public Property Get DataMembers() As DataMembers
Attribute DataMembers.VB_Description = "Returns a collection of data members to show at design time for this data source."
   Set DataMembers = UserControl.DataMembers
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub DataMemberChanged(ByVal DataMember As String)
Attribute DataMemberChanged.VB_Description = "Notify data consumers that a data member of this data source has changed."
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,DataMember
Public Property Get DataMember() As String
Attribute DataMember.VB_Description = "Returns or sets the data member for the control."
   DataMember = ObjGrid.DataMember
End Property

Public Property Let DataMember(ByVal New_DataMember As String)
   ObjGrid.DataMember() = New_DataMember
   PropertyChanged "DataMember"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,DataBindings
Public Property Get DataBindings() As DataBindings
Attribute DataBindings.VB_Description = "Returns/sets a DataBindings collection object that collects the bindable properties that are available to the developer."
   Set DataBindings = ObjGrid.DataBindings
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
   CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
   UserControl.CurrentY() = New_CurrentY
   PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
   CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
   UserControl.CurrentX() = New_CurrentX
   PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
Attribute Controls.VB_Description = "A collection whose elements represent each control on a form, including elements of control arrays. "
   Set Controls = UserControl.Controls
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ContainerHwnd
Public Property Get ContainerHwnd() As Long
Attribute ContainerHwnd.VB_Description = "Returns a handle (from Microsoft Windows) to the window a UserControl is contained in."
   ContainerHwnd = UserControl.ContainerHwnd
End Property

Private Sub ObjGrid_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
   RaiseEvent Compare(Row1, Row2, Cmp)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColWordWrapOptionHeader
Public Property Get ColWordWrapOptionHeader(ByVal BandNumber As Long, ByVal BandColIndex As Long) As Integer
Attribute ColWordWrapOptionHeader.VB_Description = "Returns or sets how the text is displayed per column."
   ColWordWrapOptionHeader = ObjGrid.ColWordWrapOptionHeader(BandNumber, BandColIndex)
End Property

Public Property Let ColWordWrapOptionHeader(ByVal BandNumber As Long, ByVal BandColIndex As Long, ByVal New_ColWordWrapOptionHeader As Integer)
   ObjGrid.ColWordWrapOptionHeader(BandNumber, BandColIndex) = New_ColWordWrapOptionHeader
   PropertyChanged "ColWordWrapOptionHeader"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColWordWrapOptionFixed
Public Property Get ColWordWrapOptionFixed(ByVal Index As Long) As Integer
Attribute ColWordWrapOptionFixed.VB_Description = "Returns or sets how the text is displayed per column."
   ColWordWrapOptionFixed = ObjGrid.ColWordWrapOptionFixed(Index)
End Property

Public Property Let ColWordWrapOptionFixed(ByVal Index As Long, ByVal New_ColWordWrapOptionFixed As Integer)
   ObjGrid.ColWordWrapOptionFixed(Index) = New_ColWordWrapOptionFixed
   PropertyChanged "ColWordWrapOptionFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColWordWrapOptionBand
Public Property Get ColWordWrapOptionBand(ByVal BandNumber As Long, ByVal BandColIndex As Long) As Integer
Attribute ColWordWrapOptionBand.VB_Description = "Returns or sets how the text is displayed per column."
   ColWordWrapOptionBand = ObjGrid.ColWordWrapOptionBand(BandNumber, BandColIndex)
End Property

Public Property Let ColWordWrapOptionBand(ByVal BandNumber As Long, ByVal BandColIndex As Long, ByVal New_ColWordWrapOptionBand As Integer)
   ObjGrid.ColWordWrapOptionBand(BandNumber, BandColIndex) = New_ColWordWrapOptionBand
   PropertyChanged "ColWordWrapOptionBand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColWordWrapOption
Public Property Get ColWordWrapOption(ByVal Index As Long) As Integer
Attribute ColWordWrapOption.VB_Description = "Returns or sets how the text is displayed per column."
   ColWordWrapOption = ObjGrid.ColWordWrapOption(Index)
End Property

Public Property Let ColWordWrapOption(ByVal Index As Long, ByVal New_ColWordWrapOption As Integer)
   ObjGrid.ColWordWrapOption(Index) = New_ColWordWrapOption
   PropertyChanged "ColWordWrapOption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColWidth
Public Property Get ColWidth(ByVal Index As Long, ByVal BandNumber As Long) As Long
Attribute ColWidth.VB_Description = "Determines the width of the specified column, in Twips. Not available at design time."
   ColWidth = ObjGrid.ColWidth(Index, BandNumber)
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal BandNumber As Long, ByVal New_ColWidth As Long)
   ObjGrid.ColWidth(Index, BandNumber) = New_ColWidth
   PropertyChanged "ColWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Cols
Public Property Get Cols(ByVal BandNumber As Long) As Long
Attribute Cols.VB_Description = "Determines the total number of columns or rows in the Hierarchical FlexGrid."
   Cols = ObjGrid.Cols(BandNumber)
End Property
Public Property Let Cols(ByVal BandNumber As Long, ByVal New_Cols As Long)
   ObjGrid.Cols(BandNumber) = New_Cols
   PropertyChanged "Cols"
   Call CarregaControles
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColPosition
'Public Property Get ColPosition(ByVal Index As Long, ByVal BandNumber As Long) As Long
'   ColPosition = ObjGrid.ColPosition(Index, BandNumber)
'End Property

Public Property Let ColPosition(ByVal Index As Long, ByVal BandNumber As Long, ByVal New_ColPosition As Long)
Attribute ColPosition.VB_Description = "Returns the distance, in Twips, between the upper left corner of the control and the upper left corner of a specified column."
   ObjGrid.ColPosition(Index, BandNumber) = New_ColPosition
   PropertyChanged "ColPosition"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColPos
Public Property Get ColPos(ByVal Index As Long) As Long
Attribute ColPos.VB_Description = "Returns the distance, in Twips, between the upper left corner of the control and the upper left corner of a specified column."
   ColPos = ObjGrid.ColPos(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,CollapseAll
Public Sub CollapseAll(ByVal BandNumber As Long)
Attribute CollapseAll.VB_Description = "Collapses all rows in specified band or all bands."
   ObjGrid.CollapseAll BandNumber
End Sub

Private Sub ObjGrid_Collapse(Cancel As Boolean)
   RaiseEvent Collapse(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColIsVisible
Public Property Get ColIsVisible(ByVal Index As Long) As Boolean
Attribute ColIsVisible.VB_Description = "Returns True if the specified column is visible."
   ColIsVisible = ObjGrid.ColIsVisible(Index)
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColHeaderCaption
Public Property Get ColHeaderCaption(ByVal BandNumber As Long, ByVal BandColIndex As Long) As String
Attribute ColHeaderCaption.VB_Description = "Returns or sets a band column header caption."
   ColHeaderCaption = ObjGrid.ColHeaderCaption(BandNumber, BandColIndex)
End Property

Public Property Let ColHeaderCaption(ByVal BandNumber As Long, ByVal BandColIndex As Long, ByVal New_ColHeaderCaption As String)
   ObjGrid.ColHeaderCaption(BandNumber, BandColIndex) = New_ColHeaderCaption
   PropertyChanged "ColHeaderCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColData
Public Property Get ColData(ByVal Index As Long) As Long
Attribute ColData.VB_Description = "Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time."
   ColData = ObjGrid.ColData(Index)
End Property

Public Property Let ColData(ByVal Index As Long, ByVal New_ColData As Long)
   ObjGrid.ColData(Index) = New_ColData
   PropertyChanged "ColData"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColAlignmentHeader
Public Property Get ColAlignmentHeader(ByVal BandNumber As Long, ByVal BandColIndex As Long) As Integer
Attribute ColAlignmentHeader.VB_Description = "Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
   ColAlignmentHeader = ObjGrid.ColAlignmentHeader(BandNumber, BandColIndex)
End Property

Public Property Let ColAlignmentHeader(ByVal BandNumber As Long, ByVal BandColIndex As Long, ByVal New_ColAlignmentHeader As Integer)
   ObjGrid.ColAlignmentHeader(BandNumber, BandColIndex) = New_ColAlignmentHeader
   PropertyChanged "ColAlignmentHeader"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColAlignmentFixed
Public Property Get ColAlignmentFixed(ByVal Index As Long) As Integer
Attribute ColAlignmentFixed.VB_Description = "Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
   ColAlignmentFixed = ObjGrid.ColAlignmentFixed(Index)
End Property

Public Property Let ColAlignmentFixed(ByVal Index As Long, ByVal New_ColAlignmentFixed As Integer)
   ObjGrid.ColAlignmentFixed(Index) = New_ColAlignmentFixed
   PropertyChanged "ColAlignmentFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColAlignmentBand
Public Property Get ColAlignmentBand(ByVal BandNumber As Long, ByVal BandColIndex As Long) As Integer
Attribute ColAlignmentBand.VB_Description = "Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
   ColAlignmentBand = ObjGrid.ColAlignmentBand(BandNumber, BandColIndex)
End Property

Public Property Let ColAlignmentBand(ByVal BandNumber As Long, ByVal BandColIndex As Long, ByVal New_ColAlignmentBand As Integer)
   ObjGrid.ColAlignmentBand(BandNumber, BandColIndex) = New_ColAlignmentBand
   PropertyChanged "ColAlignmentBand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ColAlignment
Public Property Get ColAlignment(ByVal Index As Long) As Integer
Attribute ColAlignment.VB_Description = "Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString property)."
   ColAlignment = ObjGrid.ColAlignment(Index)
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal New_ColAlignment As Integer)
   ObjGrid.ColAlignment(Index) = New_ColAlignment
   PropertyChanged "ColAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
   UserControl.Cls
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ClipControls
Public Property Get ClipControls() As Boolean
Attribute ClipControls.VB_Description = "Determines whether graphics methods in Paint events repaint an entire object or newly exposed areas."
   ClipControls = UserControl.ClipControls
End Property

Public Property Let ClipControls(ByVal New_ClipControls As Boolean)
   UserControl.ClipControls() = New_ClipControls
   PropertyChanged "ClipControls"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ClipBehavior
Public Property Get ClipBehavior() As Integer
Attribute ClipBehavior.VB_Description = "Indicates the manner in which a windowless UserControl's appearance is clipped."
   ClipBehavior = UserControl.ClipBehavior
End Property

Public Property Let ClipBehavior(ByVal New_ClipBehavior As Integer)
   UserControl.ClipBehavior() = New_ClipBehavior
   PropertyChanged "ClipBehavior"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,ClearStructure
Public Sub ClearStructure()
Attribute ClearStructure.VB_Description = "Clears information about the order and name of columns displayed."
   ObjGrid.ClearStructure
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of the Hierarchical FlexGrid. This includes all text, pictures, and cell formatting."
   ObjGrid.Clear
End Sub

'The Underscore following "Circle" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Circle
Public Sub Circle_(x As Single, Y As Single, Radius As Single, Color As Long, StartPos As Single, EndPos As Single, Aspect As Single)
   UserControl.Circle (x, Y), Radius, Color, StartPos, EndPos, Aspect
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,CellWidth
Public Property Get CellWidth() As Long
Attribute CellWidth.VB_Description = "Returns the width of the current cell, in twips."
   CellWidth = ObjGrid.CellWidth
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,CellTop
Public Property Get CellTop() As Long
Attribute CellTop.VB_Description = "Returns the top position of the current cell, in twips."
   CellTop = ObjGrid.CellTop
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,CellLeft
Public Property Get CellLeft() As Long
Attribute CellLeft.VB_Description = "Returns the left position of the current cell, in twips."
   CellLeft = ObjGrid.CellLeft
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,CellHeight
Public Property Get CellHeight() As Long
Attribute CellHeight.VB_Description = "Returns the height of the current cell, in Twips."
   CellHeight = ObjGrid.CellHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,CausesValidation
Public Property Get CausesValidation() As Boolean
Attribute CausesValidation.VB_Description = "Returns/sets whether validation occurs on the control which lost focus."
   CausesValidation = ObjGrid.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
   ObjGrid.CausesValidation() = New_CausesValidation
   PropertyChanged "CausesValidation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function CanPropertyChange(ByVal PropertyName As String) As Boolean
Attribute CanPropertyChange.VB_Description = "Asks the container if a property bound to a data source can be changed.  The CanPropertyChange method is most useful if the property specified in PropertyName is bound to a data source."

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CancelAsyncRead
Public Sub CancelAsyncRead(Optional ByVal Property As Variant)
Attribute CancelAsyncRead.VB_Description = "Cancel an asynchronous data request."
   UserControl.CancelAsyncRead Property
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Bands
Public Property Get Bands() As Long
Attribute Bands.VB_Description = "Returns the number of bands."
   Bands = ObjGrid.Bands
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BandIndent
Public Property Get BandIndent(ByVal BandNumber As Long) As Long
Attribute BandIndent.VB_Description = "Returns or sets the indent for a band."
   BandIndent = ObjGrid.BandIndent(BandNumber)
End Property

Public Property Let BandIndent(ByVal BandNumber As Long, ByVal New_BandIndent As Long)
   ObjGrid.BandIndent(BandNumber) = New_BandIndent
   PropertyChanged "BandIndent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BandDisplay
Public Property Get BandDisplay() As BandDisplaySettings
Attribute BandDisplay.VB_Description = "Returns or sets the band display style."
   BandDisplay = ObjGrid.BandDisplay
End Property

Public Property Let BandDisplay(ByVal New_BandDisplay As BandDisplaySettings)
   ObjGrid.BandDisplay() = New_BandDisplay
   PropertyChanged "BandDisplay"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BandData
'Public Property Get BandData(ByVal BandData As Long) As Long
'   BandData = ObjGrid.BandData(BandData)
'End Property

Public Property Let BandData(ByVal BandData As Long, ByVal New_BandData As Long)
Attribute BandData.VB_Description = "Returns or sets a user-determined long value associated with each band."
   ObjGrid.BandData(BandData) = New_BandData
   PropertyChanged "BandData"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BackColorUnpopulated
Public Property Get BackColorUnpopulated() As OLE_COLOR
Attribute BackColorUnpopulated.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
   BackColorUnpopulated = ObjGrid.BackColorUnpopulated
End Property

Public Property Let BackColorUnpopulated(ByVal New_BackColorUnpopulated As OLE_COLOR)
   ObjGrid.BackColorUnpopulated() = New_BackColorUnpopulated
   PropertyChanged "BackColorUnpopulated"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BackColorSel
Public Property Get BackColorSel() As OLE_COLOR
Attribute BackColorSel.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
   BackColorSel = ObjGrid.BackColorSel
End Property

Public Property Let BackColorSel(ByVal New_BackColorSel As OLE_COLOR)
   ObjGrid.BackColorSel() = New_BackColorSel
   PropertyChanged "BackColorSel"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BackColorFixed
Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
   BackColorFixed = ObjGrid.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
   ObjGrid.BackColorFixed() = New_BackColorFixed
   PropertyChanged "BackColorFixed"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,BackColorBkg
Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Returns or sets the background color of various elements of the Hierarchical FlexGrid."
   BackColorBkg = ObjGrid.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
   ObjGrid.BackColorBkg() = New_BackColorBkg
   PropertyChanged "BackColorBkg"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
   AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
   UserControl.AutoRedraw() = New_AutoRedraw
   PropertyChanged "AutoRedraw"
End Property

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
   RaiseEvent AsyncReadProgress(AsyncProp)
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
   RaiseEvent AsyncReadComplete(AsyncProp)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AsyncRead
Public Sub AsyncRead(ByVal Target As String, ByVal AsyncType As Long, Optional ByVal PropertyName As Variant, Optional ByVal AsyncReadOptions As Variant)
Attribute AsyncRead.VB_Description = "Read in data asynchronously from a path or a URL and receive AsyncReadComplete event."
   UserControl.AsyncRead Target, AsyncType, PropertyName, AsyncReadOptions
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,Appearance
Public Property Get Appearance() As AppearanceSettings
Attribute Appearance.VB_Description = "Returns or sets whether a control should be painted with 3-D effects."
   Appearance = ObjGrid.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceSettings)
   ObjGrid.Appearance() = New_Appearance
   PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,AllowUserResizing
Public Property Get AllowUserResizing() As AllowUserResizeSettings
Attribute AllowUserResizing.VB_Description = "Returns or sets whether the user is allowed to resize rows and columns with the mouse."
   AllowUserResizing = ObjGrid.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
   ObjGrid.AllowUserResizing() = New_AllowUserResizing
   PropertyChanged "AllowUserResizing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,AllowBigSelection
Public Property Get AllowBigSelection() As Boolean
Attribute AllowBigSelection.VB_Description = "Returns or sets whether clicking on a column or row header causes the entire column or row to be selected."
   AllowBigSelection = ObjGrid.AllowBigSelection
End Property

Public Property Let AllowBigSelection(ByVal New_AllowBigSelection As Boolean)
   ObjGrid.AllowBigSelection() = New_AllowBigSelection
   PropertyChanged "AllowBigSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ObjGrid,ObjGrid,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds a new row to a Hierarchical FlexGrid control at run time."
   ObjGrid.AddItem Item, Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=15,0,0,0
Public Property Get ActiveControl() As Object
Attribute ActiveControl.VB_Description = "Returns the control that has focus."
   Set ActiveControl = m_ActiveControl
End Property
Private Sub DefineImgCab()
   Dim i    As Integer
   Dim j    As Integer
   If ObjGrid.FixedRows > 0 Or ObjGrid.FixedCols > 0 Then
      For j = 0 To ObjGrid.Rows - 1
         For i = 0 To ObjGrid.Cols - 1
            If j < ObjGrid.FixedRows Or i < ObjGrid.FixedCols Then
               ObjGrid.Row = j
               ObjGrid.Col = i
               
               Dim x
               x = "PctAzulClaro"
               Select Case x
                  Case "PctAzulClaro": Set ObjGrid.CellPicture = PctAzulClaro.Picture
                  Case "PctAzulEscuro": Set ObjGrid.CellPicture = PctAzulEscuro.Picture
                  Case "PctCinzaClaro": Set ObjGrid.CellPicture = PctCinzaClaro.Picture
                  Case "PctCinzaEscuro": Set ObjGrid.CellPicture = PctCinzaEscuro.Picture
                  Case "PctOlivaClaro": Set ObjGrid.CellPicture = PctOlivaClaro.Picture
                  Case "PctOlivaEscuro": Set ObjGrid.CellPicture = PctOlivaEscuro.Picture
               End Select
            End If
         Next
      Next
   End If
End Sub
Private Sub CarregaControles()
   Dim i As Integer
   On Error Resume Next
   For i = 1 To ObjGrid.Cols - 1
      If i <> 0 Then
         Load CellTextBox(i)
         CellTextBox(i).Top = CellTextBox(i - 1).Top + CellTextBox(i - 1).Height
         CellTextBox(i).Text = "Text" & i
         'CellTextBox(i).ZOrder 1
      End If
   Next
End Sub
