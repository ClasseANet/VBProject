Attribute VB_Name = "basChooseColor"



'Date Created:
'Last Updated:

Option Explicit
DefInt A-Z

Private Type TCHOOSECOLOR
 lStructSize        As Long
 hWndOwner          As Long
 hInstance          As Long
 rgbResult          As Long
 lpCustColors       As Long
 Flags              As Long
 lCustData          As Long
 lpfnHook           As Long
 lpTemplateName     As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public CustomColors(0 To 15) As Long

Public Function SelectColor(hWndParent As Long, DefColor As Long, Optional ShowExpDlg As Boolean = 0, Optional InitCustomColours As Boolean = -1) As Long
 Dim i
 Dim c As Long
 Dim CC As TCHOOSECOLOR
 Dim CT$
 'Initialise Custom Colours
 If InitCustomColours Then
  For i = 0 To 15
   CT$ = GetSetting$("Ariad Non-ADL User Settings", "CustomColours", CStr(i))
   CustomColors(i) = IIf(Len(CT$), Val(CT$), QBColor(15))
  Next
 End If
 'Show Dialog
 With CC
  .rgbResult = DefColor
  .hWndOwner = hWndParent
  .lpCustColors = VarPtr(CustomColors(0))
  .Flags = &H101
  If ShowExpDlg Then .Flags = .Flags Or &H2
  .lStructSize = Len(CC)
  c = ChooseColor(CC)
  If c Then
   SelectColor = .rgbResult
  Else
   SelectColor = -1
  End If
 End With
End Function
