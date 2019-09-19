Attribute VB_Name = "basFileDialogs"

'-------------------------------'
' R Development Library 2.0 '
'-------------------------------'
'       API File Common Dialogs '
'                   Version 1.0 '
'-------------------------------'
'Copyright © 1998-9 by R Software. All Rights Reserved
'
'Date Created:
'Last Updated:

Option Explicit
DefInt A-Z

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 128

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST

Type OPENFILENAME
 lStructSize        As Long
 hWndOwner          As Long
 hInstance          As Long
 lpstrFilter        As String
 lpstrCustomFilter  As String
 nMaxCustFilter     As Long
 nFilterIndex       As Long
 lpstrFile          As String
 nMaxFile           As Long
 lpstrFileTitle     As String
 nMaxFileTitle      As Long
 lpstrInitialDir    As String
 lpstrTitle         As String
 Flags              As Long
 nFileOffset        As Integer
 nFileExtension     As Integer
 lpstrDefExt        As String
 lCustData          As Long
 lpfnHook           As Long
 lpTemplateName     As String
End Type

Public OFN As OPENFILENAME

Public Declare Function CommDlgExtendedError Lib "COMDLG32.DLL" () As Long
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Enum FileDlgModes
 fdmOpenFile = 1
 fdmSaveFile
 fdmSaveFileNoConfirm
 fdmOpenFileOrPrompt
End Enum

Public Const PIC_FILTER1$ = "Pictures (*.bmp;*.dib;*.ico;*.gif;*.jpg;*.rle)|*.bmp;*.dib;*.ico;*.gif;*.jpg;*.rle|Bitmaps (*.bmp;*.dib;*.rle)|*.bmp;*.dib;*.rle|Icons (*.ico)|*.ico|Internet Images (*.gif;*.jpg)|*.gif;*.jpg"

Public Function SelectFile$(OwnerHWnd As Long, Optional Title$ = "", Optional Filter$ = "All Files (*.*)|*.*", Optional FilterIDX As Long = 0, Optional DefFile$, Optional DefPath$, Optional DefExt$, Optional ByVal FileMode As FileDlgModes = fdmOpenFile)
 Dim R As Long, SP As Long, ShortSize As Long, Z As Long
 With OFN
  .lStructSize = Len(OFN)
  .hWndOwner = OwnerHWnd
  .hInstance = App.hInstance
  .lpstrFilter = Replace$(Filter$, "|", Chr$(0)) & Chr$(0)
  .nFilterIndex = FilterIDX
  .lpstrFile = DefFile$ & String$(257 - Len(DefFile$), 0)
  .nMaxFile = Len(.lpstrFile) - 1
  .lpstrFileTitle = .lpstrFile
  .nMaxFileTitle = .nMaxFile
  .lpstrDefExt = DefExt$ & Chr$(0)
  .lpstrInitialDir = IIf(Len(DefPath$), DefPath$, CurDir$) & Chr$(0)
  .lpstrTitle = Title$ & Chr$(0)
  If FileMode = fdmSaveFile Or FileMode = fdmSaveFileNoConfirm Then
   .Flags = OFS_FILE_SAVE_FLAGS
   If FileMode = fdmSaveFile Then .Flags = .Flags Or OFN_OVERWRITEPROMPT
   R = GetSaveFileName(OFN)
  Else
   .Flags = OFS_FILE_OPEN_FLAGS
   If FileMode = fdmOpenFileOrPrompt Then .Flags = .Flags Or OFN_CREATEPROMPT
   R = GetOpenFileName(OFN)
  End If
  If R Then
   SP = InStr(.lpstrFile, Chr$(0))
   If SP Then .lpstrFile = Left$(.lpstrFile, SP - 1)
   SelectFile$ = Trim$(Replace$(.lpstrFile, Chr$(0), ""))
  Else
   Z = CommDlgExtendedError()
   If Z Then MsgBox "Unable to get filename(s)." & vbCr & vbCr & "CommDlgExtendedError returned " & Z, vbCritical
  End If
 End With
End Function
