VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
''' Creates a new pidl of the given size
''' (calling proc is responsible for freeing the new pidl)
'''Public Function CreatePIDL(cb As Long) As Long
'''  Dim pidl As Long
'''  pidl = isMalloc.Alloc(cb)
'''  If pidl Then
'''    FillMemory ByVal pidl, cb, 0 ' initialize to zero, set by caller
'''    CreatePIDL = pidl
'''  End If
'''End Function
''
''' Returns a copy of a relative or absolute pidl
''' (calling proc is responsible for freeing the new pidl)
'''Public Function CopyPIDL(pidl As Long) As Long
'''  Dim cb As Long
'''  Dim pidlNew As Long
'''  cb = GetPIDLSize(pidl)
'''  If cb Then
'''    pidlNew = CreatePIDL(cb)
'''    MoveMemory ByVal pidlNew, ByVal pidl, cb
'''    CopyPIDL = pidlNew
'''  End If
'''End Function

'Public Function BrowseDialog(hWnd As Long, sPrompt As String, ulFlags As BF_Flags, Optional pidlRoot As Long = 0, Optional pidlPreSel As Long = 0) As Long
Public Function BrowseDialog(hWnd As Long, sPrompt As String, ulFlags As Long, Optional pidlRoot As Long = 0, Optional pidlPreSel As Long = 0) As Long
  Dim bi As BROWSEINFO
  With bi
    .hwndOwner = hWnd
    .pidlRoot = pidlRoot
    .lpszTitle = sPrompt
    .ulFlags = ulFlags
    .lParam = pidlPreSel
'    .lpfn = FARPROC(AddressOf BrowseCallbackProc)
  End With
  BrowseDialog = SHBrowseForFolder(bi)
End Function

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                                   ByVal lParam As Long, ByVal lpData As Long) As Long
'  Dim sPath As String * MAX_PATH
  
  Select Case uMsg
    
    Case BFFM_INITIALIZED
    Dim BFFM_SETSELECTIONA, cFalse
      ' Set the dialog's pre-selected folder from the pidl we set
      ' bi.lParam to above (passed in the lpData param).
      Call SendMessage(hWnd, BFFM_SETSELECTIONA, ByVal cFalse, ByVal lpData)
      
'    Case BFFM_SELCHANGED
'      If SHGetPathFromIDList(lParam, sPath) Then
'        ' Return the path
'        Debug.Print Left$(sPath, InStr(sPath, vbNullChar) - 1)
'      End If
    
  End Select

End Function

' Fills a GUID
Public Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function
' Returns a pointer to the next Item ID in a pidl.
' Returns 0 if the next Item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer   ' SHItemID.cb, 2 bytes
  cb = GetItemIDSize(pidl)
  ' Make sure it's not the zero value terminator.
  If cb Then GetNextItemID = pidl + cb
End Function
' Frees the specified pidl and zeros it
' Copies and returns the specified Item ID from a complex pidl
'   pidl -    pointer to an Item ID list from which to copy
'   nItem - 1-based position in the pidl of the Item ID to copy

' If successful, returns a new Item ID (single-element pidl)
' from the specified element positon. Returns 0 on failure.
' If nItem exceeds the number of Item IDs in the pidl,
' the last Item ID is returned.
' (calling proc is responsible for freeing the new pidl)

'Public Function GetItemID(ByVal pidl As Long, ByVal nItem As Integer) As Long
'  Dim nCount As Integer
'  Dim i As Integer
'  Dim cb As Integer
'  Dim pidlNew As Long
'
'  nCount = GetItemIDCount(pidl)
'  If (nItem > nCount) Or (nItem = GIID_LAST) Then nItem = nCount
'
'  ' GetNextItemID returns the 2nd Item ID
'  For i = 1 To nItem - 1: pidl = GetNextItemID(pidl): Next
'
'  ' Get the size of the specified Item identifier.
'  ' If cb = 0 (the zero terminator), the we'll return a desktop pidl, proceed
'  cb = GetItemIDSize(pidl)
'
'  ' Allocate a new Item identifier list.
'  pidlNew = isMalloc.Alloc(cb + 2)
'  If pidlNew Then
'
'    ' Copy the specified Item identifier.
'    ' and append the zero terminator.
'    MoveMemory ByVal pidlNew, ByVal pidl, cb
'    MoveMemory ByVal pidlNew + cb, 0, 2
'
'    GetItemID = pidlNew
'  End If
'
'End Function
'Public Function FreePIDL(pidl As Long) As Boolean
'  On Error GoTo Out
'  If pidl Then
'    ' Free the pidl and zero it's *value* only
'    ' (not what it points to!, i.e. ZeroMemory = FE...)
'    isMalloc.Free ByVal pidl: pidl = 0
'Out:
'    If Err Then
'      Call CoTaskMemFree(pidl): pidl = 0
'    End If
'    FreePIDL = True
'  End If
'End Function

Public Function GetAPIErrStr(dwErrCode As Long) As String
  Dim sErrDesc As String * 256   ' max string resource len
  Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS Or _
                                 FORMAT_MESSAGE_MAX_WIDTH_Mask, _
                                 ByVal 0&, dwErrCode, LANG_USER_DEFAULT, _
                                 ByVal sErrDesc, 256, 0)
  GetAPIErrStr = GetStrFromBufferA(sErrDesc)
End Function
Public Function MyPtInRectRegion(pLeft As Long, pTop As Long, pRight As Long, pBotton As Long, AutoScroll As Long, X As Long, Y As Long) As RectFlags
   Dim lpRect As RECT
   Dim lpPoint As PointAPI
   lpPoint.X = X
   lpPoint.Y = Y
   
   lpRect.Left = pLeft
   lpRect.Top = pTop
   lpRect.Right = pRight
   lpRect.Bottom = pBotton
'*******************
   MyPtInRectRegion = PtInRectRegion(lpRect, AutoScroll, lpPoint)
'*******************
   X = lpPoint.X
   Y = lpPoint.Y
   pLeft = lpRect.Left
   pTop = lpRect.Top
   pRight = lpRect.Right
   pBotton = lpRect.Bottom
End Function
Public Function MyGetClientRect(ByVal hWnd As Long, pLeft As Long, pTop As Long, pRight As Long, pBotton As Long) As Long 'Rect As RECT) As Long
   Dim lpRect As RECT
   lpRect.Left = pLeft
   lpRect.Top = pTop
   lpRect.Right = pRight
   lpRect.Bottom = pBotton
'*******************
   MyGetClientRect = GetClientRect(hWnd, lpRect)
'*******************
   pLeft = lpRect.Left
   pTop = lpRect.Top
   pRight = lpRect.Right
   pBotton = lpRect.Bottom
End Function
Public Function MyGetScrollPos(ByVal hWnd As Long, ByVal nBar As Long) As Long
   MyGetScrollPos = GetScrollPos(hWnd, nBar)
End Function
Public Function MyScreenToClient(ByVal hWnd As Long, ByRef X As Long, ByRef Y As Long) As Long   ' lpPoint As POINTAPI) As Long
   Dim lpPoint As PointAPI
   lpPoint.X = X
   lpPoint.Y = Y
'*******************
   MyScreenToClient = ScreenToClient(hWnd, lpPoint)
'*******************
   X = lpPoint.X
   Y = lpPoint.Y
End Function
Public Function MySendMessage(ByVal hWnd As Long, ByVal wMsg As Long, wParam As Variant, lParam As Variant) As Long
   Call SendMessage(hWnd, wMsg, wParam, lParam)
End Function
Public Sub MyUpdateWindow(ByVal hWnd As Long)
   Call UpdateWindow(hWnd)
End Sub
Public Function MyReleaseCapture() As Long
   Call ReleaseCapture
End Function
Public Function MyImageList_BeginDrag(ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As cBoolean
   MyImageList_BeginDrag = ImageList_BeginDrag(himlTrack, iTrack, dxHotspot, dyHotspot)
End Function
Public Function MyImageList_DragEnter(ByVal hwndLock As Long, ByVal X As Long, ByVal Y As Long) As cBoolean
   MyImageList_DragEnter = ImageList_DragEnter(hwndLock, X, Y)
End Function
Public Function MyImageList_DragLeave(ByVal hwndLock As Long)
   MyImageList_DragLeave = ImageList_DragLeave(hwndLock)
End Function
Public Function MyImageList_DragMove(ByVal X As Long, ByVal Y As Long) As cBoolean
   MyImageList_DragMove = ImageList_DragMove(X, Y)
End Function
Public Function MyImageList_DragShowNolock(ByVal fShow As Boolean) As cBoolean
   MyImageList_DragShowNolock = ImageList_DragShowNolock(fShow)
End Function
Public Function MyImageList_Destroy(ByVal himl As Long) As cBoolean
   MyImageList_Destroy = ImageList_Destroy(himl)
End Function
Public Sub MyImageList_EndDrag()
   Call ImageList_EndDrag
End Sub
Public Function MyImageList_GetImageCount(ByVal himl As Long) As Long
   MyImageList_GetImageCount = ImageList_GetImageCount(himl)
End Function
Public Function MyImageList_GetIconSize(ByVal himl As Long, lpcx As Long, lpcy As Long) As Boolean
   MyImageList_GetIconSize = ImageList_GetIconSize(himl, lpcx, lpcy)
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.
Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Returns a file's SFGAO_ attributes
'   pszPath  - must be either an absolute path or absolute pidl
Public Function GetFileAttribs(ByVal pszPath As Variant) As Long
  Dim sfi As SHFILEINFO
  If GetFileInfo(pszPath, SHGFI_ATTRIBUTES, sfi) Then
    GetFileAttribs = sfi.dwAttributes
  End If
End Function
' Rtns True on success, False on failure
' Adapted as a more efficient GetFileDisplayName()...
Public Function GetFileDisplayNameEx(pszPath As Long, sName As String) As Boolean
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(pszPath, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_DISPLAYNAME) Then
    sName = sfi.szDisplayName   ' no GetStrFromBufferA()
    GetFileDisplayNameEx = True
  End If
  ' Returns the IShellFolder for the specified relative pidl
  '   isfParent - pidl's parent folder IShellFolder
  '   pidlRel    - child folder's relative pidl we're returning the IShellFolder of.
  ' If an error occurs, the desktop's IShellFolder is returned.
End Function
Public Function GetItemIDSize(ByVal pidl As Long) As Integer
   'If we try to access memory at address 0 (NULL), then it's bye-bye...
   If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2
End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.
Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer
  ' Error handle in case we get a bad pidl and overflow cb.
  ' (most Item IDs are roughly 20 bytes in size, and since an Item ID represents
  ' a folder, a pidl can never exceed 260 folders, or 5200 bytes).
  On Error GoTo Out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop
    ' Add 2 bytes for the zero terminating Item ID
    GetPIDLSize = cb + 2
  End If
  
Out:
End Function
Public Function GetPath(CmdDlghWnd As Long, ByRef ObjTxt As Variant)
   Dim IdRoot As Long
   Dim Folder$
   IdRoot = BrowseDialog(CmdDlghWnd, "Selecione o Diretório.", &H1)
   If IdRoot Then
      Folder$ = Space$(260)
      If SHGetPathFromIDList(IdRoot, Folder$) Then
         Folder$ = Left(Folder$, InStr(Folder$, Chr$(0)) - 1)
      Else
         Folder$ = ""
      End If
   End If
   Select Case UCase(TypeName(ObjTxt))
      Case "TEXTBOX": ObjTxt.Text = IIf(Trim(Folder$) = "", ObjTxt.Text, UCase(Folder$))
      Case Else: ObjTxt = IIf(Trim(Folder$) = "", ObjTxt, UCase(Folder$))
   End Select
   GetPath = Folder
End Function
'Public Function GetPIDLParent(pidl As Long, Optional fReturnDesktop As Boolean = False, _
'                              Optional fFreeOldPidl As Boolean = False) As Long
'  Dim nCount As Integer
'  Dim pidl1 As Long
'  Dim i As Integer
'  Dim cb As Integer
'  Dim pidlNew As Long
'
'  nCount = GetItemIDCount(pidl)
'  If (nCount = 0) And (fReturnDesktop = False) Then Exit Function
'
'  ' Get the size of all but the pidl's last Item ID and zero terminator.
'  ' (maintain the value of the original pidl, it's passed ByRef !!)
'  pidl1 = pidl
'  For i = 1 To nCount - 1
'    cb = cb + GetItemIDSize(pidl1)
'    pidl1 = GetNextItemID(pidl1)
'  Next
'
'  ' Allocate a new Item ID list with a new terminating 2 bytes.
'  pidlNew = isMalloc.Alloc(cb + 2)
'
'  ' If the memory was allocated...
'  If pidlNew Then
'    ' Copy all but the last Item ID from the original pidl
'    ' to the new pidl and zero the terminating 2 bytes.
'    MoveMemory ByVal pidlNew, ByVal pidl, cb
'    FillMemory ByVal pidlNew + cb, 2, 0
'
'    If fFreeOldPidl Then Call FreePIDL(pidl)
'    GetPIDLParent = pidlNew
'
'  End If
'
'End Function
' Returns the count of Item IDs in a pidl.

Public Function GetItemIDCount(ByVal pidl As Long) As Integer
  Dim nItems As Integer
  ' If the size of an Item ID is 0, then it's the zero
  ' value terminating Item ID at the end of the pidl.
  Do While GetItemIDSize(pidl)
    pidl = GetNextItemID(pidl)
    nItems = nItems + 1
  Loop
  GetItemIDCount = nItems
End Function

' =====================================================
' interface procs

' Returns a reference to the IMalloc interface.

'Public Function isMalloc() As IMalloc
'  Static im As IMalloc
'  If (im Is Nothing) Then Call SUCCEEDED(SHGetMalloc(im))
'  Set isMalloc = im
'End Function
' Inserts a new root Item in the treeview control.
'   hwndTV  - treeview's hWnd
'   isfParent - pointer to the root's parent shell folder
'   pidlFQ    - root Item's fully qualified pidl

' Rtns True if successful, False if failure
' Called from Form1.FillControls() & Form1.mnuFileRoot_Click()

'Public Function InsertTVRoot(hwndTV As Long, isfParent As IShellFolder, pidlFQ As Long) As Long
'  Dim lptvid As Long   ' pointer to the TVItemDATA struct
'  Dim tvid As TVItemDATA
'  Dim tvi As TVItem
'  Dim sName As String       ' the Item's displayname
'  Dim ulAttrs As Long
'  Dim tvins As TVINSERTSTRUCT
'  Dim hItemRoot  As Long
'
''  Call EmptyTreeView(hwndTV, True)
'
'  ' Allocate memory & get a pointer to our ItemDATA struct
'  lptvid = isMalloc.Alloc(Len(tvid))
'  If lptvid = 0 Then Exit Function
'
'  ' Indicate what TVItem members contain data
'  tvi.Mask = TVIF_CHILDREN Or TVIF_TEXT Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE Or TVIF_PARAM
'  tvi.cChildren = 1
'
'  ' Store the pointer to the root's friendly name that we'll put in the treeview...
'  Call GetFileDisplayNameEx(pidlFQ, sName) '& vbNullChar
'  tvi.pszText = StrPtrEx(sName)
'  tvi.cchTextMax = MAX_PATH
'
''**************************************8
'  ' Store the root's normal & selected icons in the
'  ' iImage & iSelectedImage members respectively.
'  ulAttrs = GetFileAttribs(pidlFQ)
'  If (ulAttrs And SFGAO_SHARE) Then tvi.Mask = tvi.Mask Or TVIF_State
''  Call FillTVItemIcons(pidlFQ, tvi, (ulAttrs And SFGAO_SHARE))
'
'  ' Fill our Item data struct
'  Set tvid.isfParent = isfParent
'  ' ======================================
'  ' This is crucial!!! We must increment the parent
'  ' folder's reference count
'  If (isfParent Is Nothing) = False Then Call isfParent.AddRef
'  ' ======================================
'  tvid.pidlRel = GetItemID(pidlFQ, GIID_LAST)
'  ' Make a copy in case the calling proc frees it.
'  tvid.pidlFQ = CopyPIDL(pidlFQ)
'
'  ' Fill & store the pointer to the Item's ItemDATA struct
'  MoveMemory ByVal lptvid, tvid, Len(tvid)
'  tvi.lParam = lptvid
'
'  ' Populate the TreeVeiw Insert Struct.
'Dim TVI_ROOT
'  tvins.hParent = TVI_ROOT
'  tvins.hInsertAfter = 0   ' is root Item
'  tvins.Item = tvi
'
'  ' Add the root Item to the treeview, returning it's hItem
''  hItemRoot = TreeView_InsertItem(hwndTV, tvins)
'  If hItemRoot Then
'
'    ' Expand the root, invoking a TVN_ItemEXPANDING,
'    ' which calls FillTreeView() for the root's subfolders.
''    Call TreeView_Expand(hwndTV, hItemRoot, TVE_EXPAND)
''    Call UpdateWindow(hwndTV)
'
'    ' Select the root, invoking a TVN_SELCHANGING,
'    ' which calls FillListView() for the root's contents.
''    Call TreeView_SelectItem(hwndTV, hItemRoot)
'
'    InsertTVRoot = hItemRoot
'  End If
'
'End Function
'Public Function isfDesktop() As IShellFolder
'  Static isf As IShellFolder
'  If (isf Is Nothing) Then Call SUCCEEDED(SHGetDesktopFolder(isf))
'  Set isfDesktop = isf
'End Function
' Provides a generic test for success on any status value.
' Non-negative numbers indicate success.
' Returns the memory address of any string, ANSII or Unicode.
' (VB's StrPtr() function only rtns the address of a Unicode string)

' This is scetchy. If the sz param isn't static, or if this function's
' return value isn't used real fast, VB may move the string,
' invalidating the string's pointer...

Public Function StrPtrEx(sz As String) As Long
  MoveMemory StrPtrEx, sz, 4
End Function

' If we incur any error situation from any API or interface member
' function's call to this proc, we'll let the user know that sometime's
' not right. What happens when execution continues after the error
' is indeternimate and could possibly lead to an FE (i.e. if we don't
' get isfDesktop, then the control's dead in the water -useless).

Public Function SUCCEEDED(hr As Long) As Boolean   ' hr = HRESULT
Dim S_OK
  If (hr >= S_OK) Then
    SUCCEEDED = True
  Else
    MsgBox "Error: &H" & Hex(hr) & ", " & GetAPIErrStr(hr), vbExclamation
  End If
End Function

Public Sub HideStartButton()
    '-------------------------------------------------------------
    ' Hides the Windows Start button.
    '-------------------------------------------------------------
    '   Depends on:
    '       GetTaskbarHwnd()
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const SWP_HIDEWINDOW = &H80
    '   Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    '   Private Declare Function FindWindowExA Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    '-------------------------------------------------------------
    Dim hwndTB As Long
    Dim hwndButton As Long
    
    'Get taskbar hwnd
    hwndTB = GetTaskbarHwnd()
    'Get start button hwnd
    hwndButton = FindWindowExA(hwndTB, 0&, "Button", vbNullString)
    
    Call SetWindowPos(hwndButton, 0&, 0&, 0&, 0&, 0&, SWP_HIDEWINDOW)
End Sub
Public Sub ShowStartButton()
    '-------------------------------------------------------------
    ' Shows the Windows Start button after hiding it.
    '-------------------------------------------------------------
    '   Depends on:
    '       GetTaskbarHwnd()
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const SWP_SHOWINDOW = &H40
    '   Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    '   Private Declare Function FindWindowExA Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    '-------------------------------------------------------------
    Dim hwndTB As Long
    Dim hwndButton As Long
    
    'Get taskbar hwnd
    hwndTB = GetTaskbarHwnd()
    
    'Get start button hwnd
    hwndButton = FindWindowExA(hwndTB, 0&, "Button", vbNullString)
    
    Call SetWindowPos(hwndButton, 0&, 0&, 0&, 0&, 0&, SWP_SHOWWINDOW)
End Sub
Public Sub HideTaskbar()
    '-------------------------------------------------------------
    ' Hides the Windows taskbar
    '-------------------------------------------------------------
    '   Depends on:
    '       GetTaskbarHwnd()
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const SWP_HIDEWINDOW = &H80
    '   Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    '-------------------------------------------------------------

    Call SetWindowPos(GetTaskbarHwnd(), 0&, 0&, 0&, 0&, 0&, _
        SWP_HIDEWINDOW)  'hide the Tasbar
End Sub
Public Sub ShowTaskbar()
    '-------------------------------------------------------------
    ' Shows the Windows taskbar after hiding it with HideTaskbar()
    '-------------------------------------------------------------
    '   Depends on:
    '       GetTaskbarHwnd()
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const SWP_SHOWWINDOW = &H40
    '   Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    '-------------------------------------------------------------

    Call SetWindowPos(GetTaskbarHwnd(), 0&, 0&, 0&, 0&, 0&, _
        SWP_SHOWWINDOW) 'show the Taskbar
End Sub
Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
      SetTopMostWindow = False
   End If
End Function
Public Function GetTaskbarHwnd() As Long
    '-------------------------------------------------------------
    ' Returns the hwnd of the Windows taskbar.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function FindWindowA Lib Lib "user32" _
    '       (ByVal lpClassName As String, _
    '       ByVal lpWindowName As String) As Long
    
    '-------------------------------------------------------------
    GetTaskbarHwnd = FindWindowA("Shell_traywnd", "")
End Function
Public Function GetDesktopHwnd() As Long
    '-------------------------------------------------------------
    ' Returns the hwnd of the Windows desktop.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function GetDesktopWindow Lib "user32" () As Long
    '-------------------------------------------------------------
    GetDesktopHwnd = GetDesktopWindow()
End Function
Public Function GetComputerName() As String
    '-------------------------------------------------------------
    ' Returns the name of the computer.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const MAX_PATH = 260
    '   Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
    '-------------------------------------------------------------
   Dim sBuffer As String
   Dim lRet    As Long
   Dim nSize   As Long
    
   nSize = MAX_PATH
   sBuffer = Space$(MAX_PATH)
   
   
   
   lRet = GetComputerNameA(sBuffer, nSize)
   If lRet <> 0 Then
      GetComputerName = UCase$(Left$(Trim$(sBuffer), Len(Trim$(sBuffer)) - 1))
      GetComputerName = ClsDsr.EliminarString(GetComputerName, Chr(0))
   Else
      GetComputerName = ""
   End If
End Function
Public Function GetUserName() As String
    '-------------------------------------------------------------
    ' Returns the name of the computer.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const MAX_PATH = 260
    '   Private Declare Function GetUserNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
    '-------------------------------------------------------------
    Dim r As Long
    Dim sBuffer As String
    Dim nSize As Long
    
    nSize = MAX_PATH
    
    sBuffer = Space$(MAX_PATH)
    r = GetUserNameA(sBuffer, nSize)
    
    If r Then
        GetUserName = Left$(sBuffer, nSize - 1&)
    End If
End Function
Public Function GetWindowParent(ByVal hWnd As Long) As Long
    '-------------------------------------------------------------
    ' Returns the hwnd of a window's parent window.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function GetParent Lib "user32" _
    '       (ByVal hwnd As Long) As Long
    '-------------------------------------------------------------

    GetWindowParent = GetParent(hWnd)
End Function
Public Function IsWindowChildOf(ByVal hWndChild As Long, _
    ByVal hWndParent As Long) As Boolean
    '-------------------------------------------------------------
    ' Checks to see if a window is a child of another window.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function IsChildAPI Lib "user32" _
    '       Alias "IsChild" _
    '       (ByVal hwndParent As Long, ByVal hwnd As Long) As Long
    '-------------------------------------------------------------
    
    IsWindowChildOf = Not (IsChildAPI(hWndParent, hWndChild) = 0&)
End Function
Public Function IsWindowEnabled(ByVal hWnd As Long) As Boolean
    '-------------------------------------------------------------
    ' Checks to see if a window is enabled.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function IsWindowVisibleEnabledAPI _
    '       Lib "user32" _
    '       Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
    '-------------------------------------------------------------
    
    IsWindowEnabled = Not (IsWindowEnabledAPI(hWnd) = 0&)
End Function
Public Function IsWindowVisible(ByVal hWnd As Long) As Boolean
    '-------------------------------------------------------------
    ' Checks to see if a window is a visible.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function IsWindowVisibleAPI Lib "user32" _
    '       Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
    '-------------------------------------------------------------
    
    IsWindowVisible = Not (IsWindowVisibleAPI(hWnd) = 0&)
End Function
Public Function IsWindowMinimized(ByVal hWnd As Long) As Boolean
    '-------------------------------------------------------------
    ' Checks to see if a window is a minimized.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function IsIconic Lib "user32" _
    '       (ByVal hwnd As Long) As Long
    '-------------------------------------------------------------
    
    IsWindowMinimized = Not (IsIconic(hWnd) = 0&)
End Function
Public Function IsWindowMaximized(ByVal hWnd As Long) As Boolean
    '-------------------------------------------------------------
    ' Checks to see if a window is maximized.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function IsZoomed Lib "user32" _
    '       (ByVal hwnd As Long) As Long
    '-------------------------------------------------------------
    
    IsWindowMaximized = Not (IsZoomed(hWnd) = 0&)
End Function
Public Function IsWindow(ByVal hWnd As Long) As Boolean
    '-------------------------------------------------------------
    ' Checks to see if a window is a valid window.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function IsWindowAPI Lib "user32" _
    '       Alias "IsWindow" (ByVal hwnd As Long) As Long
    '-------------------------------------------------------------
    
    IsWindow = Not (IsWindowAPI(hWnd) = 0&)
End Function
Public Function GetWindowCaption(ByVal hWnd As Long) As String
    '-------------------------------------------------------------
    ' Gets the "Caption" for the window specified by hwnd, even
    ' windows you don't own.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const MAX_PATH = 260
    '
    '   Private Declare Function GetWindowText Lib "user32" _
    '       Alias "GetWindowTextA" (ByVal hwnd As Long, _
    '       ByVal lpString As String, ByVal cch As Long) As Long
    '-------------------------------------------------------------
    
    Dim sBuffer As String
    Dim r As Long
    
    'Setup a buffer for windows to copy the name into
    sBuffer = Space$(MAX_PATH) & Chr$(0)
    
    r = GetWindowText(hWnd, sBuffer, MAX_PATH)
    
    'If this function is successful, r will contain the number of
    'characters copied to our buffer.
    
    If r Then
        'Success
        GetWindowCaption = Left$(sBuffer, r - 1&)
    End If
End Function
Public Sub SetWindowCaption(ByVal hWnd As Long, _
    ByVal sText As String)
    '-------------------------------------------------------------
    ' Sets the "Caption" for the window specified by hwnd, even
    ' windows you don't own.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Declare Function SetWindowText Lib "user32" _
    '       Alias "SetWindowTextA" (ByVal hwnd As Long, _
    '       ByVal lpString As String) As Long
    '-------------------------------------------------------------

    SetWindowText hWnd, sText
End Sub
Public Function GetWindowClass(ByVal hWnd As Long) As String
    '-------------------------------------------------------------
    ' Gets the class name for a given window
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const MAX_PATH = 260
    '
    '   Private Declare Function GetClassName Lib "user32" _
    '       Alias "GetClassNameA" (ByVal hwnd As Long, _
    '       ByVal lpClassName As String, ByVal nMaxCount As Long) _
    '       As Long
    '-------------------------------------------------------------

    Dim sBuffer As String
    Dim r As Long
    
    'Setup a buffer for windows to copy the name into
    sBuffer = Space$(MAX_PATH) & Chr$(0)
    
    r = GetClassName(hWnd, sBuffer, MAX_PATH)
    
    'If this function is successful, r will contain the number
    'of characters copied to our buffer.
    
    If r Then
        'Success
        GetWindowClass = Left$(sBuffer, r - 1&)
    End If
End Function
Public Sub MakeFormTransparent(ByVal hWnd As Long)
    '-------------------------------------------------------------
    ' Makes the background of a form transparent.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const SWP_NOSIZE = &H1
    '   Private Const SWP_NOACTIVATE = &H10
    '   Private Const SWP_SHOWWINDOW = &H40
    '   Private Const SWP_NOMOVE = &H2
    '
    '   Private Const GWL_EXStyle = (-20&)
    '   Private Const WS_EX_TRANSPARENT = &H20&
    '
    '   Private Declare Function SetWindowPos Lib "user32" _
    '       (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    '       ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    '       ByVal cy As Long, ByVal wFlags As Long) As Long
    '   Private Declare Function SetWindowLong Lib "user32" _
    '       Alias "SetWindowLongA" (ByVal hwnd As Long, _
    '       ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    '   Private Declare Function GetWindowLong Lib "user32" _
    '       Alias "GetWindowLongA" (ByVal hwnd As Long, _
    '       ByVal nIndex As Long) As Long
    '-------------------------------------------------------------

    Dim fFlags As Long
    Dim dwExStyle As Long
    
    fFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or _
        SWP_NOACTIVATE

    'Get the current extended Style bits for the window.
    dwExStyle = GetWindowLong(hWnd, GWL_EXStyle)
    
    'Combine with WS_EX_TRANSPARENT Style bits.
    dwExStyle = dwExStyle Or WS_EX_TRANSPARENT
    
    'Reset the extended Style bits.
    SetWindowLong hWnd, GWL_EXStyle, dwExStyle
    
    'Show window w/ new extended Style bits.
    SetWindowPos hWnd, 0&, 0&, 0&, 0&, 0&, fFlags
End Sub
Public Sub CenterForm(ByVal hwndForm As Long, _
    Optional hWndParent As Long, _
    Optional bShowForm As Boolean = True)
    '-------------------------------------------------------------
    ' Centers a form against the desktop window, accounting for
    ' the taskbar, or optionally against another window.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const SWP_NOSIZE = &H1
    '   Private Const SWP_NOACTIVATE = &H10
    '   Private Const SWP_SHOWWINDOW = &H40
    '
    '   Private Const SPI_GETWORKAREA = 48&
    '
    '   Private Type RECT
    '       Left As Long
    '       op As Long
    '       Right As Long
    '       Bottom As Long
    '   End Type
    '
    '   Private Declare Function SetWindowPos Lib "user32" _
    '       (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    '       ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    '       ByVal cy As Long, ByVal wFlags As Long) As Long
    '   Private Declare Function SystemParametersInfo Lib "user32" _
    '       Alias "SystemParametersInfoA" _
    '       (ByVal uAction As Long, ByVal uParam As Long, _
    '       ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    '   Private Declare Function GetWindowRect Lib "user32" _
    '       (ByVal hwnd As Long, lpRect As RECT) As Long
    '-------------------------------------------------------------
    
    Dim r As Long
    Dim uPRect As RECT, uCRect As RECT
    Dim xC As Long, yC As Long
    Dim xWid As Long, yHgt As Long
    Dim fFlags As Long
    
    'Get parent metrics
    If hWndParent = 0 Then
        'center againt desktop
        r = SystemParametersInfo(SPI_GETWORKAREA, 0&, uPRect, 0&)
    Else
        'center agaist parent
        r = GetWindowRect(hWndParent, uPRect)
    End If
    'Get child metrics
    r = GetWindowRect(hwndForm, uCRect)
    
    xC = uPRect.Left + (uPRect.Right - uPRect.Left) \ 2
    yC = uPRect.Top + (uPRect.Bottom - uPRect.Top) \ 2
    
    xWid = uCRect.Right - uCRect.Left
    yHgt = uCRect.Bottom - uCRect.Top
    
    xC = xC - xWid \ 2
    yC = yC - yHgt \ 2
        
    fFlags = SWP_NOSIZE Or SWP_NOACTIVATE
    
    If bShowForm Then
        fFlags = fFlags Or SWP_SHOWWINDOW
    End If
    
    r = SetWindowPos(hwndForm, 0&, xC, yC, 0&, 0&, fFlags)
    
End Sub
Public Sub MakeFormTopmost(ByVal hWnd As Long, Optional OnTop As Boolean = True)
    '-------------------------------------------------------------
    ' Makes the window specifed in hwnd a "topmost" window, or
    ' makes a "topmost" window not "topmost".
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Const HWND_TOPMOST = -1&
    '   Private Const HWND_NOTOPMOST = -2&
    '   Private Const SWP_NOMOVE = &H2
    '   Private Const SWP_NOSIZE = &H1
    '   Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    '
    '   Private Declare Function SetWindowPos Lib "user32" _
    '       (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    '       ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    '       ByVal cy As Long, ByVal wFlags As Long) As Long
    '-------------------------------------------------------------
    
    Dim r As Long
    
    If OnTop Then
        r = SetWindowPos(hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, TOPMOST_FLAGS)
    Else
        r = SetWindowPos(hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, TOPMOST_FLAGS)
    End If
End Sub
Public Sub SnapCursorToWindow(ByVal hWnd As Long)
    '-------------------------------------------------------------
    ' Snaps the mouse pointer to the center of the specified
    ' window. Useful for centering the cursor over the default
    ' button on a dialog, etc.
    '-------------------------------------------------------------
    '   API declarations:
    '-------------------------------------------------------------
    '   Private Type RECT
    '       Left As Long
    '       op As Long
    '       Right As Long
    '       Bottom As Long
    '   End Type
    '
    '   Private Type PointAPI
    '       x As Long
    '       y As Long
    '   End Type
    '
    '   Private Declare Function GetClientRect Lib "user32" _
    '       (ByVal hwnd As Long, lpRect As RECT) As Long
    '   Private Declare Function ClientToScreen Lib "user32" _
    '       (ByVal hwnd As Long, lpPoint As PointAPI) As Long
    '   Private Declare Function SetCursorPos Lib "user32" _
    '       (ByVal x As Long, ByVal y As Long) As Long
    '-------------------------------------------------------------

    Dim uRECT As RECT
    Dim r As Long
    Dim uPt As PointAPI
    Dim X As Long
    Dim Y As Long
    
    r = GetClientRect(hWnd, uRECT)
    uPt.X = (uRECT.Right - uRECT.Left) \ 2
    uPt.Y = (uRECT.Bottom - uRECT.Top) \ 2
    
    r = ClientToScreen(hWnd, uPt)
    r = SetCursorPos(uPt.X, uPt.Y)
    
End Sub




