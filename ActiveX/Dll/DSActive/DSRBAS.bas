Attribute VB_Name = "DSRBAS"
Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Public Declare Function ImageList_BeginDrag Lib "Comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As cBoolean
Public Declare Function ImageList_DragEnter Lib "Comctl32.dll" (ByVal hwndLock As Long, ByVal x As Long, ByVal y As Long) As cBoolean
Public Declare Function ImageList_DragLeave Lib "Comctl32.dll" (ByVal hwndLock As Long) As cBoolean
Public Declare Function ImageList_DragMove Lib "Comctl32.dll" (ByVal x As Long, ByVal y As Long) As cBoolean
Public Declare Function ImageList_DragShowNolock Lib "Comctl32.dll" (ByVal fShow As Boolean) As cBoolean
Public Declare Function ImageList_Destroy Lib "Comctl32.dll" (ByVal himl As Long) As cBoolean
Public Declare Function ImageList_GetImageCount Lib "Comctl32.dll" (ByVal himl As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "Comctl32.dll" (ByVal himl As Long, lpcx As Long, lpcy As Long) As Boolean
Public Declare Function InitCommonControlsEx Lib "Comctl32.dll" (lpInitCtrls As TagInitCommonControlsEx) As Boolean
Public Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As SB_Type, lpsi As SCROLLINFO) As Boolean


'********************
'********************
'** Classes Comuns **
'********************
'********************
Public ClsDsr     As New DSR
Public ClsDos     As New DOS
Public ClsMsg     As New Mensagem
Public ClsCtrl    As New DSControl
Public ClsOffice  As New Office
Public ClsAutoIns As New AutoInstall
Public ClsSeg     As New DS_SEGURANCA
' Returns a set of bit flags indicating whether the specified point resides in
' the specified size region with the perimeter of the specified rect. cxyRegion
' defines the rectangular region within rc, and must be a positive value
Public Function PtInRectRegion(rc As RECT, cxyRegion As Long, pt As PointAPI) As RectFlags
'Public Function PtInRectRegion(Rc As Variant, cxyRegion As Long, Pt As Variant) As Variant
  Dim dwFlags As RectFlags
  
  If PtInRect(rc, pt.x, pt.y) Then
    dwFlags = (rfLeft And (pt.x <= (rc.Left + cxyRegion)))
    dwFlags = dwFlags Or (rfRight And (pt.x >= (rc.Right - cxyRegion)))
    dwFlags = dwFlags Or (rfTop And (pt.y <= (rc.Top + cxyRegion)))
    dwFlags = dwFlags Or (rfBottom And (pt.y >= (rc.Bottom - cxyRegion)))
  End If
  
  PtInRectRegion = dwFlags

End Function

' Converts the specified window's client coords to window
' coords (relative to the window's rect origin)

Public Function ClientToWindow(hWnd As Long, pt As PointAPI) As Boolean
'Public Function ClientToWindow(hwnd As Long, Pt As Variant) As Boolean
  Dim fRtn As Boolean
  Dim rcClient As RECT
  Dim rcWindow As RECT
  
  If IsWindow(hWnd) Then
    fRtn = CBool(GetClientRect(hWnd, rcClient))
    fRtn = fRtn And CBool(ClientToScreen(hWnd, rcClient))
    fRtn = fRtn And CBool(GetWindowRect(hWnd, rcWindow))
    If fRtn Then
      pt.x = pt.x + (rcClient.Left - rcWindow.Left)
      pt.y = pt.y + (rcClient.Top - rcWindow.Top)
      ClientToWindow = True
    End If
  End If
  
End Function

' Retrieves the bounding rectangle for a tree-view Item and indicates whether the Item is visible.
' If the Item is visible and retrieves the bounding rectangle, the return value is TRUE.
' Otherwise, the TVM_GETItemRECT message returns FALSE and does not retrieve
' the bounding rectangle.
' If fItemRect = TRUE, returns label rect. Otherwise, entire Item line rect is  returned.

Public Function TreeView_GetItemRect(hWnd As Long, hItem As Long, prc As RECT, fItemRect As cBoolean) As Boolean
'Public Function TreeView_GetItemRect(hwnd As Long, hItem As Long, prc As Variant, fItemRect As Variant) As Boolean
  prc.Left = hItem
  TreeView_GetItemRect = SendMessage(hWnd, TVM_GETItemRECT, ByVal fItemRect, prc)
End Function

