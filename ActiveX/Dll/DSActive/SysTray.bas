Attribute VB_Name = "SysTray"
Option Explicit

' Type passed to Shell_NotifyIcon
Public Type NotifyIconData
  Size As Long
  Handle As Long
  ID As Long
  Flags As Long
  CallBackMessage As Long
  Icon As Long
  Tip As String * 64
End Type

' Constants for managing System Tray tasks, foudn in shellapi.h
Public Const AddIcon = &H0
Public Const ModifyIcon = &H1
Public Const DeleteIcon = &H2

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Const MessageFlag = &H1
Public Const IconFlag = &H2
Public Const TipFlag = &H4
      
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, Data As NotifyIconData) As Boolean
Public Sub AddIconToTray(ByRef ObjTray As Object, ByRef pData As NotifyIconData, Optional pToolTip = "")
   
   
   With pData
      .Size = Len(pData)
      .Handle = ObjTray.hwnd
      .ID = vbNull
      .Flags = IconFlag Or TipFlag Or MessageFlag
      .CallBackMessage = WM_MOUSEMOVE
      If .Icon = 0 Then
         If UCase(TypeName(ObjTray)) = "Forms" Then
            .Icon = ObjTray.Icon
         Else
            If ObjTray.Picture = 0 Then
               .Icon = ObjTray.Parent.Icon
            Else
               .Icon = ObjTray.Picture
            End If
         End If
      End If
      .Tip = pToolTip & vbNullChar
   End With
  Call Shell_NotifyIcon(AddIcon, pData)
End Sub
Public Sub DeleteIconFromTray(ByRef pData As NotifyIconData)
  Call Shell_NotifyIcon(DeleteIcon, pData)
End Sub

