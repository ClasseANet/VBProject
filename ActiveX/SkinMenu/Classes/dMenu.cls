VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Name As String
Public Caption As String
Attribute Caption.VB_VarMemberFlags = "200"
Public Ident As Long
Public Visible As Boolean
Public Enabled As Boolean
Public Checked As Boolean
Public UseMaskColor As Boolean
Public MaskColor As OLE_COLOR
Public Description As String

Public X As Long
Attribute X.VB_VarMemberFlags = "40"
Public Y As Long
Attribute Y.VB_VarMemberFlags = "40"

Dim m_Picture(2) As StdPicture
Public TemporaryPicture As StdPicture
Attribute TemporaryPicture.VB_VarMemberFlags = "40"

Private RECT As RECT

Public Expanded As Boolean
Public RootIndex As Long
Public IsRoot As Boolean
Public Accelerator As String
Public Property Get APicture(WhichPicture As IFCWhatPictures) As StdPicture
Attribute APicture.VB_MemberFlags = "40"
 Set APicture = TemporaryPicture 'm_Picture(WhichPicture)
End Property '(Public) Property Get APicture () As StdPicture

Property Set APicture(WhichPicture As IFCWhatPictures, ByVal APicture As StdPicture)
 'Set m_Picture(WhichPicture) = APicture
 Set TemporaryPicture = APicture
 
End Property ' Property Set APicture

Function PointInside(X As Single, Y As Single) As Boolean
Attribute PointInside.VB_MemberFlags = "40"
  PointInside = X >= RECT.Left And X <= RECT.Right And Y >= RECT.Top And Y <= RECT.Bottom
End Function

Sub SetRect(Left As Long, Top As Long, Right As Long, Bottom As Long)
Attribute SetRect.VB_MemberFlags = "40"
    RECT.Left = Left
    RECT.Top = Top
    RECT.Right = Right
    RECT.Bottom = Bottom
End Sub

Sub GetRect(Left As Long, Top As Long, Right As Long, Bottom As Long)
Attribute GetRect.VB_MemberFlags = "40"
    Left = RECT.Left
    Top = RECT.Top
    Right = RECT.Right
    Bottom = RECT.Bottom
End Sub
Private Sub Class_Initialize()
 Visible = True
 Enabled = True
 MaskColor = QBColor(13)
 UseMaskColor = True
End Sub


