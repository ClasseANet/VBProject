VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Set_AddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Set_Constru"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Set_Edit"
Attribute VB_Ext_KEY = "Member2" ,"Set_Proj"
Private mvarProj As Set_Proj
Private mvarConstru As Set_Constru
Private mvarEdit As Set_Edit
Public Property Get Edit() As Set_Edit
   If mvarEdit Is Nothing Then
       Set mvarEdit = New Set_Edit
   End If
   Set Edit = mvarEdit
End Property
Public Property Set Edit(vData As Set_Edit)
   Set mvarEdit = vData
End Property
Public Property Get Constru() As Set_Constru
   If mvarConstru Is Nothing Then
       Set mvarConstru = New Set_Constru
   End If
   Set Constru = mvarConstru
End Property
Public Property Set Constru(vData As Set_Constru)
   Set mvarConstru = vData
End Property
Public Property Get Proj() As Set_Proj
   If mvarProj Is Nothing Then
       Set mvarProj = New Set_Proj
   End If
   Set Proj = mvarProj
End Property
Public Property Set Proj(vData As Set_Proj)
   Set mvarSet_Proj = vData
End Property
Private Sub Class_Terminate()
  Set mvarEdit = Nothing
  Set mvarConstru = Nothing
  Set mvarProj = Nothing
End Sub




