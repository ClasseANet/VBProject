VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SenhaEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function InputBoxPassword(prompt, Optional Title, Optional Default) As String
   Const NV_INPUTBOX = &H5000&
   Dim Ret As String
   
   SetTimer 0, NV_INPUTBOX, 10, AddressOf TimerProc
   Ret = InputBox(prompt) ', Title, Default)
   InputBoxPassword = Ret
End Function

