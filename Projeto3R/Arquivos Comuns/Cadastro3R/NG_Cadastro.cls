VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum eTipoSMS
    PosicaoDia
    LembrarSessao
    LembrarCreme
End Enum
Private mvarPane        As Object

Private mvarSys         As Object
Private mvarIDCOLIGADA  As Integer
Private mvarIDLOJA      As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   
   mvarIDCOLIGADA = Val(mvarSys.Propriedades("IDCOLIGADA"))
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Sub F_SHOW(pItem As Variant)
   Dim xMDI As TL_FrmMDI
      
   Set xMDI = New TL_FrmMDI
   With xMDI
      Set .Sys = mvarSys
      
      If Not IsNumeric(pItem) Then
         .IDMODU = pItem
         If mvarSys.XDb.AbreTabela("Select ID From MODULO Where IDMODU='" & pItem & "'") Then
            pItem = xVal(mvarSys.XDb.RSAux("ID") & "")
         Else
            If Mid(pItem, 1, 3) = "TL_" Then
               If mvarSys.XDb.AbreTabela("Select ID From MODULO Where IDMODU='" & Mid(pItem, 4) & "'") Then
                  pItem = xVal(mvarSys.XDb.RSAux("ID") & "")
               End If
            End If
         End If
      End If
      .ItemCad = pItem
      .Show
   End With
   Set xMDI = Nothing
End Sub

