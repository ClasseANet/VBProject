VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Identidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mvarIDLOJA As Integer
Public mvarNMLOJA As String
Public mvarIDFUNC As Integer
Public mvarNMFUNC As String
Public mvarDTHORA As String

Private WithEvents mvarMe  As FrmIdentidade
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys      As Object
Private mvarClFinger As Object
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Set ClFinger(ByVal vData As Object)
   Set mvarClFinger = vData
End Property
Public Property Get ClFinger() As Object
    Set ClFinger = mvarClFinger
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
    IDLOJA = mvarIDLOJA
End Property
Public Property Let NMLOJA(ByVal vData As String)
   mvarNMLOJA = vData
End Property
Public Property Get NMLOJA() As String
    NMLOJA = mvarNMLOJA
End Property
Public Property Let IDFUNC(ByVal vData As Integer)
   mvarIDFUNC = vData
End Property
Public Property Get IDFUNC() As Integer
    IDFUNC = mvarIDFUNC
End Property
Public Property Let NMFUNC(ByVal vData As String)
   mvarNMFUNC = vData
End Property
Public Property Get NMFUNC() As String
    NMFUNC = mvarNMFUNC
End Property
Public Property Let DTHORA(ByVal vData As String)
   mvarDTHORA = vData
End Property
Public Property Get DTHORA() As String
    DTHORA = mvarDTHORA
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmIdentidade
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_CmdOkClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
      Unload mvarMe
   End If
End Sub
Private Sub mvarMe_Load()
   With mvarMe
      If CDate(Format(mvarDTHORA, "hh:mm")) < CDate("12:00") Then
         .LblSaudacao.Caption = "Bom Dia!"
      ElseIf CDate(Format(mvarDTHORA, "hh:mm")) < CDate("18:00") Then
         .LblSaudacao.Caption = "Boa Tarde!"
      Else
         .LblSaudacao.Caption = "Boa Noite!"
      End If
      
      .LblNome.Caption = mvarNMFUNC
      
      .LblHora.ForeColor = vbBlack
      .LblHora.Caption = Format(mvarDTHORA, "hh:mm") & "h"
      
      Select Case Weekday(DTHORA)
         Case 1: .LblDia.Caption = "Domingo"
         Case 2: .LblDia.Caption = "Segunda-Feria"
         Case 3: .LblDia.Caption = "Terça-Feria"
         Case 4: .LblDia.Caption = "Quarta-Feria"
         Case 5: .LblDia.Caption = "Quinta-Feria"
         Case 6: .LblDia.Caption = "Sexta-Feria"
         Case 7: .LblDia.Caption = "Sábado"
      End Select
      .LblDia.Caption = .LblDia.Caption & ", " & Format(mvarDTHORA, "dd/mm/yyyy")
      
      .Image1.Width = 1600
      .Image1.Height = 1800
      .Image1.Stretch = True
      If ExisteArquivo(mvarClFinger.fdbPath & Trim(mvarIDFUNC) & ".jpg") Then
         .Image1.Picture = LoadPicture(mvarClFinger.fdbPath & Trim(mvarIDFUNC) & ".jpg")
      End If
      .Caption = " Registro Biométrico (" & mvarNMLOJA & ")"
   End With
End Sub
