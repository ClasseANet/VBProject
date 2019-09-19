VERSION 5.00
Begin VB.Form frmODBCLog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Logon"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   4500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3165
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   380
      Left            =   1560
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   2655
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   380
      Left            =   3000
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   2655
      Width           =   1200
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   380
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   2655
      Width           =   1200
   End
   Begin VB.TextBox TxtUID 
      Height          =   300
      Left            =   1110
      TabIndex        =   5
      Top             =   720
      Width           =   3100
   End
   Begin VB.TextBox TxtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1110
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1050
      Width           =   3100
   End
   Begin VB.TextBox TxtDatabase 
      Height          =   300
      Left            =   1110
      TabIndex        =   3
      Top             =   1380
      Width           =   3100
   End
   Begin VB.ComboBox CmbDSN 
      Height          =   315
      ItemData        =   "ODBCLog2.frx":0000
      Left            =   1110
      List            =   "ODBCLog2.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "È"
      Top             =   360
      Width           =   3100
   End
   Begin VB.TextBox TxtServer 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1110
      TabIndex        =   1
      Top             =   2055
      Width           =   3100
   End
   Begin VB.ComboBox CmbDrivers 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "ODBCLog2.frx":0004
      Left            =   1110
      List            =   "ODBCLog2.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1710
      Width           =   3100
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   0
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   2295
      Index           =   1
      Left            =   75
      Top             =   255
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Index           =   0
      Left            =   60
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&DSN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   14
      Top             =   405
      Width           =   465
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&UID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   750
      Width           =   405
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   12
      Top             =   1095
      Width           =   885
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data&base:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   1440
      Width           =   885
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dri&ver:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   10
      Top             =   1785
      Width           =   585
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Server:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   9
      Top             =   2130
      Width           =   630
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmODBCLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "ODBC Logon"
Const BUTTON1 = "&OK"
Const BUTTON2 = "&Cancel"
Const BUTTON3 = "&Register"
Const FRAME1 = "Connect Values:"
Const Label1 = "&DSN:"
Const Label2 = "&UID:"
Const LABEL3 = "&Password:"
Const LABEL4 = "Data&base:"
Const LABEL5 = "Dri&ver:"
Const LABEL6 = "&Server:"
Const MSG1 = "Enter ODBC Connection Parameters"
Const MSG2 = "Opening ODBC Database"
Const MSG3 = "Enter Driver Name from ODBCINST.INI File:"
Const MSG4 = "Driver Name"
Const MSG5 = "This Datasource has not been Registered, this will now be attempted for you!"
Const MSG7 = "Invalid Parameter(s), Please try again!"
Const MSG8 = "Query Timeout Could not be set, default will be used!"
Const MSG9 = "Datasource Registration Succeeded, proceed with Open."
Const MSG10 = "Please enter a DSN!"
Const MSG11 = "Please select a Driver!"
Const MSG12 = "You must Close First!"
'>>>>>>>>>>>>>>>>>>>>>>>>

Dim mbBeenLoaded As Integer
Public DBOpened As Boolean

Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1
Private Sub CmbDrivers_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub CmbDSN_Change()
  If Len(CmbDSN.Text) = 0 Or CmbDSN.Text = "(None)" Then
    TxtServer.Enabled = True
    CmbDrivers.Enabled = True
    lblLabels(4).Enabled = True
    lblLabels(5).Enabled = True
  Else
    TxtServer.Enabled = False
    CmbDrivers.Enabled = False
    lblLabels(4).Enabled = False
    lblLabels(5).Enabled = False
  End If
End Sub
Private Sub CmbDSN_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub cmdCancel_Click()
'  gbDBOpenFlag = False
'  gsDBName = vbNullString
'  DBOpened = False
'  Me.Hide
   Unload Me
End Sub
Private Sub cmdOK_Click()
  On Error GoTo cmdOK_ClickErr

  Dim sConnect As String
  Dim dbTemp As Database

'  MsgBar MSG2, True

'  If MdiPrincipal.mnuPOpenOnStartup.Checked Then
'    Me.Refresh
'  End If
  
  Screen.MousePointer = vbHourglass
  
  If Len(CmbDSN.Text) > 0 Then
    sConnect = "ODBC;DSN=" & CmbDSN.Text & ";"
  Else
    sConnect = "ODBC;Driver={" & CmbDrivers.Text & "};"
    sConnect = sConnect & "Server=" & TxtServer.Text & ";"
  End If
  
  sConnect = sConnect & "UID=" & TxtUID.Text & ";"
  sConnect = sConnect & "PWD=" & TxtPWD.Text & ";"
  If Len(TxtDatabase.Text) > 0 Then
    sConnect = sConnect & "Database=" & TxtDatabase.Text & ";"
  End If
  
'  Set dbTemp = gwsMainWS.OpenDatabase("", 0, 0, sConnect)
  DB.dbODBC = "S"
'''''''  Call Db.SrvConecta("", "", CmbDSN.Text, TxtUID.Text, TxtPWD.Text, "")
  Call DB.SrvConecta("", "", CmbDSN.Text, TxtUID.Text, TxtPWD.Text, Me.TxtDatabase)
  
'  If gbDBOpenFlag Then
'    If gbDBOpenFlag Then
'      Beep
'      MsgBox MSG12, 48
'      Me.Hide
'      Exit Sub
'    End If
'  End If

  'success
  DBOpened = True
  'save the values
'  gsODBCDatasource = CmbDSN.Text
'  gsDBName = gsODBCDatasource
'  gsODBCDatabase = TxtDatabase.Text
'  gsODBCUserName = TxtUID.Text
'  gsODBCPassword = TxtPWD.Text
'  gsODBCDriver = CmbDrivers.Text
'  gsODBCServer = TxtServer.Text
'  gsDataType = gsSQLDB

'  Set gdbCurrentDB = dbTemp
'  GetODBCConnectParts gdbCurrentDB.Connect

'  CmbDSN.Text = gsODBCDatasource
'  TxtDatabase.Text = gsODBCDatabase
'  TxtUID.Text = gsODBCUserName
'  TxtPWD.Text = gsODBCPassword

'  frmMDI.Caption = "VisData:" & gsDBName & "." & gsODBCDatabase
'  gdbCurrentDB.QueryTimeout = glQueryTimeout

'  gbDBOpenFlag = True
'  AddMRU

  Screen.MousePointer = vbDefault
  Unload Me
  Exit Sub

cmdOK_ClickErr:
  Screen.MousePointer = vbDefault
'  gbDBOpenFlag = False
  If Len(CmbDSN.Text) > 0 Then
    If InStr(1, Error, "ODBC--connection to '" & CmbDSN.Text & "' failed") > 0 Then
      Beep
      MsgBox MSG5, 48
      TxtDatabase.Text = vbNullString
      TxtUID.Text = vbNullString
      TxtPWD.Text = vbNullString
      If RegisterDB((CmbDSN.Text)) Then
        MsgBox MSG9, 48
      End If
    ElseIf InStr(1, Error, "Login failed") > 0 Then
      Beep
      MsgBox MSG7, 48
    ElseIf InStr(1, Error, "QueryTimeout property") > 0 Then
'      If glQueryTimeout <> 5 Then
'        Beep
'        MsgBox MSG8, 48
'      End If
      Resume Next
    Else
      ShowError
    End If
  End If
  
'  MsgBar MSG1, False
  If Err = 3059 Then
    Unload Me
  End If

End Sub
Private Sub cmdRegister_Click()
  On Error GoTo Fim
  If Len(CmbDSN.Text) = 0 Then
    MsgBox MSG10, vbInformation, Me.Caption
    Exit Sub
  End If
  If Len(CmbDrivers.Text) = 0 Then
    MsgBox MSG11, vbInformation, Me.Caption
    Exit Sub
  End If
  'try to register it
  DBEngine.RegisterDatabase CmbDSN.Text, CmbDrivers.Text, False, vbNullString

  MsgBox MSG9, vbInformation
  
  Exit Sub
Fim:
  ShowError
End Sub

Private Sub Form_Activate()
   Screen.MousePointer = vbDefault
   Set Sys.MDIFilho = Me
   Select Case ""
      Case Trim(CmbDSN): CmbDSN.SetFocus
      Case Trim(TxtUID): TxtUID.SetFocus
      Case Trim(TxtPWD): TxtPWD.SetFocus
      Case Else
         If TxtDatabase.Enabled Then
            TxtDatabase.SetFocus
         Else
            cmdOK.SetFocus
         End If
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: Unload Me
      Case Else: KeyAscii = SendTab(Me, KeyAscii)
   End Select
End Sub
Private Sub Form_Load()
  Dim i As Integer

  Me.Caption = FORMCAPTION
  cmdOK.Caption = BUTTON1
  cmdCancel.Caption = BUTTON2
  cmdRegister.Caption = BUTTON3
  LblTitle.Caption = FRAME1
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
  lblLabels(2).Caption = LABEL3
  lblLabels(3).Caption = LABEL4
  lblLabels(4).Caption = LABEL5
  lblLabels(5).Caption = LABEL6
  GetDSNsAndDrivers

'  MsgBar MSG1, False
  CmbDSN.Text = "UNB01_32"
'  TxtDatabase.Text = "RIO_TST"
'  TxtUID.Text = "ORDSR"
'  TxtPWD.Text = "P678694694"
  
  TxtDatabase.Text = "" ' "RIO07"
  TxtUID.Text = "TECA"
  TxtPWD.Text = "TECAPLUS"
  
  TxtDatabase.Enabled = False
  Me.lblLabels(3).Enabled = False
  
'  CmbDSN.Text = gsODBCDatasource
'  TxtDatabase.Text = gsODBCDatabase
'  TxtUID.Text = gsODBCUserName
'  TxtPWD.Text = gsODBCPassword
'  If Len(gsODBCDriver) > 0 Then
'    For i = 0 To CmbDrivers.ListCount - 1
'      If CmbDrivers.List(i) = gsODBCDriver Then
'        CmbDrivers.ListIndex = i
'        Exit For
'      End If
'    Next
'  End If
'  TxtServer.Text = gsODBCServer
  Call ConfigForm(Me, SysMdi.Icon, Sys.FundoTela)
  mbBeenLoaded = True
End Sub
Private Sub CmbDSN_Click()
  CmbDSN_Change
End Sub
Sub GetDSNsAndDrivers()
  On Error Resume Next
  
  Dim i As Integer
  Dim sDSNItem As String * 1024
  Dim sDRVItem As String * 1024
  Dim sDSN As String
  Dim sDRV As String
  Dim iDSNLen As Integer
  Dim iDRVLen As Integer
  Dim lHenv As Long     'handle to the environment

  CmbDSN.AddItem "(None)"

  'get the DSNs
  If SQLAllocEnv(lHenv) <> -1 Then
    Do Until i <> SQL_SUCCESS
      sDSNItem = Space(1024)
      sDRVItem = Space(1024)
      i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
      sDSN = VBA.Left(sDSNItem, iDSNLen)
      sDRV = VBA.Left(sDRVItem, iDRVLen)
        
      If sDSN <> Space(iDSNLen) Then
        CmbDSN.AddItem sDSN
        CmbDrivers.AddItem sDRV
      End If
    Loop
  End If
  'remove the dupes
  If CmbDSN.ListCount > 0 Then
    With CmbDrivers
      If .ListCount > 1 Then
        i = 0
        While i < .ListCount
          If .List(i) = .List(i + 1) Then
            .RemoveItem (i)
          Else
            i = i + 1
          End If
        Wend
      End If
    End With
  End If
  CmbDSN.ListIndex = 0
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set Sys.MDIFilho = Nothing
'  MsgBar vbNullString, False
End Sub
Private Function RegisterDB(rsDatasource As String) As Integer
   On Error GoTo RDBErr

   Dim sDriver As String

'   sDriver = InputBox(MSG3, MSG4, gsDEFAULT_DRIVER)
'   If sDriver <> gsDEFAULT_DRIVER Then
'     DBEngine.RegisterDatabase rsDatasource, sDriver, False, vbNullString
'   Else
'     DBEngine.RegisterDatabase rsDatasource, sDriver, True, vbNullString
'   End If

   RegisterDB = True
   Exit Function

RDBErr:
   RegisterDB = False
End Function

Private Sub Image1_Click()

End Sub

Private Sub TxtDatabase_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub TxtPWD_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub TxtServer_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub
Private Sub TxtUID_GotFocus()
   Call SelecionarTexto(ActiveControl)
End Sub

