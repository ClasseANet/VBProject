Attribute VB_Name = "NBioAPI"
Option Explicit
' #################################################################################################
'   NBioAPI.bas   : Define constants for NBioAPI
'   Copyright     : NITGEN Co., Ltd.
' #################################################################################################
' -------------------------------------------------------------------------------------------------
'   Error code
' -------------------------------------------------------------------------------------------------
Global Const NBioAPIERROR_NONE = 0
' -------------------------------------------------------------------------------------------------
'   General
' -------------------------------------------------------------------------------------------------
' True / False
Global Const NBioAPI_TRUE = 1
Global Const NBioAPI_FALSE = 0
' -------------------------------------------------------------------------------------------------
'   Device
' -------------------------------------------------------------------------------------------------
' Constant for DeviceID
Global Const NBioAPI_DEVICE_ID_NONE = 0
Global Const NBioAPI_DEVICE_ID_FDP02_0 = 1
Global Const NBioAPI_DEVICE_ID_FDU01_0 = 2
Global Const NBioAPI_DEVICE_ID_OSU02_0 = 3
Global Const NBioAPI_DEVICE_ID_FDU11_0 = 4
Global Const NBioAPI_DEVICE_ID_FSC01_0 = 5
Global Const NBioAPI_DEVICE_ID_FDU03_0 = 6
Global Const NBioAPI_DEVICE_ID_AUTO_DETECT = 255
' Constant for Device Name
Global Const NBioAPI_DEVICE_NAME_FDP02 = 1
Global Const NBioAPI_DEVICE_NAME_FDU01 = 2
Global Const NBioAPI_DEVICE_NAME_OSU02 = 3
Global Const NBioAPI_DEVICE_NAME_FDU11 = 4
Global Const NBioAPI_DEVICE_NAME_FSC01 = 5
Global Const NBioAPI_DEVICE_NAME_FDU03 = 6
' -------------------------------------------------------------------------------------------------
'   BSP
' -------------------------------------------------------------------------------------------------
' Constant for Security Level
Global Const NBioAPI_FIR_SECURITY_LEVEL_LOWEST = 1
Global Const NBioAPI_FIR_SECURITY_LEVEL_LOWER = 2
Global Const NBioAPI_FIR_SECURITY_LEVEL_LOW = 3
Global Const NBioAPI_FIR_SECURITY_LEVEL_BELOW_NORMAL = 4
Global Const NBioAPI_FIR_SECURITY_LEVEL_NORMAL = 5
Global Const NBioAPI_FIR_SECURITY_LEVEL_ABOVE_NORMAL = 6
Global Const NBioAPI_FIR_SECURITY_LEVEL_HIGH = 7
Global Const NBioAPI_FIR_SECURITY_LEVEL_HIGHER = 8
Global Const NBioAPI_FIR_SECURITY_LEVEL_HIGHEST = 9

' Purpose for FIR
Global Const NBioAPI_FIR_PURPOSE_VERIFY = 1
Global Const NBioAPI_FIR_PURPOSE_IDENTIFY = 2
Global Const NBioAPI_FIR_PURPOSE_ENROLL = 3
Global Const NBioAPI_FIR_PURPOSE_ENROLL_FOR_VERIFICATION_ONLY = 4
Global Const NBioAPI_FIR_PURPOSE_ENROLL_FOR_IDENTIFICATION_ONLY = 5
Global Const NBioAPI_FIR_PURPOSE_AUDIT = 6
Global Const NBioAPI_FIR_PURPOSE_UPDATE = 10

' Finger ID
Global Const NBioAPI_FINGER_ID_UNKNOWN = 0
Global Const NBioAPI_FINGER_ID_RIGHT_THUMB = 1
Global Const NBioAPI_FINGER_ID_RIGHT_INDEX = 2
Global Const NBioAPI_FINGER_ID_RIGHT_MIDDLE = 3
Global Const NBioAPI_FINGER_ID_RIGHT_RING = 4
Global Const NBioAPI_FINGER_ID_RIGHT_LITTLE = 5
Global Const NBioAPI_FINGER_ID_LEFT_THUMB = 6
Global Const NBioAPI_FINGER_ID_LEFT_INDEX = 7
Global Const NBioAPI_FINGER_ID_LEFT_MIDDLE = 8
Global Const NBioAPI_FINGER_ID_LEFT_RING = 9
Global Const NBioAPI_FINGER_ID_LEFT_LITTLE = 10

' Window Style
Global Const NBioAPI_WINDOW_STYLE_POPUP = 0
Global Const NBioAPI_WINDOW_STYLE_INVISIBLE = 1     'only for NBioAPI_Capture()
Global Const NBioAPI_WINDOW_STYLE_CONTINUOUS = 2

Global Const NBioAPI_WINDOW_STYLE_NO_FPIMG = 65536
Global Const NBioAPI_WINDOW_STYLE_TOPMOST = 131072  ' currently not used (after v2.3)
Global Const NBioAPI_WINDOW_STYLE_NO_WELCOME = 262144
Global Const NBioAPI_WINDOW_STYLE_NO_TOPMOST = 524288

' -------------------------------------------------------------------------------------------------
'   Export Data
' -------------------------------------------------------------------------------------------------
Global Const MINCONV_TYPE_FDP = 0
Global Const MINCONV_TYPE_FDU = 1
Global Const MINCONV_TYPE_FDA = 2
Global Const MINCONV_TYPE_OLD_FDA = 3
Global Const MINCONV_TYPE_FDAC = 4
Global Const MINCONV_TYPE_FIM10_HV = 5
Global Const MINCONV_TYPE_FIM10_LV = 6
Global Const MINCONV_TYPE_FIM01_HV = 7
Global Const MINCONV_TYPE_FIM01_HD = 8
Global Const MINCONV_TYPE_FELICA = 9
Global Const MINCONV_TYPE_EXTENSION = 10
Global Const MINCONV_TYPE_TEMPLATESIZE_32 = 11
Global Const MINCONV_TYPE_TEMPLATESIZE_48 = 12
Global Const MINCONV_TYPE_TEMPLATESIZE_64 = 13
Global Const MINCONV_TYPE_TEMPLATESIZE_80 = 14
Global Const MINCONV_TYPE_TEMPLATESIZE_96 = 15
Global Const MINCONV_TYPE_TEMPLATESIZE_112 = 16
Global Const MINCONV_TYPE_TEMPLATESIZE_128 = 17
Global Const MINCONV_TYPE_TEMPLATESIZE_144 = 18
Global Const MINCONV_TYPE_TEMPLATESIZE_160 = 19
Global Const MINCONV_TYPE_TEMPLATESIZE_176 = 20
Global Const MINCONV_TYPE_TEMPLATESIZE_192 = 21
Global Const MINCONV_TYPE_TEMPLATESIZE_208 = 22
Global Const MINCONV_TYPE_TEMPLATESIZE_224 = 23
Global Const MINCONV_TYPE_TEMPLATESIZE_240 = 24
Global Const MINCONV_TYPE_TEMPLATESIZE_256 = 25
Global Const MINCONV_TYPE_TEMPLATESIZE_272 = 26
Global Const MINCONV_TYPE_TEMPLATESIZE_288 = 27
Global Const MINCONV_TYPE_TEMPLATESIZE_304 = 28
Global Const MINCONV_TYPE_TEMPLATESIZE_320 = 29
Global Const MINCONV_TYPE_TEMPLATESIZE_336 = 30
Global Const MINCONV_TYPE_TEMPLATESIZE_352 = 31
Global Const MINCONV_TYPE_TEMPLATESIZE_368 = 32
Global Const MINCONV_TYPE_TEMPLATESIZE_384 = 33
Global Const MINCONV_TYPE_TEMPLATESIZE_400 = 34
' -------------------------------------------------------------------------------------------------
'   Export Image
' -------------------------------------------------------------------------------------------------

' Constant for FP Image
Global Const NBioAPI_IMG_TYPE_RAW = 1
Global Const NBioAPI_IMG_TYPE_BMP = 2
Global Const NBioAPI_IMG_TYPE_JPG = 3

' Declaration global variables
Public bBiometria As Boolean

Public objNBioBSP As Object      'NBioBSPCOMLib.NBioBSP
Public objDevice As Object       'IDevice                ' Device object
Public objExtraction As Object   'IExtraction        ' Extraction object
Public objNSearch As Object      'INSearch              ' NSearch object
Public szTextEncodeFIR As String
Public Sub Init_Finger()
   If Not bBiometria Then Exit Sub
   On Error Resume Next
   
   'Set objNBioBSP = New NBioBSPCOMLib.NBioBSP
   Set objNBioBSP = CriarObjeto("NBioBSPCOM.NBioBSP")
   If objNBioBSP Is Nothing Then Exit Sub
   Set objDevice = objNBioBSP.Device
   Set objExtraction = objNBioBSP.Extraction        ' Extraction object
   Set objNSearch = objNBioBSP.NSearch
   
   ' Check initialize object
   If objNSearch.ErrorCode <> NBioAPIERROR_NONE Then
       MsgBox objNSearch.ErrorDescription & " [" & objNSearch.ErrorCode & "]", vbOKOnly
   End If
End Sub
Public Sub Terminate_Finger()
   If Not bBiometria Then Exit Sub
   On Error Resume Next
   Set objDevice = Nothing
   Set objExtraction = Nothing
   Set objNSearch = Nothing
   Set objNBioBSP = Nothing
End Sub
Public Sub Identify_Finger()
   If Not bBiometria Then Exit Sub
    
   Dim i As Integer
   Dim szTextEncodeFIR As String
   Dim ListItem As ListItem
   Dim nUserID            As Long

   On Error Resume Next
        
   If objNBioBSP Is Nothing Then
      Call Init_Finger
   End If
   nUserID = 0
   szTextEncodeFIR = ""
'    Call ListResult.ListItems.Clear
   
   ' Get FIR data
   Call objDevice.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
   Call objExtraction.Capture(NBioAPI_FIR_PURPOSE_VERIFY)
   If objExtraction.ErrorCode <> NBioAPIERROR_NONE Then
      If objExtraction.ErrorCode = 513 Then     '* Cancelado
      ElseIf objExtraction.ErrorCode = 515 Then '* Timeout
      Else
        MsgBox objExtraction.ErrorDescription & " [" & objExtraction.ErrorCode & "]"
        Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
      End If
      Exit Sub
    End If

    Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)

    szTextEncodeFIR = objExtraction.TextEncodeFIR

    ' Identify FIR to NSearch DB
    Call objNSearch.IdentifyUser(szTextEncodeFIR, NBioAPI_FIR_SECURITY_LEVEL_NORMAL)

    If objNSearch.ErrorCode <> 0 Then
      If objNSearch.ErrorCode = 777 Then
         Call ExibirStop("Digital não identificada." & vbNewLine & vbNewLine & "Tente novamente.")
      Else
        MsgBox objNSearch.ErrorDescription & " [" & objNSearch.ErrorCode & "]", vbOKOnly, "Error"
      End If
      Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
      Exit Sub
    End If

    '* Add item to list of result
    
    'Set ListItem = ListResult.ListItems.Add
    'ListItem.Text = objNSearch.UserID
    'ListItem.SubItems(1) = "-"
    'ListItem.SubItems(2) = "-"
    'ListItem.SubItems(3) = "-"
    'Set ListItem = Nothing

End Sub
'Private Sub Register_Finger()
'
'    Dim i, j As Integer
'    Dim nUserID As Long
'    Dim szFir As String
'    Dim ListItem As ListItem
'
'    Dim objResult As ICandidateList         ' CandidateList or Result object
'
'    nUserID = 0
'    szFir = ""
'
'    ' Get User ID
'    If Not IsNumeric(txtUserID.Text) Then
'        MsgBox "User ID must be have numeric type and greater than 0.", vbOKOnly, "Error"
'        Exit Sub
'    End If
'
'    nUserID = CLng(txtUserID.Text)
'
'    ' Get FIR data
'    Call objDevice.Open(NBioAPI_DEVICE_ID_AUTO_DETECT)
'    Call objExtraction.Enroll(Null)
'    If objExtraction.ErrorCode <> NBioAPIERROR_NONE Then
'        MsgBox objExtraction.ErrorDescription & " [" & objExtraction.ErrorCode & "]"
'        Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
'        Exit Sub
'    End If
'
'    Call objDevice.Close(NBioAPI_DEVICE_ID_AUTO_DETECT)
'
'    szFir = objExtraction.TextEncodeFIR
'
'    ' Regist FIR to NSearch DB
'    Call objNSearch.AddFIR(szFir, nUserID)
'    If objNSearch.ErrorCode <> NBioAPIERROR_NONE Then
'        MsgBox objNSearch.ErrorDescription & " [" & objNSearch.ErrorCode & "]"
'        Exit Sub
'    End If
'
'
'    ' Add item to list of SearchDB
'    For Each objResult In objNSearch
'
'        Set ListItem = ListSearchDB.ListItems.Add
'        ListItem.Text = objResult.UserID
'        ListItem.SubItems(1) = objResult.FingerID
'        ListItem.SubItems(2) = objResult.SampleNumber
'        Set ListItem = Nothing
'
'    Next
'
'    txtUserID.Text = CLng(txtUserID.Text) + 1
'    ListSearchDB.SelectedItem.Selected = False
'
'End Sub

