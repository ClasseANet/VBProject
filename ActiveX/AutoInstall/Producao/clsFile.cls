VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Private Enum enumFileAttributes
'    efaARCHIVE = &H20
'    efaCOMPRESSED = &H800
'    efaDIRECTORY = &H10
'    efaHIDDEN = &H2
'    efaNORMAL = &H80
'    efaREADONLY = &H1
'    efaSYSTEM = &H4
'    efaTEMPORARY = &H100
'End Enum
'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Public Property Get FixedDisk() As Boolean
    FixedDisk = FileOpenStructure(mvarFileName).fFixedDisk
End Property
Public Property Get Created() As Date
    Created = TimeFileToDate(FileInformation(mvarFileName).ftCreationTime)
End Property
Public Property Get LastAccessed() As Date
    LastAccessed = TimeFileToDate(FileInformation(mvarFileName).ftLastAccessTime)
End Property
Public Property Get ExpandedName() As String
    ExpandedName = FileExpandedName(mvarFileName)
End Property
Public Property Get ShortName() As String
    ShortName = FileShortName(mvarFileName)
End Property
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Public Property Get FileName() As String
Attribute FileName.VB_UserMemId = 0
    FileName = mvarFileName
End Property
'''Public Function Edit() As Boolean
'''End Function
'''
'''Public Function Move(sDestination As String, Optional bVerbose As Boolean) As Boolean
'''End Function
'''
'''Public Function Copy(sNewFileName As String, Optional bVerbose As Boolean) As Boolean
'''End Function
'''
'''Public Function Delete() As Boolean
'''End Function
'''
'''Public Function Launch() As Boolean
'''End Function
'''
'''Public Function Rename(ByVal sNewFileName As String, Optional ByVal bVerbose As Boolean) As Boolean
'''End Function
Public Property Let Attributes(ByVal vData As Long)
    SetFileAttributes Me.FileName, vData
End Property
Public Property Get Attributes() As Long
    Attributes = FileAttributes(mvarFileName)
End Property
Public Property Get Size() As Long
    On Error Resume Next
    Size = FileLen(mvarFileName)
    If Err.Number <> 0 Then
        Size = FileLength(mvarFileName)
    End If
End Property
Public Property Get LastModified() As Date
    LastModified = TimeFileToDate(FileInformation(FileShortName(mvarFileName)).ftLastWriteTime)
End Property
Public Property Get Name() As String
    Name = sFileName(mvarFileName, efpFileName)
End Property
Public Property Get NameAndExtension() As String
    NameAndExtension = sFileName(mvarFileName, efpFileNameAndExt)
End Property
Public Property Get Extension() As String
    Extension = sFileName(mvarFileName, efpFileExt)
End Property
Public Property Get Path() As String
    Path = sFileName(mvarFileName, efpFilePath)
End Property
Public Property Get DriveLetter() As String
    Dim lNo As Long
    lNo = PathGetDriveNumber(mvarFileName)
    If lNo > -1 Then
        DriveLetter = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", 1 + lNo, 1) & ":"
    Else
        DriveLetter = ""
    End If
    
End Property
Public Property Get PathRoot() As String
    Dim sRet As String
    
    sRet = FileRoot(mvarFileName)
    If Right(sRet, 1) <> "\" And Trim(sRet) <> "" Then
        sRet = sRet & "\"
    End If
    PathRoot = sRet
End Property
Public Property Get NetworkPath() As Boolean
    NetworkPath = PathIsNetworkPath(mvarFileName)
End Property
Public Property Get UNCServerShare() As Boolean
    UNCServerShare = PathIsUNCServerShare(mvarFileName)
End Property
Public Property Get UNCServer() As Boolean
    UNCServer = PathIsUNCServer(mvarFileName)
End Property
Public Property Get UNC() As Boolean
    UNC = PathIsUNC(mvarFileName)
End Property
Public Property Get VolumeType() As Long
    VolumeType = GetDriveType(Me.PathRoot)
End Property
Public Property Get VolumeSerialNo() As Long
    VolumeSerialNo = FileInformation(mvarFileName).dwVolumeSerialNumber
End Property
Public Property Get VolumeName() As String
    VolumeName = VolumeInformation(Me.PathRoot).sVolumeName
End Property
Public Property Get Version() As String
   Dim MyInstall As New AutoInstall
   Version = L_GetFileVersion(mvarFileName)
   Set MyInstall = Nothing
End Property
Private Function L_GetFileVersion(ByVal pFilename As String, Optional ByVal fIsRemoteServerSupportFile) As String
   Dim sVerInfo As VERINFO
   Dim strVer As String
   
   On Error GoTo GFVError
   
   If IsMissing(fIsRemoteServerSupportFile) Then
      fIsRemoteServerSupportFile = False
   End If
   
   '
   'Get the file version into a VERINFO struct, and then assemble a version string
   'from the appropriate elemen
   '
   If GetFileVerStruct(pFilename, sVerInfo, fIsRemoteServerSupportFile) = True Then
      strVer = ""
      strVer = strVer & Format$(sVerInfo.nMSHi, "000") & "."
      strVer = strVer & Format$(sVerInfo.nMSLo, "000") & "."
      strVer = strVer & Format$(sVerInfo.nLSHi, "000") & "."
      strVer = strVer & Format$(sVerInfo.nLSLo, "000")
      L_GetFileVersion = strVer
   Else
      L_GetFileVersion = ""
   End If
   
   Exit Function
    
GFVError:
   L_GetFileVersion = ""
   If Err = 48 Then
      MsgBox "ERRO : " & Err & " - " & Error
   End If
   Err = 0
End Function



