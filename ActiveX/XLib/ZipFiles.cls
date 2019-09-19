VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipFiles"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'
'
Public Enum ZTranslate
    CRLFtoLF = 1
    LFtoCRLF = 2
End Enum

'
' Collection of Files to Zip
'
Private mCollection As Collection
'
' Recurse Folders ?
'
Private miRecurseFolders As Integer
'
' Zip File Name
'
Private msZipFileName As String
'
' Encryption ?
'
Private miEncrypt As Integer
'
' System Files
'
Private miSystem As Integer
'
' Root Directory
'
Private msRootDirectory As String
'
' Verbose Zip
'
Private miVerbose As Integer
'
' Quiet Zip
'
Private miQuiet As Integer
'
' Translate CRLF / LF Chars
'
Private miTranslateCRLF As ZTranslate
'
' Updating Existing Zip ?
'
Private miUpdateZip As Integer

Private Sub Class_Initialize()
'
' Initialise the collection
'
    Set mCollection = New Collection
'
' We have to add in a dummy file into the collection because
' the Zip routines fall over otherwise.
'
' I think this is a bug, but it's not documented anywhere
' on the InfoZip website.
'
' The Zip process *always* fails on the first file,
' regardless of whether it's a valid file or not!
'
    mCollection.Add "querty", "querty"
    miEncrypt = 0
    miSystem = 0
    msRootDirectory = "\"
    miQuiet = 0
    miUpdateZip = 0
    
End Sub

Private Sub Class_Terminate()
'
' Terminate the collection
'
    Set mCollection = Nothing
End Sub

Public Property Get RecurseFolders() As Boolean
    RecurseFolders = miRecurseFolders = 1
End Property

Public Property Let RecurseFolders(ByVal bRecurse As Boolean)
    miRecurseFolders = IIf(bRecurse, 1, 0)
End Property

Public Property Get ZipFileName() As String
    ZipFileName = msZipFileName
End Property

Public Property Let ZipFileName(ByVal sZipFileName As String)
    msZipFileName = sZipFileName '& vbNullChar
End Property

Public Property Get Encrypted() As Boolean
    Encrypted = miEncrypt = 1
End Property

Public Property Let Encrypted(ByVal bEncrypt As Boolean)
    miEncrypt = IIf(bEncrypt, 1, 0)
End Property

Public Property Get IncludeSystemFiles() As Boolean
    IncludeSystemFiles = miSystem = 1
End Property

Public Property Let IncludeSystemFiles(ByVal bInclude As Boolean)
    miSystem = IIf(bInclude, 1, 0)
End Property

Public Property Get ZipFileCount() As Long
    If mCollection Is Nothing Then
        ZipFileCount = 0
    Else
        ZipFileCount = mCollection.Count - 1
    End If
End Property


Public Property Get RootDirectory() As String
    RootDirectory = msRootDirectory
End Property

Public Property Let RootDirectory(ByVal sRootDir As String)
    msRootDirectory = sRootDir ' & vbNullChar
End Property

Public Property Get UpdatingZip() As Boolean
    UpdatingZip = miUpdateZip = 1
End Property

Public Property Let UpdatingZip(ByVal bUpdating As Boolean)
    miUpdateZip = IIf(bUpdating, 1, 0)
End Property

Public Function AddFile(ByVal sFileName As String)
    Dim lCount As Long
    Dim sFile As String
    
    On Error Resume Next
    
    sFile = mCollection.Item(sFileName)
    
    If Len(sFile) = 0 Then
        Err.Clear
        On Error GoTo 0
        mCollection.Add sFileName, sFileName
    Else
        On Error GoTo 0
        Err.Raise vbObjectError + 2001, "FWZip::AddFile", "File is already in Zip List"
    End If
    
End Function

Public Function RemoveFile(ByVal sFileName As String)
    Dim lCount As Long
    Dim sFile As String
    
    On Error Resume Next
    
    sFile = mCollection.Item(sFileName)
    
    If Len(sFile) = 0 Then
        Err.Raise vbObjectError + 2002, "CGZip::RemoveFile", "File is not in Zip List"
    Else
        mCollection.Remove sFileName
    End If
    
End Function
Public Function UnZipFile(pPath As String, pFile As String, Optional pPathDest As String, Optional pHonorDir As Boolean = True) As Boolean
   On Error GoTo TrataErro
   'Dim oUnZip As CGUnzipFiles
   
   'Call RegServer(App.Path & "\Unzip.dll")
        
   If Trim(pPathDest) = "" Then pPathDest = pPath
   Call ClsAmbiente.CriarDiretorio(pPathDest)
   
'    Set oUnZip = New CGUnzipFiles
'    With oUnZip
'      .ZipFileName = ResolvePathName(pPath) & pFile
'      .ExtractDir = ResolvePathName(pPathDest) 'GetTempPathName
'      .HonorDirectories = pHonorDir
'      If .UnZip <> 0 Then
'         pMessage = .GetLastMessage
'         DescompactarPacote = True
'      End If
'    End With
'    Set oUnZip = Nothing
   UnZipFile = VBUnzip(ClsAmbiente.ResolvePathName(pPath) & pFile, pPathDest, 1, 0, 0, 1, 1, 0, 1, 0, 0, IIf(pHonorDir, 1, 0), 0, 0, 0, 0, 0)
    
    'MsgBox "\ZIPTEST.ZIP Extracted Successfully to " & GetTempPathName

    Exit Function

TrataErro:
    MsgBox Err.Number & " " & "Form1::cmdUnZip_Click" & " " & Err.Description
End Function
Public Function ZipFiles(pCollection As Collection, pZipFileName As String) As Long
   ZipFiles = MakeZipFile(pCollection, pZipFileName)
End Function
Public Function MakeZipFile(pCollection As Collection, pZipFileName As String) As Long
   Dim zFileArray As ZIPnames
   Dim sFileName As Variant
   Dim lFileCount As Long
   Dim iIgnorePath As Integer
   Dim iRecurse As Integer
   
   On Error GoTo vbErrorHandler
    
   NumArquivos = NumArquivos + 0
   
   lFileCount = 0
    
   For Each sFileName In pCollection
        zFileArray.s(lFileCount) = sFileName
        lFileCount = lFileCount + 1
    Next
    
    MakeZipFile = VBZip(CInt(lFileCount), pZipFileName, _
        zFileArray, iIgnorePath, miRecurseFolders, _
        miUpdateZip, 0, msRootDirectory, _
        0, 1, 0)
    

    Exit Function

vbErrorHandler:
    MakeZipFile = -99
    Err.Raise Err.Number, "FWZipFiles::MakeZipFile", Err.Description

End Function

Public Function GetLastMessage() As String
    GetLastMessage = msOutput
End Function
'***********************************************************************
'***********************************************************************
'***********************************************************************
Private Function VBZip(argc As Integer, zipname As String, _
        mynames As ZIPnames, junk As Integer, _
        recurse As Integer, updat As Integer, _
        freshen As Integer, basename As String, _
        Optional Encrypt As Integer = 0, _
        Optional IncludeSystem As Integer = 0, _
        Optional IgnoreDirectoryEntries As Integer = 0, _
        Optional Verbose As Integer = 0, _
        Optional Quiet As Integer = 0, _
        Optional CRLFtoLF As Integer = 0, _
        Optional LFtoCRLF As Integer = 0, _
        Optional Grow As Integer = 0, _
        Optional Force As Integer = 0, _
        Optional iMove As Integer = 0, _
        Optional DeleteEntries As Integer = 0) As Long
    
    Dim hMem As Long, xx As Integer
    Dim retcode As Long
    Dim MYUSER As ZIPUSERFUNCTIONS
    Dim MYOPT As ZPOPT
    
    On Error Resume Next ' nothing will go wrong :-)
    
    msOutput = ""
    
    ' Set address of callback functions
    MYUSER.DLLPrnt = FnPtr(AddressOf DLLPrnt)
    MYUSER.DLLPASSWORD = FnPtr(AddressOf DllPass)
    MYUSER.DLLCOMMENT = FnPtr(AddressOf DllComm)
    MYUSER.DLLSERVICE = 0& ' not coded yet :-)
'    retcode = ZpInit(MYUSER)
    
    ' Set zip options
    MYOPT.fSuffix = 0        ' include suffixes (not yet implemented)
    MYOPT.fEncrypt = Encrypt     ' 1 if encryption wanted
    MYOPT.fSystem = IncludeSystem        ' 1 to include system/hidden files
    MYOPT.fVolume = 0        ' 1 if storing volume label
    MYOPT.fExtra = 0         ' 1 if including extra attributes
    MYOPT.fNoDirEntries = IgnoreDirectoryEntries  ' 1 if ignoring directory entries
    MYOPT.fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
    MYOPT.fIncludeDate = 0   ' 1 if including files earlier than a specified date
    MYOPT.fVerbose = Verbose       ' 1 if full messages wanted
    MYOPT.fQuiet = Quiet         ' 1 if minimum messages wanted
    MYOPT.fCRLF_LF = CRLFtoLF        ' 1 if translate CR/LF to LF
    MYOPT.fLF_CRLF = LFtoCRLF ' 1 if translate LF to CR/LF
    MYOPT.fJunkDir = junk    ' 1 if junking directory names
    MYOPT.fRecurse = recurse ' 1 if recursing into subdirectories
    MYOPT.fGrow = Grow          ' 1 if allow appending to zip file
    MYOPT.fForce = Force         ' 1 if making entries using DOS names
    MYOPT.fMove = iMove          ' 1 if deleting files added or updated
    MYOPT.fDeleteEntries = DeleteEntries ' 1 if files passed have to be deleted
    MYOPT.fUpdate = updat    ' 1 if updating zip file--overwrite only if newer
    MYOPT.fFreshen = freshen ' 1 if freshening zip file--overwrite only
    MYOPT.fJunkSFX = 0       ' 1 if junking sfx prefix
    MYOPT.fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
    MYOPT.fComment = 0       ' 1 if putting comment in zip file
    MYOPT.fOffsets = 0       ' 1 if updating archive offsets for sfx Files
    MYOPT.fPrivilege = 0     ' 1 if not saving privelages
    MYOPT.fEncryption = 0    'Read only property!
    MYOPT.fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
    MYOPT.flevel = 0         ' compression level - should be 0!!!
    MYOPT.date = vbNullString ' "12/31/79"? US Date?
    MYOPT.szRootDir = UCase$(basename)
    
    retcode = ZpInit(MYUSER)
    ' Set options
    retcode = ZpSetOptions(MYOPT)
    
    ' ZCL not needed in VB
    ' MYZCL.argc = 2
    ' MYZCL.filename = "c:\wiz\new.zip"
    ' MYZCL.fileArray = MYNAMES
    
    ' Go for it!
    
    retcode = ZpArchive(argc, zipname, mynames)
    
    VBZip = retcode
End Function
Public Function VBUnzip(ByRef sZipFileName, ByRef sUnzipDirectory As String, _
    ByRef iExtractNewer As Integer, _
    ByRef iSpaceUnderScore As Integer, _
    ByRef iPromptOverwrite As Integer, _
    ByRef iQuiet As Integer, _
    ByRef iWriteStdOut As Integer, _
    ByRef iTestZip As Integer, _
    ByRef iExtractList As Integer, _
    ByRef iExtractOnlyNewer As Integer, _
    ByRef iDisplayComment As Integer, _
    ByRef iHonorDirectories As Integer, _
    ByRef iOverwriteFiles As Integer, _
    ByRef iConvertCR_CRLF As Integer, _
    ByRef iVerbose As Integer, _
    ByRef iCaseSensitivty As Integer, _
    ByRef iPrivilege As Integer) As Long


On Error GoTo vbErrorHandler

    
    Dim lRet As Long
    
    Dim UZDCL As DCLIST
    Dim UZUSER As USERFUNCTION
    Dim UZVER As UZPVER
    Dim uExcludeNames As UNZIPnames
    Dim uZipNames     As UNZIPnames
    
    msOutput = ""
    
    uExcludeNames.uzFiles(0) = vbNullString
    uZipNames.uzFiles(0) = vbNullString
    
    uZipNumber = 0
    uZipMessage = vbNullString
    uZipInfo = vbNullString
    uVBSkip = 0
    
    With UZDCL
        .ExtractOnlyNewer = iExtractOnlyNewer
        .SpaceToUnderScore = iSpaceUnderScore
        .PromptToOverwrite = iPromptOverwrite
        .fQuiet = iQuiet
        .ncflag = iWriteStdOut
        .ntflag = iTestZip
        .nvflag = iExtractList
        .nUflag = iExtractNewer
        .nzflag = iDisplayComment
        .ndflag = iHonorDirectories
        .noflag = iOverwriteFiles
        .naflag = iConvertCR_CRLF
        .nZIflag = iVerbose
        .C_flag = iCaseSensitivty
        .fPrivilege = iPrivilege
        .Zip = sZipFileName
        .ExtractDir = sUnzipDirectory
    End With
    
    With UZUSER
        .UZDLLPrnt = FnPtr(AddressOf UZDLLPrnt)
        .UZDLLSND = 0&
        .UZDLLREPLACE = FnPtr(AddressOf UZDLLRep)
        .UZDLLPASSWORD = FnPtr(AddressOf UZDLLPass)
        .UZDLLMESSAGE = FnPtr(AddressOf UZReceiveDLLMessage)
        .UZDLLSERVICE = FnPtr(AddressOf UZDLLServ)
    End With
    
    With UZVER
        .structlen = Len(UZVER)
        .beta = Space$(9) & vbNullChar
        .date = Space$(19) & vbNullChar
        .zlib = Space$(9) & vbNullChar
    End With
    
    UzpVersion2 UZVER
    
    lRet = Wiz_SingleEntryUnzip(0, uZipNames, 0, uExcludeNames, UZDCL, UZUSER)
    VBUnzip = lRet
    

    Exit Function

vbErrorHandler:
    Err.Raise Err.Number, "CodeModule::VBUnzip", Err.Description

End Function


