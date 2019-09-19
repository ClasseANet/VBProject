VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsFile"
Attribute VB_Ext_KEY = "Member0" ,"clsFile"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


'local variable to hold collection
Private mCol As Collection
Public Function LoadFiles( _
                        ByVal sFileSpec As String, _
                        Optional ByVal bRecursive As Boolean = False, _
                        Optional ByVal pb) _
                As Long
    
    Static lLevel As Long
    lLevel = lLevel + 1
    If lLevel = 1 Then
        Set mCol = Nothing
        Set mCol = New Collection
    End If
    
    On Error Resume Next
    pb.Min = 0
    pb.value = 0
    pb.Max = fileCount(sFileSpec) + 1
    On Error GoTo 0
    
    Dim lFind As Long, lMatch As Long
    Dim tInfo As WIN32_FIND_DATA
    Dim oFile As File
    
    ' Scan Subdirs First
    If bRecursive Then
        Dim sDirSpec As String
        Dim sSpec As String
        sSpec = sFileName(sFileSpec, efpFileNameAndExt)
        sDirSpec = sAppend(sFileName(sFileSpec, efpFilePath), "\")
        lFind = FindFirstFile(sDirSpec & "*.*", tInfo)
        lMatch = 99
        Do While lFind > 0 And lMatch > 0
            If (tInfo.dwFileAttributes And efaDIRECTORY) > 0 Then
                Dim sDirName As String
                sDirName = sNT(tInfo.cFileName)
                If sDirName <> "." And sDirName <> ".." Then
                    LoadFiles sDirSpec & sAppend(sDirName, "\") & sSpec, bRecursive, pb
                End If
            End If
            lMatch = FindNextFile(lFind, tInfo)
        Loop
        FindClose lFind
    End If
    
    lFind = FindFirstFile(sFileSpec, tInfo)
    lMatch = 99
    Do While lFind > 0 And lMatch > 0
        If Not (tInfo.dwFileAttributes And efaDIRECTORY) > 0 Then
            Set oFile = New File
            oFile.FileName = sFileName(sFileSpec, efpFilePath) & sNT(tInfo.cFileName)
            mCol.Add oFile
        End If
        lMatch = FindNextFile(lFind, tInfo)
        On Error Resume Next
        pb.value = pb.value + 1
        pb.Refresh
        On Error GoTo 0
    Loop
    FindClose lFind
    LoadFiles = mCol.Count
    lLevel = lLevel - 1
    
End Function

Public Function Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Function
Public Function Add(ByVal sFileName As String, Optional sKey As String) As File
    'create a new object
    Dim objNewMember As File
    Set objNewMember = New File
    objNewMember.FileName = sFileName
    If Len(sKey) = 0 Then
        On Error Resume Next
        mCol.Add objNewMember, sFileName
        On Error GoTo 0
    Else
        mCol.Add objNewMember, sKey
    End If

    'set the properties passed into the method
'    objNewMember.sPathNetwork = sPathNetwork
'    objNewMember.sPathDOS = sPathDOS
'    objNewMember.sExtension = sExtension
'    objNewMember.sName = sName
'    objNewMember.dLastModified = dLastModified
'    objNewMember.lSize = lSize
'    If IsObject(lAttributes) Then
'        Set objNewMember.lAttributes = lAttributes
'    Else
'        objNewMember.lAttributes = lAttributes
'    End If
'    objNewMember.sINIsection = sINIsection
'    objNewMember.sFilename = sFilename
'    objNewMember.sFullPathNetwork = sFullPathNetwork
'    objNewMember.sFullPathDOS = sFullPathDOS
'    If Len(sKey) = 0 Then
'        mCol.Add objNewMember
'    Else
'        mCol.Add objNewMember, sKey
'    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As File
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As File
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

