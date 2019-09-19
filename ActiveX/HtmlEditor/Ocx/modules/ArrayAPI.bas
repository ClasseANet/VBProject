Attribute VB_Name = "mArrayAPI"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft Visual Html Editor
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit

Public Type OLEVARIANT
    vt As Integer
    wReserved1 As Integer
    wReserved2 As Integer
    wReserved3 As Integer
    lVal As Integer
    iVal As Integer
    bstrVal As Long
    pUnkVal As Long
    pArray As Long
    pvRecord As Long
    pRecInfo As Long
End Type

' =============================================================================================
' SAFEARRAY API FUNCTIONS
' =============================================================================================
Type SAFEARRAYBOUNDTYPE
    cElements As Long
    lLbound As Long
End Type

Type SAFEARRAYTYPE
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgsabound(0 To 1) As SAFEARRAYBOUNDTYPE
End Type

'Declare Function SafeArrayAccessData Lib "OLEAUT32.DLLL" Alias " SafeArrayAccessData" (ByVal psa As Long, ByRef ppvData As Long) As Long
'Declare Function SafeArrayAllocData Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeAllocDescriptor Lib "oleaut32.dll" (ByVal cDims As Long, ByRef ppsaOut As Long) As Long
'Declare Function SafeAllocDescriptorEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef ppsaOut As Long) As Long
'Declare Function SafeArrayCopy Lib "oleaut32.dll" (ByVal psa As Long, ByRef ppsaOut As Long) As Long
'Declare Function SafeArrayCopyData Lib "oleaut32.dll" (ByVal psaSource As Long, ByVal psaTarget As Long) As Long
'Declare Function SafeArrayCreate Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUNDTYPE) As Long
'Declare Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUNDTYPE, ByVal pvExtra As Long) As Long
'Declare Function SafeArrayCreateVector Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long) As Long
'Declare Function SafeArrayCreateVectorEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long, ByVal pvExtra As Long) As Long
'Declare Function SafeArrayDestroy Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeArrayDestroyData Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Long, ByVal rgIndices As Long, ByVal pv As Long) As Long
'Declare Function SafeArrayGetElemsize Lib "oleaut32.dll" (ByVal psa As Long) As Long
''Declare Function SafeArrayGetIID Lib "OLEAUT32.DLL" (ByVal psa As Long, ByRef pguid As Guid) As Long
'Declare Function SafeArrayGetLBound Lib "oleaut32.dll" (ByVal psa As Long, ByVal nDim As Long, ByRef plLbound As Long) As Long
'Declare Function SafeArrayGetUBound Lib "oleaut32.dll" (ByVal psa As Long, ByVal nDim As Long, ByRef plUbound As Long) As Long
'Declare Function SafeArrayGetRecordInfo Lib "oleaut32.dll" Alias " SafeArrayGetRecordInfo" (ByVal psa As Long, ByRef prinfo As Long) As Long
'Declare Function SafeArrayGetVartype Lib "oleaut32.dll" Alias " SafeArrayGetVartype" (ByVal psa As Long, ByVal pvt As Long) As Long
'Declare Function SafeArrayLock Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeArrayPtrOfIndex Lib "oleaut32.dll" Alias " SafeArrayPtrOfIndex" (ByVal psa As Long, ByVal rgIndices As Long, ByRef ppvData As Long) As Long
'Declare Function SafeArrayPutElement Lib "oleaut32.dll" (ByVal psa As Long, ByVal rgIndices As Long, ByVal pv As Long) As Long
'Declare Function SafeArrayRedim Lib "oleaut32.dll" (ByVal psa As Long, ByRef psaboundNew As SAFEARRAYBOUNDTYPE) As Long
''Declare Function SafeArraySetIID Lib "OLEAUT32.DLL" (ByVal psa As Long, ByRef pguid As Guid) As Long
'Declare Function SafeArraySetRecordInfo Lib "oleaut32.dll" (ByVal psa As Long, ByRef psaboundNew As SAFEARRAYBOUNDTYPE) As Long
'Declare Function SafeArrayUnaccessData Lib "oleaut32.dll" (ByVal psa As Long) As Long
'Declare Function SafeArrayUnlock Lib "oleaut32.dll" (ByVal psa As Long) As Long
' =============================================================================================
' =============================================================================================
'http://groups-beta.google.com/group/hr.comp.programiranje.vb/browse_thread/thread/c3ee538b7940116/7f9134d8a007b513?q=SafeArrayGetElement+byval&rnum=4&hl=en#7f9134d8a007b513

Public Declare Function ArrayCreate Lib "oleaut32.dll" Alias "SafeArrayCreate" (ByVal vartypes As Long, ByVal Dimension As Long, boundData As Any) As Long
Public Declare Function ArrayDestroy Lib "oleaut32.dll" Alias "SafeArrayDestroyDescriptor" (Arrays As Any) As Long
Public Declare Function ArrayAllocate Lib "oleaut32.dll" Alias "SafeArrayAllocDescriptorEx" (ByVal vartypes As Long, ByVal Dimension As Long, Arrays As Any) As Long
Public Declare Function ArrayAllocateData Lib "oleaut32.dll" Alias "SafeArrayAllocData" (Arrays As Any) As Long
Public Declare Function ArrayRedim Lib "oleaut32.dll" Alias "SafeArrayRedim" (Arrays As Any, boundData As Any) As Long
Public Declare Function ArrayLock Lib "oleaut32.dll" Alias "SafeArrayLock" (Arrays As Any) As Long
Public Declare Function ArrayUnLock Lib "oleaut32.dll" Alias "SafeArrayUnlock" (Arrays As Any) As Long
Public Declare Function ArrayAccess Lib "oleaut32.dll" Alias "SafeArrayAccessData" (Arrays As Any, PointerToFirstElement As Long) As Long
Public Declare Function ArrayUnAccess Lib "oleaut32.dll" Alias "SafeArrayUnaccessData" (Arrays As Any) As Long
Public Declare Function ArrayPut Lib "oleaut32.dll" Alias "SafeArrayPutElement" (Arrays As Any, Element As Long, Data As Any) As Long
Public Declare Function ArrayGet Lib "oleaut32.dll" Alias "SafeArrayGetElement" (Arrays As Any, Element As Long, Data As Any) As Long
Public Declare Function ArrayClone Lib "oleaut32.dll" Alias "SafeArrayCopy" (Arrays As Any, NewArrays As Any) As Long
Public Declare Function ArrayCopy Lib "oleaut32.dll" Alias "SafeArrayCopyData" (SourceArrays As Any, DestinationArrays As Any) As Long
Public Declare Function ArrayElemenPointAPIer Lib "oleaut32.dll" Alias "SafeArrayPtrOfIndex" (Arrays As Any, Element As Long, PointerToData As Long) As Long
Public Declare Function ArrayDelete Lib "oleaut32.dll" Alias "SafeArrayDestroyData" (Arrays As Any) As Long
Public Declare Function ArrayDim Lib "oleaut32.dll" Alias "SafeArrayGetDim" (Arrays As Any) As Long
Public Declare Function ArrayLBound Lib "oleaut32.dll" Alias "SafeArrayGetLBound" (Arrays As Any, ByVal wDim As Long, dataLBound As Long) As Long
Public Declare Function ArrayUBound Lib "oleaut32.dll" Alias "SafeArrayGetUBound" (Arrays As Any, ByVal wDim As Long, dataLBound As Long) As Long
Public Declare Function ArrayElements Lib "oleaut32.dll" Alias "SafeArrayGetElemsize" (Arrays As Any) As Long
Public Declare Function ArrayCloneToVB Lib "oleaut32.dll" Alias "SafeArrayCopy" (Arrays As Any, VBArrays() As Any) As Long
Public Declare Function ArrayGetType Lib "oleaut32.dll" Alias "SafeArrayGetVartype" (Arrays As Any, ArrType As Long) As Long
' =============================================================================================
' =============================================================================================

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' =============================================================================================
Public Sub ParseOleVariantStrArray(varArray As Variant, Elements() As String)
'
' Given a stream of bytes that represent a Microsoft SafeArray, wade through
'the garbage and get the actual data found in the array.  Write the data to disk.

'      SafeArray looks like
'typedef struct FARSTRUCT tagSAFEARRAY {
'    unsigned short cDims;           // Count of dimensions in this array.
'    unsigned short fFeatures;     // Flags used by the SafeArray
'                                    // routines documented below.
'    unsigned long cbElements; // Size of an element of the array.
'                               // Does not include size of
'                               // pointed-to data.
'    unsigned long cLocks;  // Number of times the array has been
'    void HUGEP* pvData;   // Pointer to the data.
'    SAFEARRAYBOUND rgsabound[1]; // One bound for each dimension.
'} SAFEARRAY;
'SAFEARRAYBOUND is a two-longword structure, the first 32 bits hold the # of
'elements in the array, the second 32 bits hold the lower bound of the array.
'There is one structure for every dimension of the array.
'http://groups-beta.google.com/group/microsoft.public.vc.vcce/browse_thread/thread/7f2bfbdb6ef17e98/a90475b71136f83b?q=OleSafeArray+visual+basic&rnum=2&hl=en#a90475b71136f83b

'Public Type OLEVARIANT
'    vt As Integer 'variable type
'    wReserved1 As Integer
'    wReserved2 As Integer
'    wReserved3 As Integer
'    lVal As Integer
'    iVal As Integer
'    bstrVal As Long
'    pUnkVal As Long
'    pArray As Long
'    pvRecord As Long
'    pRecInfo As Long
'End Type

    Dim pArray As Long
    Dim ppArray As Long
    Dim ppVarStruct As Long
    
    Dim lLbound As Long
    Dim lUbound As Long
    
    Dim Element As String
    Dim X As Long
    Dim Ret As Long
    
    '----------------------------------------------------------------
    ' First get a pointer to the variant.
    ppVarStruct = VarPtr(varArray)

    ' Get the pointer to the data *inside* the variant.  The VARTYPE is 2 bytes,
    ' and each of the three reserved words are also 2 bytes, giving 8 bytes total.
    CopyMemory ppArray, ByVal ppVarStruct + 8, 4
    '--------------------------------------------------
    ArrayLBound ByVal ppArray, ByVal 1&, ByVal VarPtr(lLbound)
    ArrayUBound ByVal ppArray, ByVal 1&, ByVal VarPtr(lUbound)
    'ArrayGetType ByVal ppArray, b
    '------------------------------------
    Element = Space(200)
    ReDim Elements(lUbound + 1)
    For X = lLbound To lUbound
        Ret = ArrayGet(ByVal ppArray, X, Element)
        Elements(X) = StrConv(Element, vbFromUnicode)
    Next X
    
    ArrayDelete ByVal ppArray
    
End Sub
'====================================================================


