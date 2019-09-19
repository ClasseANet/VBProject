Attribute VB_Name = "BlType"
Option Explicit

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Type VERINFO                  'Version FIXEDFILEINFO
    strPad1 As Long           'Pad out struct version
    strPad2 As Long           'Pad out struct signature
    nMSLo As Integer          'Low word of ver # MS DWord
    nMSHi As Integer          'High word of ver # MS DWord
    nLSLo As Integer          'Low word of ver # LS DWord
    nLSHi As Integer          'High word of ver # LS DWord
    strPad3(1 To 16) As Byte  'Skip some of VERINFO struct (16 bytes)
    FileOS As Long            'Information about the OS this file is targeted for.
    strPad4(1 To 16) As Byte  'Pad out the resto of VERINFO struct (16 bytes)
End Type
Public Type PointAPI   ' pt
  X As Long
  y As Long
End Type
'Public Type ZIPnames
'    s(0 To 99) As String
'End Type
Public Type ZIPUSERFUNCTIONS
    DLLPrnt As Long
    DLLPASSWORD As Long
    DLLCOMMENT As Long
    DLLSERVICE As Long
End Type
Public Type ZPOPT
    fSuffix As Long
    fEncrypt As Long
    fSystem As Long
    fVolume As Long
    fExtra As Long
    fNoDirEntries As Long
    fExcludeDate As Long
    fIncludeDate As Long
    fVerbose As Long
    fQuiet As Long
    fCRLF_LF As Long
    fLF_CRLF As Long
    fJunkDir As Long
    fRecurse As Long
    fGrow As Long
    fForce As Long
    fMove As Long
    fDeleteEntries As Long
    fUpdate As Long
    fFreshen As Long
    fJunkSFX As Long
    fLatestTime As Long
    fComment As Long
    fOffsets As Long
    fPrivilege As Long
    fEncryption As Long
    fRepair As Long
    flevel As Byte
    date As String ' 8 bytes long
    szRootDir As String ' up to 256 bytes long
End Type
'Salve BMP to JPG
Public Type imgdes
  ibuff As Long
  stx As Long
  sty As Long
  endx As Long
  endy As Long
  buffwidth As Long
  palette As Long
  colors As Long
  imgtype As Long
  bmh As Long
  hBitmap As Long
End Type
Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

