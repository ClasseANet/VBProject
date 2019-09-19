VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SOBRE"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Option Explicit
'with mvarME
Private MemStat As MemoryStatus
Private MemoryLoad As Long
Private TotalPhys As Long
Private AvailPhys As Long
Private TotalPageFile As Long
Private AvailPageFile As Long
Private TotalVirtual As Long
Private AvailVirtual As Long
  
Private disco As DISK_INFO
Private volume As VOL_INFO

Private sBuffer As String
Private sDrives As String
Private sDriveID As String
Private lDrive As Long
Private Enum ConstDRIVE
   DRIVE_ERROR = 0
   DRIVE_NOTPRESENT
   DRIVE_REMOVABLE
   DRIVE_FIXED
   DRIVE_REMOTE
   DRIVE_CDROM
   DRIVE_RAMDISK
End Enum

'End With

Private WithEvents mvarMe As frmSobre
Attribute mvarMe.VB_VarHelpID = -1
Private glngReturnStatus As Long
Private Const SUCCESS = 1&
Private Const FAILURE = 0&
Dim VersaoInfo As OSVERSIONINFO

Private mvarName As String 'local copy
Private mvarTitle As String 'local copy
Private mvarAnoDsvm As String 'local copy
Private mvarEmpresa As String 'local copy
Private mvarVersao As String 'local copy

'local variable(s) to hold property value(s)
Private mvarAplic As Object 'local copy
Public Property Let Aplic(ByVal vData As Object)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Aplicacao = 5
    Set mvarAplic = vData
End Property
Public Property Get Aplic() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Aplicacao
   Set Aplic = mvarAplic
End Property
Public Sub AboutShow(vData As Object)
   Aplic = vData
   mvarMe.Show vbModal
'   frmSobre.Show vbModal
End Sub
Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
  ' Exemplos: CreateNewKey "TestKey", HKEY_CURRENT_USER
  '           CreateNewKey "TestKey\SubKey1\SubKey2", HKEY_LOCAL_MACHINE
  Dim hNewKey As Long         'handle to the new key
  Dim lRetVal As Long         'result of the RegCreateKeyEx function

'  lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
  lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, 0, &H3F, 0&, hNewKey, lRetVal)
  RegCloseKey (hNewKey)

End Sub
Function QueryValue(ByVal hPrincipal As Long, sKeyName As String, sValueName As String, Optional cMensagem As Variant) As String
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long            'handle of opened key
  Dim vValue As Variant       'setting of queried value

  If RegOpenKeyEx(hPrincipal, sKeyName, 0, KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
    If QueryValueEx(hKey, sValueName, vValue) = ERROR_SUCCESS Then
      QueryValue = vValue
    Else
      QueryValue = cMensagem
    End If
    
    RegCloseKey (hKey)
  Else
    QueryValue = cMensagem
  End If
  
End Function
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
  Dim cch As Long
  Dim lrc As Long
  Dim lType As Long
  Dim lValue As Long
  Dim sValue As String

  On Error GoTo QueryValueExError

  ' Determine the size and type of data to be read
  lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
  If lrc <> ERROR_NONE Then Error 5
  Select Case lType
    ' For strings
    Case REG_SZ:
      sValue = String(cch, 0)
      lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, _
        sValue, cch)
      If lrc = ERROR_NONE Then
        vValue = Left$(sValue, cch)
      Else
        vValue = Empty
      End If
      ' For DWORDS
    Case REG_DWORD:
      lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, _
        lValue, cch)
      If lrc = ERROR_NONE Then vValue = lValue
    Case Else
      'all other data types not supported
      lrc = -1
  End Select

QueryValueExExit:
  QueryValueEx = lrc
  Exit Function

QueryValueExError:
  Resume QueryValueExExit

End Function
Public Sub SetKeyValue(sKeyName As String, sValueName As String, _
  vValueSetting As Variant, lValueType As Long)
  ' Exemplo: SetKeyValue "TestKey\SubKey1", "StringValue", "Hello", REG_SZ
  Dim lRetVal As Long         'result of the SetValueEx function
  Dim hKey As Long         'handle of open key

  'open the specified key
  lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0, _
    KEY_ALL_ACCESS, hKey)
  
  lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
  
  RegCloseKey (hKey)

End Sub
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
  Dim lValue As Long
  Dim sValue As String
  
  Select Case lType
    Case REG_SZ
      sValue = vValue
      SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
        lType, sValue, Len(sValue))
    Case REG_DWORD
      lValue = vValue
      SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, _
        lType, lValue, 4)
  End Select

End Function
Private Function ComputerName() As String
   Dim lpBuffer As String * 255
   Dim lRet As Long
   
   lRet = GetComputerName(lpBuffer, 255)
   If lRet <> 0 Then
     ComputerName = UCase$(Left$(Trim$(lpBuffer), Len(Trim$(lpBuffer)) - 1))
     ComputerName = ClsDsr.EliminarString(ComputerName, Chr(0))
   Else
     ComputerName = "Não Fornecido."
   End If
End Function
Public Function ExtractNullTermString(ByVal NullTermString As String) As String
  'Looks for first null in the string and returns
  'everything left of the null

  If InStr(NullTermString, Chr(0)) Then
    ExtractNullTermString = Left$(NullTermString, InStr(NullTermString, Chr(0)) - 1)
  Else
    ExtractNullTermString = NullTermString
  End If

End Function
'---------------------------------------------------------------
' FabProcessador()
' Retorna o fabricante do processador da máquina.
'---------------------------------------------------------------
Public Function FabProcessador() As String
  Dim si As SystemInfo
  Dim cRetorno As String
  
  Call GetSystemInfo(si)
  
  Select Case (si.dwOemID)
    Case PROCESSOR_ARCHITECTURE_INTEL
      cRetorno = "Intel"
    Case PROCESSOR_ARCHITECTURE_MIPS
      cRetorno = "MIPS"
    Case PROCESSOR_ARCHITECTURE_ALPHA
      cRetorno = "Alpha"
    Case PROCESSOR_ARCHITECTURE_PPC
      cRetorno = "PowerPC"
    Case PROCESSOR_ARCHITECTURE_UNKNOWN
      cRetorno = "Desconhecido"
  End Select
  
  FabProcessador = cRetorno
  
End Function
Public Function FreeDrive() As String
  Dim DriveNum As Integer
  Dim FirstFreeDrive As String
  Dim FirstDrive As Integer
  Dim NextDrive As String
  
  On Error Resume Next
  
  DriveNum = -1
  
  Do
    DriveNum = DriveNum + 3 '1   ' start at drive zero.
    NextDrive = Chr$(DriveNum + 65) + ":\"
    FirstDrive = ClsAutoIns.GetDriveType(NextDrive)
    If Err <> 0 Then
       FirstDrive = ClsAutoIns.GetDriveType(CInt(DriveNum + 65))
    End If
    
    ' GetDriveType returns zero if it cannot determine drive
    ' type or returns 1 if the specified drive does not exist.
  Loop Until FirstDrive = 1
  ' DriveNum of 0 means Drive A, 1=B, 2=C, 3=D, 4=E, 5=F, and so on:
  FirstFreeDrive = Chr$(DriveNum + 65) + ":"
  FreeDrive = FirstFreeDrive

End Function
'---------------------------------------------------------------
' GetSerialNumber()
' Retorna o número serial do disco
'---------------------------------------------------------------
Function GetSerialNumber(sRoot As String) As Long
  Dim lSerialNum As Long
  Dim r As Long
  Dim sTemp1 As String
  Dim sTemp2 As String

  sTemp1 = String$(255, Chr$(0))
  sTemp2 = String$(255, Chr$(0))
  r = GetVolumeInformation(sRoot, sTemp1, Len(sTemp1), lSerialNum, 0, 0, sTemp2, Len(sTemp2))
   
  GetSerialNumber = lSerialNum

End Function
'---------------------------------------------------------------
' NetworkUserID()
' Retorna o nome do usuário logado na rede.
'---------------------------------------------------------------
Public Function NetworkUserID() As String
  Dim lngBufferSize As Long
  Dim strUser As String
  
  Dim glngReturnStatus As Long
    
  On Error GoTo NetworkUserID_EH

  NetworkUserID = "Usuário desconhecido."
    
  lngBufferSize = 255
  strUser = Space$(lngBufferSize)

  glngReturnStatus = GetUserName(strUser, lngBufferSize)
  If glngReturnStatus = SUCCESS Then
    strUser = Left$(strUser, lngBufferSize - 1)
  Else
    Err = glngReturnStatus
  End If
  NetworkUserID = strUser
  Exit Function

NetworkUserID_EH:
  NetworkUserID = "ErrorInCall"
  Exit Function
  
End Function
'---------------------------------------------------------------
' OSPlatform()
' Retorna o sistema operacional que está sendo executado.
'---------------------------------------------------------------
Public Function OSPlatform() As String
  VersaoInfo.dwOSVersionInfoSize = Len(VersaoInfo)
  GetVersionEx VersaoInfo

  Select Case VersaoInfo.dwPlatformId
    Case VER_PLATFORM_WIN32s
      OSPlatform = "Windows 3.x" 'OS_WIN31
    Case VER_PLATFORM_WIN32_WINDOWS
      OSPlatform = "Windows 95"  'OS_WIN95
    Case VER_PLATFORM_WIN32_NT
      OSPlatform = "Windows NT"  'OS_WINNT
    Case Else
      OSPlatform = "Sistema desconhecido" 'OS_UNKNOWN
  End Select

End Function
'---------------------------------------------------------------
' OSVersion()
' Retorna a versão do sistema operacional que está sendo executado.
'---------------------------------------------------------------
Public Function OSVersion()
  VersaoInfo.dwOSVersionInfoSize = Len(VersaoInfo)
  GetVersionEx VersaoInfo
  
  OSVersion = VersaoInfo.dwMajorVersion & "." & _
    Format$((VersaoInfo.dwMinorVersion / 10), "0#") & "." & _
    Format$(VersaoInfo.dwBuildNumber Mod 65536, "0##")

End Function
Public Function TipoProcessador() As String
  Dim si As SystemInfo
  Dim cRetorno As String
  
  Call GetSystemInfo(si)
  
  Select Case (si.dwProcessorType)
    Case PROCESSOR_INTEL_386
      cRetorno = "386"
    Case PROCESSOR_INTEL_486
      cRetorno = "486"
    Case PROCESSOR_INTEL_PENTIUM
      cRetorno = "Pentium"
    Case PROCESSOR_MIPS_R2000
      cRetorno = "MIPS R2000"
    Case PROCESSOR_MIPS_R3000
      cRetorno = "MIPS R3000"
    Case PROCESSOR_MIPS_R4000
      cRetorno = "MIPS R4000"
    Case PROCESSOR_ALPHA_21064
      cRetorno = "DEC Alpha 21064"
  End Select
  
  TipoProcessador = cRetorno

End Function

Private Sub Class_Initialize()
   Set mvarMe = New frmSobre
End Sub

Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Load()
   Dim ClsApp As New DS_LOAD
   Dim Txt As String
   Dim ArqImg As String
   
   On Error Resume Next
   With mvarMe
      
      Screen.MousePointer = vbHourglass
      
      Dim SaveMemory As MemoryStatus
      
      ArqImg = mvarAplic.Path & "\Imagem\LogoEmpr.bmp"
      If ClsDos.FileExists(ArqImg) Then
         'mvarMe.PctEmpresa.Visible = False
         mvarMe.ImgEmpresa.Visible = True
         mvarMe.ImgEmpresa.Picture = LoadPicture(ArqImg)
      Else
         ArqImg = mvarAplic.Path & "\Imagem\LogoEmpr.jpg"
         If ClsDos.FileExists(ArqImg) Then
            'mvarMe.PctEmpresa.Visible = False
            mvarMe.ImgEmpresa.Visible = True
            mvarMe.ImgEmpresa.Picture = LoadPicture(ArqImg)
         End If
      End If
      'CenterForm Me
      
      'Me.Picture = LoadResPicture("FUNDO", vbResBitmap)
      
      'Zera controles de Time Out
      'Set glbForm = Me
      'glbLimiteTempo = 0
      'MDITecaplus.Tim_Time_Out.Enabled = False
      'MDITecaplus.Tim_Time_Out.Enabled = True
      Call ClsCtrl.LoadPctMouse("PRESS", vbResCursor, .LbleMail)
      .Caption = "Sobre o programa..."
      
      Call GlobalMemoryStatus(MemStat)
      With MemStat
         If .dwMemoryLoad <> SaveMemory.dwMemoryLoad Then
            MemoryLoad = .dwMemoryLoad
            TotalPhys = .dwTotalPhys
            AvailPhys = .dwAvailPhys
            TotalPageFile = .dwTotalPageFile
            AvailPageFile = .dwAvailPageFile
            TotalVirtual = .dwTotalVirtual
            AvailVirtual = .dwAvailVirtual
         End If
      End With
      SaveMemory = MemStat
      .LblAppName.Caption = ClsApp.GetAppName(mvarAplic)
      If Len(ClsApp.GetAppTitle(mvarAplic)) >= 20 Then
         .lbl(0).Font.SIZE = 16
      End If
      .lbl(0).Caption = ClsApp.GetAppTitle(mvarAplic)
      
      Txt = "Este sistema foi desenvolvido para a "
      Txt = Txt & UCase(ClsApp.GetAppLegalCopyright(mvarAplic))
      Txt = Txt & " por " & ClsApp.GetAppEmpresa(mvarAplic)
      .lbl(1).Caption = Txt
      
      Txt = "(c) " & UCase(ClsApp.GetAppLegalCopyright(mvarAplic))
      Txt = Txt & " - Todos os Direitos reservados."
      .lbl(2).Caption = Txt
      
      Txt = "Programa : " & ClsApp.GetAppName(mvarAplic) & vbCrLf
      Txt = Txt + "Descrição : " & ClsApp.GetAppTitle(mvarAplic) & vbCrLf
      Txt = Txt + ClsApp.GetAppVersao(mvarAplic) & vbCrLf
      Txt = Txt + "Compilado em " & Format$(FileDateTime(mvarAplic.Path & "\" & mvarAplic.ExeName & ".EXE"), "Short Date") & vbCrLf
      Txt = Txt + "Tamanho do executável : " & Format$((FileLen(mvarAplic.Path & "\" & mvarAplic.ExeName & ".EXE")), "#,##0") & " bytes"
      
      .Label2.Caption = Txt
      .Label3.Caption = "Sistema operacional: " & OSPlatform() & vbCrLf & _
      "Versão: " & OSVersion()
      
      '  Label4.Caption = "Percentual de memória em uso:"
      .Label5.Caption = Format$(MemoryLoad / 100, "###" & "%")
      '  Label6.Caption = "Memória física total:"
      .Label7.Caption = Format$(TotalPhys, "#,##0") & " bytes"
      '  Label8.Caption = "Memória física disponível:"
      .Label9.Caption = Format$(AvailPhys, "#,##0") & " bytes"
      '  LblAppName0.Caption = "Arquivo de paginação total:"
      .Label11.Caption = Format$(TotalPageFile, "#,##0") & " bytes"
      '  LblAppName12.Caption = "Arquivo de paginação disponível:"
      .Label13.Caption = Format$(AvailPageFile, "#,##0") & " bytes"
      '  LblAppName14.Caption = "Memória virtual total:"
      .Label15.Caption = Format$(TotalVirtual, "#,##0") & " bytes"
      '  LblAppName16.Caption = "Memória virtual disponível:"
      .Label17.Caption = Format$(AvailVirtual, "#,##0") & " bytes"
      
      '  Label18.Caption = "Usuário:"
      .Label19.Caption = NetworkUserID()
      '  Label20.Caption = "Nome do computador:"
      .Label21.Caption = ComputerName
      '  Label22.Caption = "Grupo de trabalho:"
      .Label23.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\VxD\VNETSUP", _
      "Workgroup", "Não fornecido.")
      '  Label24.Caption = "Rede primária:"
      
      '  Label26.Caption = "Rede secundária:"
      
      Dim cRede As String
      cRede = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\NetClient\0000", _
      "DriverDesc", "")
      
      If cRede = "" Then
         .Label25.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\NetClient\0001", _
         "DriverDesc", "Não fornecido.")
         .Label27.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\NetClient\0002", _
         "DriverDesc", "Não fornecido.")
      Else
         .Label25.Caption = cRede
         .Label27.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\NetClient\0001", _
         "DriverDesc", "Não fornecido.")
      End If
      
      '  Label28.Caption = "Caminho:"
      .Label29.Caption = globalCaminho
      
      '  label30.Caption = "Fabricante do processador:"
      .Label31.Caption = FabProcessador()
      '  Label32.Caption = "Tipo do processador:"
      .Label33.Caption = TipoProcessador()
      '  label34.Caption = "Resolução de vídeo:"
      .Label35.Caption = GetSystemMetrics(SM_CXSCREEN) & " x " & GetSystemMetrics(SM_CYSCREEN) & " pixels"
      '  Label36.Caption = "Driver de vídeo:"
      .Label37.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Display\0000", _
      "DriverDesc", "Não fornecido.")
      '  Label38.Caption = "Adaptador primário:"
      
      '  Label40.Caption = "Adaptador secundário:"
      
      Dim cAdaptador As String
      cAdaptador = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Net\0000", _
      "DriverDesc", "")
      
      If cAdaptador = "" Then
         .Label39.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Net\0001", _
         "DriverDesc", "Não fornecido.")
         .Label41.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Net\0002", _
         "DriverDesc", "Não fornecido.")
      Else
         .Label39.Caption = cAdaptador
         .Label41.Caption = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\Class\Net\0001", _
         "DriverDesc", "Não fornecido.")
      End If
      
      'Drive Information
      Dim lResult As Long
      sDrives = String$(128, 0) 'four characters * 26 + extra Null = 105
      'Cool! returns a string: A:\(NULL)B:\(NULL)C:\(NULL)...
      lResult = GetLogicalDriveStrings(1024, sDrives)
      Do While Left$(sDrives, 1) <> Chr$(0)
         'pull one off the front
         sDriveID = UCase$(Left$(sDrives, 3))
         sDrives = Mid$(sDrives, 5)
         'what kind is it?
         lDrive = ClsAutoIns.GetDriveType(sDriveID)
         Select Case lDrive
            Case DRIVE_ERROR
               .CmbDrv.AddItem sDriveID & "     [erro]"
            Case DRIVE_NOTPRESENT
               .CmbDrv.AddItem sDriveID & "     [não encontrado]"
            Case DRIVE_REMOVABLE
               .CmbDrv.AddItem sDriveID & "     [removível]"
            Case DRIVE_FIXED
               .CmbDrv.AddItem sDriveID & "     [disco rígido]"
            Case DRIVE_REMOTE
               'get its UNC name
               sBuffer = String$(255, 0)
               'don't forget to drop the "\" from the ID for this call
               lResult = WNetGetConnection(Left$(sDriveID, 2), sBuffer, Len(sBuffer))
               If lResult = 0 Then
                  .CmbDrv.AddItem sDriveID & "     " & sBuffer
               End If
            Case DRIVE_CDROM
               .CmbDrv.AddItem sDriveID & "     [CD-ROM]"
            Case DRIVE_RAMDISK
               .CmbDrv.AddItem sDriveID & "     [disco RAM]"
         End Select
         'Select the drive containing the Windows directory
         sBuffer = String$(255, 0)
         lResult = GetWindowsDirectory(sBuffer, Len(sBuffer))
         If Left$(.CmbDrv.List(.CmbDrv.ListCount - 1), 2) = Left$(sBuffer, 2) Then
            .CmbDrv.ListIndex = .CmbDrv.ListCount - 1 ' will activate CmbDrv_Click
         End If
      Loop
      
      .Label42.Caption = "O próximo disco disponível é " & ClsSobre.FreeDrive
      
      Screen.MousePointer = vbDefault
   End With
End Sub

