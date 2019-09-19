VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendMail2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Option Explicit
''********************************************
''*          SSL code by GioRock 2009        *
''********************************************
''*         Assembled by GioRock 2009        *
''*              giorock@libero.it           *
''********************************************
''*  Thanks to two Authors having published  *
''*    SSL https HTML connections VB code    *
''*              Jason K. Resch              *
''*                 Anonimous                *
''********************************************
'
''ReadMe:
''       This program is created and tested on HOTMAIL server where configuration
''       is about so:
''       SMTP: smtp.live.com
''       PORT: 587
''       ACCESS: USERNAME + PASSWORD
''       AUTHENTICATE: Need authentication of Server
''       SSL: Need a protected connection
''       For other Server I'don't know but HOTMAIL working fine and faster.
''-----------------------------------------------------------------------------
''       Do not use this program to SPAM messages or other abuse sending eMail
''-----------------------------------------------------------------------------
'
''PS: Not a Visual Basic example on the World I have found
''    about SSL Mailer, so after 3 days of full immersion work
''    this is a functional program.
''    Only 2 sample with SSL https HTML connections (out of date) and
''    not fully explained and good working.
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''    --This is my first time writing a program to send eMail--    ^
''^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'
''SSL is based on public key cryptography, it works in the following manner:
''
''client-hello         C -> S: challenge, cipher_specs
''server-hello         S -> C: connection-id,server_certificate,cipher_specs
''client-master-key    C -> S: {master_key}server_public_key
''client-finish        C -> S: {connection-id}client_write_key
''server-verify        S -> C: {challenge}server_write_key
''server-finish        S -> C: {new_session_id}server_write_key
''
''First the Client sends some random data known as the CHALLENGE, along with a list of ciphers it can use, for simplicity we will only use 128-bit RC4 with MD5
''The Server responds with a random data, known as the CONNECTION-ID, and the Server's Certificate and list of cipher specs
''The Client extracts the Public Key from the Server's Certificate then uses it to Encrypt a randomly generated Master Key, this Key then sent to the Server
''The Client and Server both generate 2 keys each by hashing the Master Key with other values, and the client sends a finish message, encrypted with the client write key
''The Server Responds by returning the CHALLENGE encrypted using the Client Read Key, this proves to the Clinet that the Server is who it says its is
''The Server sends its finish message, which consists of a randomly generated value, this value can be used to re-create the session in a new connection, but that is not supported in this example
'
''You can take a look on Wikipedia for an exaustive explanation about SSL connections
'
''TODO:
''re-create the session in a new connection
''trap all error
''add POP3 connection to receive messages
''simplify and ameliorate all code and routines
''and much more....
'
''YOU CAN SEND ME SUGGESTIONS AND CRITICS TO: giorock@libero.it
''OR DIRECTLY ON PSC PAGE
'
''CryptoAPI Functions
'Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
'Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
'Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
'Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
'Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
'Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
'Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
'Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
'Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
'Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
'
''CryptoAPI Constants
'Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
'Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar
'Private Const PROV_RSA_FULL As Long = 1
'Private Const CRYPT_NEWKEYSET As Long = 8
'Private Const CRYPT_EXPORTABLE As Long = 1
'Private Const ALG_CLASS_HASH = (4 * 2 ^ 13)
'Private Const ALG_SID_MD5 = 3
'Private Const ALG_TYPE_ANY = 0
'Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
'Private Const ALG_CLASS_DATA_ENCRYPT = (3 * 2 ^ 13)
'Private Const ALG_SID_RC4 = 1
'Private Const ALG_TYPE_STREAM = (4 * 2 ^ 9)
'Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
'Private Const HP_HASHVAL As Long = 2
'Private Const SIMPLEBLOB As Long = 1
'Private Const GEN_KEY_BITS As Long = &H800000
'
''Class Variables
'Private hCryptProv As Long
'Private hClientWriteKey As Long
'Private hClientReadKey As Long
'Private hMasterKey As Long
'Private lngType As Long
'
''Variables for Parsing
'Private Layer As Integer
'Private InBuffer As String
'Private Processing As Boolean
'Private SeekLen As Integer
'
''Encryption Keys
'Private MASTER_KEY As String
'Private CLIENT_READ_KEY As String
'Private CLIENT_WRITE_KEY As String
'
''Server Attributes
'Private Private_KEY As String
'Private ENCODED_CERT As String
'Private CONNECTION_ID As String
'
''Counters
'Private SEND_SEQUENCE_NUMBER As Double
'Private RECV_SEQUENCE_NUMBER As Double
'
''Hand Shake Variables
'Private CLIENT_HELLO As String
'Private CHALLENGE_DATA As String
'Private PUBLIC_KEY As String
'
'Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'
'
'''SENDING VARIABLES
'Private gstrServerName As String
'Private glngPort As Long
'Private gblnSSL As Boolean
'Private gstrFromAddress As String
'Private gstrDomainName As String
'Private gstrUserID As String
'Private gstrUserPass As String
'Private gintSent As Integer
'
'''MESSAGE VARIABLES
'Private gstrSubject As String
'Private gstrFromName As String
'Private gstrMessage As String
'Private gstrToAddress() As String
'Private gblnDone As Boolean
'Private gblnBCCMode As Boolean
'Private gstrBCC() As String
'Private gstrCC() As String
'Private gblnBCC As Boolean
'Private gstrFileNames() As String
'Private gblnAttachments As Boolean
'Private gstrData() As String
'Private gstrMessageFileName As String
'Private gstrAllAddresses() As String
'''OTHER VARIABLES
'Private gblnConnected As Boolean
'Private ProgBar As ProgressBar
'Private gintTransfer As Integer
'Private Enum EState
'    Connect
'    helo
'    MailFrom
'    sendto
'    Data
'    MessageData
'    EndMessage
'End Enum
'
'Private Type FileInfo
'    Filename As String
'    Code As String
'    Send As Boolean
'End Type
'Private State As EState
'
'Private gintsentTo As Integer
'Private gintsentCC As Integer
'Private pintSend As Integer
'Private pintSent As Integer
'
'Private Const MAX_LINELENGTH As Long = 76 ' Must be a multiple of 4
'Private Const CHAR_EQUAL As Byte = 61
'Private Const CHAR_CR As Byte = 13
'Private Const CHAR_LF As Byte = 10
'
'
'Private m_Index1(0 To 255) As Byte
'Private m_Index2(0 To 255) As Byte
'Private m_Index3(0 To 255) As Byte
'Private m_Index4(0 To 63) As Byte
'Private m_ReverseIndex1(0 To 255) As Byte
'Private m_ReverseIndex2(0 To 255, 0 To 1) As Byte
'Private m_ReverseIndex3(0 To 255, 0 To 1) As Byte
'Private m_ReverseIndex4(0 To 255) As Byte
'
'Private FileCode() As FileInfo
'Private Const conBoundary = "SendNextMIME_121_32"
'
'Private Enum DecodeType
'    base64
'    uu
'End Enum
'
'Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (dst As Any, src As Any, ByVal cb As Long)
'
'Public Event DecodeProgress(ByVal Percent As Single, ByVal Total As Long)
'Public Event EncodeProgress(ByVal Percent As Single, ByVal Total As Long)
'
'Private Const PortionSize As Integer = 45
'Private Const BufLen As Integer = 1024
'
'Private DestFileName As String
'Private mstrOutFile
'
'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
'() '    (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, _
'    ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
'
'Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
'() '    (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, _
'    ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'
'Private Declare Function InternetCloseHandle Lib "wininet.dll" _
'() '    (ByVal hInet As Long) As Long
'
'Private Declare Function InternetDial Lib "wininet.dll" Alias "InternetOpenUrlA" _
'() '    (ByVal hwndParent As Long, ByVal strConnection As String, ByVal dwFlags As Long, _
'    ByRef dwConnection As Long, ByVal dwReserved As Long) As Long
'
'Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
'Private Const INTERNET_FLAG_RELOAD = &H80000000
'Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
'Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
'Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000
'Private Const INTERNET_DIAL_SHOW_OFFLINE = &H4000
'Private Const INTERNET_DIAL_UNATTENDED = &H8000
'
'Private WithEvents SocketSSL As Winsock
'Private RTBX As RichTextBox
'
'Private Function GetDomainName(ByVal pstrFromAddress As String) As String
'    Dim pintInstr As Integer
'
'    pintInstr = InStr(pstrFromAddress, "@")
'    GetDomainName = Mid(pstrFromAddress, pintInstr + 1)
'End Function
'
'Private Function CheckAddressFrom(pstrAddy) As String
'    Dim pintC As Integer
'
'    If Left(pstrAddy, 1) <> "<" Then
'        pstrAddy = "<" & pstrAddy
'    End If
'
'    If Right(pstrAddy, 1) <> ">" Then
'        pstrAddy = pstrAddy & ">"
'    End If
'
'    CheckAddressFrom = pstrAddy
'End Function
'
'Private Function GetMIME(MIMEBoundary As String) As String
'    Dim strMIMEMessageContent As String
'    Dim strMIMEVersion As String
'    Dim strMIMEContent As String
'
'
'
'    strMIMEContent = "Content-Type: multipart/mixed; " _
'                     & "boundary=" & """" & MIMEBoundary & """"
'
'    strMIMEMessageContent = "--" & MIMEBoundary & vbCrLf & _
'                            "Content-Type: text/plain; " & "charset=" & """" & "iso-8859-1" & """" & vbCrLf & _
'                            "Content-Transfer-Encoding: 8bit"
'    GetMIME = strMIMEContent & vbCrLf & vbCrLf & vbCrLf & strMIMEMessageContent
'
'End Function
'
'
'Private Function ProcessHeader(sTo As String, sFrom As String, sCC As String, sSubject As String) As String
'
'    Dim strMIMEMessage          As String
'    Dim strMIMEContent          As String
'    Dim strMIMEMessageContent   As String
'    Dim strMIMEClient           As String
'    Dim strMIMEMailInfo         As String
'    Dim strMIMEVersion          As String
'
'    strMIMEMailInfo = "DATE: " & Format(Now, "dd mmm yy ttttt") & vbCrLf & _
'                      "FROM: " & Trim$(sFrom) & vbCrLf & _
'                      "TO: " & Trim$(sTo) & vbCrLf & _
'                      "CC: " & Trim$(sCC) & vbCrLf & _
'                      "SUBJECT: " & sSubject
'
'    strMIMEClient = "X-Mailer: " & App.ProductName & vbCrLf & _
'                    "X-Version: " & App.Major & "." & App.Minor & vbCrLf & _
'                    "X-CompanyName: " & App.CompanyName
'
'    strMIMEVersion = "MIME-Version: 1.0"
'
'    ProcessHeader = strMIMEMailInfo & vbCrLf & _
'                    strMIMEVersion & vbCrLf & _
'                    strMIMEClient
'
'
'
'End Function
'
'
'Private Function GetTo(pstrGroup() As String) As String
'    Dim pintC
'
'    For pintC = 0 To UBound(pstrGroup)
'        GetTo = GetTo & ", " & pstrGroup(pintC)
'    Next pintC
'
'    GetTo = Mid(GetTo, 3)
'End Function
'
'
'Private Sub CreateDocToSend()
'    Dim pstrHeader As String
'    Dim pstrTMPFile As String
'    Dim pintFF As Integer
'    Dim pstrMIMEInfo As String
'    Dim pstrBuffer As String
'    Dim pstrTo As String
'    Dim pstrCC As String
'    Dim pstrFrom As String
'    Dim pstrTotal As String
'
'    pstrTo = GetTo(gstrToAddress)
'    pstrCC = GetTo(gstrCC)
'    If pstrTo = "" Then
'        pstrTo = "Undisclosed"
'    Else
'        pstrTo = Replace(pstrTo, "<", "")
'        pstrTo = Replace(pstrTo, ">", "")
'    End If
'
'    If pstrCC <> "" Then
'        pstrCC = Replace(pstrCC, "<", "")
'        pstrCC = Replace(pstrCC, ">", "")
'    End If
'
'    pstrFrom = gstrFromAddress
'    pstrHeader = ProcessHeader(pstrTo, Chr(34) & gstrFromName & Chr(34) & " " & pstrFrom, pstrCC, gstrSubject)
'    pstrTMPFile = GetTmpFile
'    pintFF = FreeFile
'    pstrHeader = pstrHeader & vbCrLf
'
'    If gblnAttachments = True Then
'        pstrMIMEInfo = GetMIME(conBoundary)
'        pstrHeader = pstrHeader & pstrMIMEInfo & vbCrLf
'    End If
'
'    ''WRITE
'    Open pstrTMPFile For Output As #pintFF
'        Print #pintFF, pstrHeader
'        Print #pintFF, gstrMessage
'        If gblnAttachments = True Then
'            Dim pintC As Integer
'            For pintC = 0 To UBound(FileCode)
'                If FileCode(pintC).Send = True Then
'                    Print #pintFF, FileCode(pintC).Code
'                    If pintC = UBound(FileCode) Then
'                        Print #pintFF, "--" & conBoundary & "--"
'                    End If
'                End If
'            Next pintC
'        End If
'
'    Close #pintFF
'    gstrMessageFileName = pstrTMPFile
'End Sub
'
'
'Private Sub CheckAddress()
'    Dim pintC As Integer
'
'    On Error Resume Next
'
'    For pintC = 0 To UBound(gstrToAddress)
'        If Left(gstrToAddress(pintC), 1) <> "<" Then
'            gstrToAddress(pintC) = "<" & gstrToAddress(pintC)
'        End If
'
'        If Right(gstrToAddress(pintC), 1) <> ">" Then
'            gstrToAddress(pintC) = gstrToAddress(pintC) & ">"
'        End If
'    Next pintC
'
'    For pintC = 0 To UBound(gstrCC)
'        If Left(gstrCC(pintC), 1) <> "<" Then
'            gstrCC(pintC) = "<" & gstrCC(pintC)
'        End If
'
'        If Right(gstrCC(pintC), 1) <> ">" Then
'            gstrCC(pintC) = gstrCC(pintC) & ">"
'        End If
'    Next pintC
'
'    For pintC = 0 To UBound(gstrBCC)
'        If Left(gstrBCC(pintC), 1) <> "<" Then
'            gstrBCC(pintC) = "<" & gstrBCC(pintC)
'        End If
'
'        If Right(gstrBCC(pintC), 1) <> ">" Then
'            gstrBCC(pintC) = gstrBCC(pintC) & ">"
'        End If
'    Next pintC
'
'End Sub
'
'
'Private Function CheckConnectionPing() As Boolean
'   Dim sTmp As String
'   Dim hInet As Long
'   Dim hUrl As Long
'   Dim Flags As Long
'   Dim url As Variant
'   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
'   If hInet Then
'      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
'      hUrl = InternetOpenUrl(hInet, "http://www.yahoo.com", vbNullString, 0, Flags, 0)
'      If hUrl Then
'         CheckConnectionPing = True
'         Call InternetCloseHandle(hUrl)
'    Else
'         hUrl = InternetOpenUrl(hInet, "http://www.google.com", vbNullString, 0, Flags, 0)
'         If hUrl Then
'            CheckConnectionPing = True
'            Call InternetCloseHandle(hUrl)
'         End If
'      End If
'   End If
'   Call InternetCloseHandle(hInet)
'End Function
'
'
'Private Function StartConnection() As Long
'    Dim pintRes As Integer
'    pintRes = InternetDial(0, "", INTERNET_DIAL_FORCE_PROMPT, 0, 0)
'    StartConnection = pintRes
'End Function
'Public Function ConnectToInternet() As Long
'    ConnectToInternet = StartConnection
'End Function
'
'
'Public Function SendEmail(FromName As String, Subject As String, Message As String, arrToAddresses() As String, arrBCCAddresses() As String, arrCCAddresses() As String, WSock As Winsock, RichText As RichTextBox, Optional ProgsBar As ProgressBar) As Integer
'    Dim pintC As Integer
'    Dim pintRec As Integer
'    Dim pblnConnected As Boolean
'
'    pblnConnected = CheckConnectionPing
'    If pblnConnected = False Then
'        SendEmail = 1
'        Exit Function
'    End If
'
'    Set SocketSSL = WSock
'    Set RTBX = RichText
'
'    gstrFromName = FromName
'    gstrSubject = Subject
'    gstrMessage = Message
'    gstrToAddress = arrToAddresses
'    gstrCC = arrCCAddresses
'    gstrBCC = arrBCCAddresses
'    gblnDone = False
'    gblnConnected = False
'    gintTransfer = 0
'    gintSent = 0
'
'    Call CheckAddress
'
'    For pintC = 0 To UBound(gstrCC)
'        If gstrCC(pintC) = "" Or gstrCC(pintC) = "<>" Then Exit For
'        ReDim Preserve gstrAllAddresses(pintRec)
'        gstrAllAddresses(pintRec) = gstrCC(pintC)
'        pintRec = pintRec + 1
'    Next pintC
'
'    For pintC = 0 To UBound(gstrToAddress)
'        If gstrToAddress(pintC) = "" Or gstrToAddress(pintC) = "<>" Then Exit For
'        ReDim Preserve gstrAllAddresses(pintRec)
'        gstrAllAddresses(pintRec) = gstrToAddress(pintC)
'        pintRec = pintRec + 1
'    Next pintC
'
'    On Error GoTo AllAddressError
'    If gstrAllAddresses(0) = "" Then
'        ReDim gstrAllAddresses(0)
'    End If
'
'    On Error Resume Next
'    Set ProgBar = ProgsBar
'
'    Call CreateDocToSend
'
'    If gblnSSL = True Then
'        Call ConnectSSLSocket
'    Else
''        Here you can add a not SSL eMail procedure
''        Dim Soc As New RegSocket
''        Call Soc.ConnectRegSocket
'    End If
'
'    Kill gstrMessageFileName
'    Call Reset
'
'    Set ProgBar = Nothing
'
'    SendEmail = 0
'    Exit Function
'
'AllAddressError:
'    If Err.Number = 9 Then
'        ReDim gstrAllAddresses(0)
'        Resume Next
'    End If
'
'End Function
'
'
'Public Sub SetUp(ServerName As String, Port As Long, FromAddress As String, SSL As Boolean, UserID As String, UserPass As String)
'
'    gstrServerName = ServerName
'    glngPort = Port
'    gblnSSL = SSL
'    gstrFromAddress = CheckAddressFrom(FromAddress)
'    gstrDomainName = GetDomainName(gstrFromAddress)
'    gstrUserID = UserID
'    gstrUserPass = UserPass
'
'End Sub
'Private Function SelectExt(ByVal vsFullPathname As String) As String
'
'    Dim extension As String
'    extension = StrReverse(LCase$(Left(StrReverse(vsFullPathname), InStr(1, StrReverse(vsFullPathname), "."))))
'    If Len(extension) = 0 Then extension = vsFullPathname
'
'    Select Case extension
'
'        Case ".323"
'            SelectExt = "text/h323"
'
'        Case ".aab"
'            SelectExt = "application/x-authorware-bin"
'
'        Case ".aam"
'            SelectExt = "application/x-authorware-map"
'
'        Case ".ace"
'            SelectExt = "application/x-compressed"
'
'        Case ".acp"
'            SelectExt = "audio/x-mei-aac"
'
'        Case ".ai"
'            SelectExt = "application/postscript"
'
'        Case ".aif"
'            SelectExt = "audio/aiff"
'
'        Case ".aifc"
'            SelectExt = "audio/aiff"
'
'        Case ".aiff"
'            SelectExt = "audio/aiff"
'
'        Case ".aip"
'            SelectExt = "text/x-audiosoft-intra"
'
'        Case ".ARJ"
'            SelectExt = "application/x-compressed"
'
'        Case ".art"
'            SelectExt = "image/x-jg"
'
'        Case ".asa"
'            SelectExt = "text/asa"
'
'        Case ".asf"
'            SelectExt = "video/x-ms-asf"
'
'        Case ".asp"
'            SelectExt = "text/asp"
'
'        Case ".asx"
'            SelectExt = "video/x-ms-asf"
'
'        Case ".asx"
'            SelectExt = "video/x-ms-asx"
'
'        Case ".au"
'            SelectExt = "audio/basic"
'
'        Case ".aut"
'            SelectExt = "application/pbautomation"
'
'        Case ".avi"
'            SelectExt = "video/avi"
'
'        Case ".avi"
'            SelectExt = "video/x-msvideo"
'
'        Case ".bmo"
'            SelectExt = "audio/blue-matter-offer"
'
'        Case ".bmp"
'            SelectExt = "image/bmp"
'
'        Case ".bmp"
'            SelectExt = "image/x-bmp"
'
'        Case ".bmr"
'            SelectExt = "text/blue-matter-content-ref"
'
'        Case ".bmt"
'            SelectExt = "audio/blue-matter-song"
'
'        Case ".bub"
'            SelectExt = "application/photobubble"
'
'        Case ".cat"
'            SelectExt = "application/vnd.ms-pki.seccat"
'
'        Case ".cdf"
'            SelectExt = "application/x-cdf"
'
'        Case ".cel"
'            SelectExt = "video/flc"
'
'        Case ".cer"
'            SelectExt = "application/pkix-cert"
'
'        Case ".cer"
'            SelectExt = "application/x-x509-ca-cert"
'
'        Case ".class"
'            SelectExt = "java/*"
'
'        Case ".crl"
'            SelectExt = "application/pkix-crl"
'
'        Case ".crt"
'            SelectExt = "application/pkix-cert"
'
'        Case ".crt"
'            SelectExt = "application/x-x509-ca-cert"
'
'        Case ".css"
'            SelectExt = "text/css"
'
'        Case ".dcr"
'            SelectExt = "application/x-director"
'
'        Case ".der"
'            SelectExt = "application/pkix-cert"
'
'        Case ".der"
'            SelectExt = "application/x-x509-ca-cert"
'
'        Case ".dib"
'            SelectExt = "image/bmp"
'
'        Case ".dib"
'            SelectExt = "image/x-bmp"
'
'        Case ".dif"
'            SelectExt = "video/x-dv"
'
'        Case ".dir"
'            SelectExt = "application/x-director"
'
'        Case ".dll"
'            SelectExt = "application/x-msdownload"
'
'        Case ".doc"
'            SelectExt = "application/msword"
'
'        Case ".dot"
'            SelectExt = "application/msword"
'
'        Case ".dpg"
'            SelectExt = "application/vnd.dpgraph"
'
'        Case ".dpgraph"
'            SelectExt = "application/vnd.dpgraph"
'
'        Case ".dv"
'            SelectExt = "video/x-dv"
'
'        Case ".dxr"
'            SelectExt = "application/x-director"
'
'        Case ".eml"
'            SelectExt = "message/rfc822"
'
'        Case ".emm"
'            SelectExt = "application/x-emms-content"
'
'        Case ".eps"
'            SelectExt = "application/postscript"
'
'        Case ".exe"
'            SelectExt = "application/x-msdownload"
'
'        Case ".fdf"
'            SelectExt = "application/vnd.fdf"
'
'        Case ".fif"
'            SelectExt = "application/fractals"
'
'        Case ".flc"
'            SelectExt = "video/flc"
'
'        Case ".fli"
'            SelectExt = "video/flc"
'
'        Case ".fml"
'            SelectExt = "application/file-mirror-list"
'
'        Case ".fpx"
'            SelectExt = "image/x-xbitmap"
'
'        Case ".gif"
'            SelectExt = "image/gif"
'
'        Case ".grv"
'            SelectExt = "application/vnd.groove-injector"
'
'        Case ".gz"
'            SelectExt = "application/x-compressed"
'
'        Case ".gz"
'            SelectExt = "application/x-gzip"
'
'        Case ".hpf"
'            SelectExt = "application/x-icq-hpf"
'
'        Case ".hqx"
'            SelectExt = "application/mac-binhex40"
'
'        Case ".hta"
'            SelectExt = "application/hta"
'
'        Case ".htc"
'            SelectExt = "text/x-component"
'
'        Case ".htm"
'            SelectExt = "text/html"
'
'        Case ".html"
'            SelectExt = "text/html"
'
'        Case ".htt"
'            SelectExt = "text/webviewhtml"
'
'        Case ".htx"
'            SelectExt = "text/html"
'
'        Case ".ico"
'            SelectExt = "image/x-icon"
'
'        Case ".iii"
'            SelectExt = "application/x-iphone"
'
'        Case ".ins"
'            SelectExt = "application/x-internet-signup"
'
'        Case ".ips"
'            SelectExt = "application/x-ipscript"
'
'        Case ".ipx"
'            SelectExt = "application/x-ipix"
'
'        Case ".isp"
'            SelectExt = "application/x-internet-signup"
'
'        Case ".IVF"
'            SelectExt = "video/x-ivf"
'
'        Case ".ivr"
'            SelectExt = "i-world/i-vrml"
'
'        Case ".java"
'            SelectExt = "java/*"
'
'        Case ".java"
'            SelectExt = "text/java"
'
'        Case ".jfif"
'            SelectExt = "image/pjpeg"
'
'        Case ".jpe"
'            SelectExt = "image/jpeg"
'
'        Case ".jpeg"
'            SelectExt = "image/jpeg"
'
'        Case ".jpg"
'            SelectExt = "image/jpeg"
'
'        Case ".JS"
'            SelectExt = "application/x-javascript"
'
'        Case ".la1"
'            SelectExt = "audio/x-liquid-file"
'
'        Case ".lar"
'            SelectExt = "application/x-laplayer-reg"
'
'        Case ".latex"
'            SelectExt = "application/x-latex"
'
'        Case ".lav"
'            SelectExt = "audio/x-liquid"
'
'        Case ".lavs"
'            SelectExt = "audio/x-liquid-secure"
'
'        Case ".lha"
'            SelectExt = "application/x-compressed"
'
'        Case ".lks"
'            SelectExt = "application/x-lk-rlestream"
'
'        Case ".lmsff"
'            SelectExt = "audio/x-la-lms"
'
'        Case ".lqt"
'            SelectExt = "audio/x-liquid-file"
'
'        Case ".ls"
'            SelectExt = "application/x-javascript"
'
'        Case ".lsf"
'            SelectExt = "video/x-la-asf"
'
'        Case ".lsx"
'            SelectExt = "video/x-la-asf"
'
'        Case ".LZH"
'            SelectExt = "application/x-compressed"
'
'        Case ".m1v"
'            SelectExt = "video/mpeg"
'
'        Case ".m3u"
'            SelectExt = "audio/mpegurl"
'
'        Case ".m3u"
'            SelectExt = "audio/x-mpegurl"
'
'        Case ".mac"
'            SelectExt = "image/x-macpaint"
'
'        Case ".man"
'            SelectExt = "application/x-troff-man"
'
'        Case ".mbc"
'            SelectExt = "application/x-pn-virtualink"
'
'        Case ".mbo"
'            SelectExt = "application/x-previewsystems-vbox-music"
'
'        Case ".mbox"
'            SelectExt = "application/x-previewsystems-vbox-music"
'
'        Case ".mdb"
'            SelectExt = "application/msaccess"
'
'        Case ".med"
'            SelectExt = "application/x-att-a2bmusic-purchase"
'
'        Case ".mes"
'            SelectExt = "application/x-att-a2bmusic"
'
'        Case ".mht"
'            SelectExt = "message/rfc822"
'
'        Case ".mhtml"
'            SelectExt = "message/rfc822"
'
'        Case ".mid"
'            SelectExt = "audio/mid"
'
'        Case ".midi"
'            SelectExt = "audio/mid"
'
'        Case ".mix"
'            SelectExt = "image/x-xbitmap"
'
'        Case ".mjf"
'            SelectExt = "audio/x-vnd.AudioExplosion.MjuiceMediaFile"
'
'        Case ".mjv"
'            SelectExt = "audio/audio/mjuice_voucher"
'
'        Case ".mmjb_mime"
'            SelectExt = "application/x-musicmatch-mmjb5.20detect"
'
'        Case ".mmz"
'            SelectExt = "application/x-mmjb-mmz"
'
'        Case ".mocha"
'            SelectExt = "application/x-javascript"
'
'        Case ".mov"
'            SelectExt = "video/quicktime"
'
'        Case ".movie"
'            SelectExt = "video/x-sgi-movie"
'
'        Case ".mp1"
'            SelectExt = "audio/mpeg"
'
'        Case ".mp2"
'            SelectExt = "video/mpeg"
'
'        Case ".mp2v"
'            SelectExt = "video/mpeg"
'
'        Case ".mp3"
'            SelectExt = "audio/mpeg"
'
'        Case ".mpa"
'            SelectExt = "video/mpeg"
'
'        Case ".mpe"
'            SelectExt = "video/mpeg"
'
'        Case ".mpeg"
'            SelectExt = "video/mpeg"
'
'        Case ".mpg"
'            SelectExt = "video/mpeg"
'
'        Case ".mpga"
'            SelectExt = "audio/mpeg"
'
'        Case ".mpv"
'            SelectExt = "video/mpg"
'
'        Case ".mpv2"
'            SelectExt = "video/mpeg"
'
'        Case ".mwc"
'            SelectExt = "application/vnd.dpgraph"
'
'        Case ".mxp"
'            SelectExt = "application/x-mmxp"
'
'        Case ".npi"
'            SelectExt = "application/x-pn-npistream"
'
'        Case ".nws"
'            SelectExt = "message/rfc822"
'
'        Case ".p10"
'            SelectExt = "application/pkcs10"
'
'        Case ".p12"
'            SelectExt = "application/x-pkcs12"
'
'        Case ".p7b"
'            SelectExt = "application/x-pkcs7-certificates"
'
'        Case ".p7c"
'            SelectExt = "application/pkcs7-mime"
'
'        Case ".p7m"
'            SelectExt = "application/pkcs7-mime"
'
'        Case ".p7r"
'            SelectExt = "application/x-pkcs7-certreqresp"
'
'        Case ".p7s"
'            SelectExt = "application/pkcs7-signature"
'
'        Case ".pct"
'            SelectExt = "image/pict"
'
'        Case ".pdf"
'            SelectExt = "application/pdf"
'
'        Case ".pfx"
'            SelectExt = "application/x-pkcs12"
'
'        Case ".pic"
'            SelectExt = "image/pict"
'
'        Case ".pict"
'            SelectExt = "image/pict"
'
'        Case ".pko"
'            SelectExt = "application/vnd.ms-pki.pko"
'
'        Case ".pl"
'            SelectExt = "application/x-perl"
'
'        Case ".plg"
'            SelectExt = "text/html"
'
'        Case ".pls"
'            SelectExt = "audio/scpls"
'
'        Case ".pls"
'            SelectExt = "audio/x-scpls"
'
'        Case ".png"
'            SelectExt = "image/png"
'
'        Case ".pnq"
'            SelectExt = "application/x-icq-pnq"
'
'        Case ".pntg"
'            SelectExt = "image/x-macpaint"
'
'        Case ".POT"
'            SelectExt = "application/vnd.ms-powerpoint"
'
'        Case ".ppa"
'            SelectExt = "application/vnd.ms-powerpoint"
'
'        Case ".pps"
'            SelectExt = "application/vnd.ms-powerpoint"
'
'        Case ".ppt"
'            SelectExt = "application/x-mspowerpoint"
'
'        Case ".prf"
'            SelectExt = "application/pics-rules"
'
'        Case ".ps"
'            SelectExt = "application/postscript"
'
'        Case ".pwz"
'            SelectExt = "application/vnd.ms-powerpoint"
'
'        Case ".py"
'            SelectExt = "text/plain"
'
'        Case ".pyw"
'            SelectExt = "text/plain"
'
'        Case ".qht"
'            SelectExt = "text/x-html-insertion"
'
'        Case ".qhtm"
'            SelectExt = "text/x-html-insertion"
'
'        Case ".qt"
'            SelectExt = "video/quicktime"
'
'        Case ".qti"
'            SelectExt = "image/x-quicktime"
'
'        Case ".qtif"
'            SelectExt = "image/x-quicktime"
'
'        Case ".qtl"
'            SelectExt = "application/x-quicktimeplayer"
'
'        Case ".r00"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r01"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r02"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r03"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r04"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r05"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r06"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r07"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r08"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r09"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r10"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r11"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r12"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r13"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r14"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r15"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r16"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r17"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r18"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r19"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r20"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r21"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r22"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r23"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r24"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r25"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r26"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r27"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r28"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r29"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r30"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r31"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r32"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r33"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r34"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r35"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r36"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r37"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r38"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r39"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r3t"
'            SelectExt = "text/vnd.rn-realtext3d"
'
'        Case ".r40"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r41"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r42"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r43"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r44"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r45"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r46"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r47"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r48"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r49"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r50"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r51"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r52"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r53"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r54"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r55"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r56"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r57"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r58"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r59"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r60"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r61"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r62"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r63"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r64"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r65"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r66"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r67"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r68"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r69"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r70"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r71"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r72"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r73"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r74"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r75"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r76"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r77"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r78"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r79"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r80"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r81"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r82"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r83"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r84"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r85"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r86"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r87"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r88"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r89"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r90"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r91"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r92"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r93"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r94"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r95"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r96"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r97"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r98"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".r99"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".ra"
'            SelectExt = "audio/vnd.rn-realaudio"
'
'        Case ".ram"
'            SelectExt = "audio/x-pn-realaudio"
'
'        Case ".rar"
'            SelectExt = "application/x-rar-compressed"
'
'        Case ".rat"
'            SelectExt = "application/rat-file"
'
'        Case ".rf"
'            SelectExt = "image/vnd.rn-realflash"
'
'        Case ".rjs"
'            SelectExt = "application/vnd.rn-realsystem-rjs"
'
'        Case ".rjt"
'            SelectExt = "application/vnd.rn-realsystem-rjt"
'
'        Case ".rm"
'            SelectExt = "application/vnd.rn-realmedia"
'
'        Case ".rmi"
'            SelectExt = "audio/mid"
'
'        Case ".rmj"
'            SelectExt = "application/vnd.rn-realsystem-rmj"
'
'        Case ".rmm"
'            SelectExt = "audio/x-pn-realaudio"
'
'        Case ".rmp"
'            SelectExt = "application/vnd.rn-rn_music_package"
'
'        Case ".rmx"
'            SelectExt = "application/vnd.rn-realsystem-rmx"
'
'        Case ".rnx"
'            SelectExt = "application/vnd.rn-realplayer"
'
'        Case ".rob"
'            SelectExt = "application/vnd.rn-objects"
'
'        Case ".rp"
'            SelectExt = "image/vnd.rn-realpix"
'
'        Case ".rpm"
'            SelectExt = "audio/x-pn-realaudio-plugin"
'
'        Case ".rsml"
'            SelectExt = "application/vnd.rn-rsml"
'
'        Case ".rt"
'            SelectExt = "text/vnd.rn-realtext"
'
'        Case ".rtf"
'            SelectExt = "application/msword"
'
'        Case ".rtsp"
'            SelectExt = "application/x-rtsp"
'
'        Case ".rv"
'            SelectExt = "video/vnd.rn-realvideo"
'
'        Case ".sc"
'            SelectExt = "application/vnd.optx-screenwatch"
'
'        Case ".scm"
'            SelectExt = "application/x-icq-scm"
'
'        Case ".sct"
'            SelectExt = "text/scriptlet"
'
'        Case ".sd2"
'            SelectExt = "audio/x-sd2"
'
'        Case ".sdf"
'            SelectExt = "application/x-server-launch"
'
'        Case ".sdp"
'            SelectExt = "application/sdp"
'
'        Case ".sgi"
'            SelectExt = "image/x-sgi"
'
'        Case ".sit"
'            SelectExt = "application/x-stuffit"
'
'        Case ".sma"
'            SelectExt = "application/x-smb-directive"
'
'        Case ".smi"
'            SelectExt = "application/smil"
'
'        Case ".smil"
'            SelectExt = "application/smil"
'
'        Case ".snd"
'            SelectExt = "audio/basic"
'
'        Case ".spc"
'            SelectExt = "application/x-pkcs7-certificates"
'
'        Case ".spl"
'            SelectExt = "application/futuresplash"
'
'        Case ".spn"
'            SelectExt = "application/vnd.spinnerplus"
'
'        Case ".ssm"
'            SelectExt = "application/streamingmedia"
'
'        Case ".sst"
'            SelectExt = "application/vnd.ms-pki.certstore"
'
'        Case ".stl"
'            SelectExt = "application/vnd.ms-pki.stl"
'
'        Case ".stm"
'            SelectExt = "text/html"
'
'        Case ".svg"
'            SelectExt = "image/svg+xml"
'
'        Case ".svg"
'            SelectExt = "image/svg-xml"
'
'        Case ".svgz"
'            SelectExt = "image/svg+xml"
'
'        Case ".svgz"
'            SelectExt = "image/svg-xml"
'
'        Case ".swf"
'            SelectExt = "application/x-shockwave-flash"
'
'        Case ".tar"
'            SelectExt = "application/x-compressed"
'
'        Case ".tar"
'            SelectExt = "application/x-tar"
'
'        Case ".tga"
'            SelectExt = "image/x-targa"
'
'        Case ".tgz"
'            SelectExt = "application/x-compressed"
'
'        Case ".tif"
'            SelectExt = "image/tiff"
'
'        Case ".tiff"
'            SelectExt = "image/tiff"
'
'        Case ".txt"
'            SelectExt = "text/plain"
'
'        Case ".uin"
'            SelectExt = "application/x-icq"
'
'        Case ".uls"
'            SelectExt = "text/iuls"
'
'        Case ".ultact"
'            SelectExt = "application/x-UltimateAction"
'
'        Case ".ulw"
'            SelectExt = "audio/basic"
'
'        Case ".urls"
'            SelectExt = "application/x-url-list"
'
'        Case ".UU"
'            SelectExt = "application/x-compressed"
'
'        Case ".UUE"
'            SelectExt = "application/x-compressed"
'
'        Case ".vcf"
'            SelectExt = "text/x-vcard"
'
'        Case ".vcg"
'            SelectExt = "application/vnd.groove-vcard"
'
'        Case ".vcl"
'            SelectExt = "text/html"
'
'        Case ".vfw"
'            SelectExt = "video/x-msvideo"
'
'        Case ".vpg"
'            SelectExt = "application/x-vpeg"
'
'        Case ".vsl"
'            SelectExt = "application/x-cnet-vsl"
'
'        Case ".wav"
'            SelectExt = "audio/wav"
'
'        Case ".wax"
'            SelectExt = "audio/x-ms-wax"
'
'        Case ".wiz"
'            SelectExt = "application/msword"
'
'        Case ".wm"
'            SelectExt = "video/x-ms-wm"
'
'        Case ".wma"
'            SelectExt = "audio/x-ms-wma"
'
'        Case ".wmd"
'            SelectExt = "application/x-ms-wmd"
'
'        Case ".wme"
'            SelectExt = "text/xml"
'
'        Case ".wmp"
'            SelectExt = "video/x-ms-wmp"
'
'        Case ".wms"
'            SelectExt = "application/x-ms-wms"
'
'        Case ".wmv"
'            SelectExt = "video/x-ms-wmv"
'
'        Case ".wmx"
'            SelectExt = "video/x-ms-wmx"
'
'        Case ".wmz"
'            SelectExt = "application/x-ms-wmz"
'
'        Case ".wsc"
'            SelectExt = "text/scriptlet"
'
'        Case ".wvx"
'            SelectExt = "video/x-ms-wvx"
'
'        Case ".xbm"
'            SelectExt = "image/x-xbitmap"
'
'        Case ".xls"
'            SelectExt = "application/vnd.ms-excel"
'
'        Case ".xls"
'            SelectExt = "application/x-msexcel"
'
'        Case ".xml"
'            SelectExt = "text/xml"
'
'        Case ".xpl"
'            SelectExt = "audio/mpegurl"
'
'        Case ".xsl"
'            SelectExt = "text/xml"
'
'        Case ".XXE"
'            SelectExt = "application/x-compressed"
'
'        Case ".ymg"
'            SelectExt = "application/ymsgr"
'
'        Case ".yps"
'            SelectExt = "application/ymsgr"
'
'        Case ".z"
'            SelectExt = "application/x-compress"
'
'        Case ".zip"
'            SelectExt = "application/x-zip-compressed"
'
'        Case "ratfile"
'            SelectExt = "application/rat-file"
'
'        Case "smafile"
'            SelectExt = "application/x-smb-directive"
'
'        Case Else
'            SelectExt = "application/octet-stream"
'    End Select
'
'End Function
'
'Private Function MIMEFileHeader(MIMEBoundary As String, FilePath As String)
'    Dim sResult As String
'    Dim sFileName As String
'    Dim pintC
'
'    pintC = InStrRev(FilePath, "\")
'    sFileName = Mid(FilePath, pintC + 1)
'
'    sResult = vbCrLf & "--" & MIMEBoundary & vbCrLf
'    sResult = sResult & "Content-Type: " & SelectExt(FilePath) & "; " & "name=" & Chr(34) & sFileName & Chr(34) & vbNewLine
'    sResult = sResult & "Content-Transfer-Encoding: base64" & vbCrLf
'    sResult = sResult & "Content-Disposition: attachment; filename=" & Chr(34) & sFileName & Chr(34) & vbCrLf
'
'    MIMEFileHeader = sResult
'End Function
'
'Public Function Base64EncodeFile(sInputFile As String, sOutputFile As String, Optional PrgBar As ProgressBar) As Byte()
'    Dim bTemp() As Byte
'    Dim fh As Long
'
'    fh = FreeFile(0)
'    Open sInputFile For Binary Access Read As fh
'        ReDim bTemp(0 To LOF(fh) - 1)
'        Get fh, , bTemp
'    Close fh
'
'    bTemp = EncodeArr(bTemp, PrgBar)
'    Open sOutputFile For Binary Access Write As fh
'        Put fh, , bTemp
'    Close fh
'
'    Base64EncodeFile = bTemp
'End Function
'
''Converts an array of bytes to a string.
'Private Sub ByteArrayToString(ByteArray() As Byte, StringOut As String)
'  Dim lBytes As Long
'
'  If LBound(ByteArray) > 0 Then Exit Sub 'lBound MUST be 0
'  lBytes = UBound(ByteArray) + 1
'  StringOut = String$(lBytes, 0)
'
'  RtlMoveMemory ByVal StringOut, ByteArray(0), lBytes
'End Sub
'
'
''Converts a string to an array of bytes. Just like Asc() on every character.
'Private Sub StringToByteArray(ByVal StringIn As String, ByteArray() As Byte)
'    Dim lBytes As Long
'
'    If Len(StringIn) = 0 Then Exit Sub
'    lBytes = Len(StringIn)
'    ReDim ByteArray(lBytes - 1)
'
'    RtlMoveMemory ByteArray(0), ByVal StringIn, lBytes
'End Sub
'
'
''Takes sets of 3 characters (binary) and returns sets of 4 characters (ASCII)
'Private Function EncodeString(ByVal InString As String) As String
'    Dim OutString As String
'    Dim i As Integer
'    Dim UnCodedArray() As Byte
'    Dim CodedArray() As Byte
'
'    'Pad will null characters if necessary
'    If Len(InString) Mod 3 <> 0 Then
'        InString = InString & String(3 - Len(InString) Mod 3, Chr$(0))
'    End If
'
'    'Convert string to a byte array. This is MUCH faster than the Asc/Chr combo.
'    StringToByteArray InString, UnCodedArray()
'
'    'Make sure our output array is the correct size
'    ReDim CodedArray(((Len(InString) / 3) * 4) - 1)
'
'    For i = 0 To (Len(InString) / 3) - 1
'        'Encode 4 bytes at a time
'        CodedArray(i * 4 + 0) = UnCodedArray(i * 3 + 0) \ 4 + 32
'        CodedArray(i * 4 + 1) = ((UnCodedArray(i * 3 + 0) Mod 4) * 16) + (UnCodedArray(i * 3 + 1) \ 16 + 32)
'        CodedArray(i * 4 + 2) = ((UnCodedArray(i * 3 + 1) Mod 16) * 4) + (UnCodedArray(i * 3 + 2) \ 64 + 32)
'        CodedArray(i * 4 + 3) = (UnCodedArray(i * 3 + 2) Mod 64) + 32
'
'        'Check for spaces and eliminate them
'        If CodedArray(i * 4 + 0) = 32 Then CodedArray(i * 4 + 0) = 96
'        If CodedArray(i * 4 + 1) = 32 Then CodedArray(i * 4 + 1) = 96
'        If CodedArray(i * 4 + 2) = 32 Then CodedArray(i * 4 + 2) = 96
'        If CodedArray(i * 4 + 3) = 32 Then CodedArray(i * 4 + 3) = 96
'    Next i
'    ByteArrayToString CodedArray(), OutString
'    EncodeString = OutString
'End Function
'
'
''Gives a character representing the number of bytes to be decoded from a line. "M" is a full line (45 bytes)
'Private Function ENC(ByVal i As Integer) As String
'    If i = 0 Then
'        ENC = "`"
'    Else
'        i = i + 32
'        ENC = Chr(i)
'    End If
'End Function
'
'
''Returns the filename part from a full filename.
'Private Function StripFilename(ByVal Filename As String) As String
'    While InStr(Filename, "\") <> 0
'        Filename = Right(Filename, Len(Filename) - InStr(Filename, "\"))
'    Wend
'    StripFilename = Filename
'End Function
'
'
'Private Function UUEncodeFile(ByVal InputFilename As String, ByVal OutputFileName As String) As Boolean
'    Dim TotalParts As Long
'    Dim Remain As Long
'    Dim InString As String
'    Dim CurrentPos As Long
'    Dim i As Long
'    Dim FNumIn As Long
'    Dim FNumOut As Long
'
'    On Error GoTo ErrorHandler
'
'    'Open the original file as binary read
'    FNumIn = FreeFile
'    Open InputFilename For Binary Access Read Shared As #FNumIn
'
'    'Open the target file as binary write
'    FNumOut = FreeFile
'    Open OutputFileName For Binary Access Write As #FNumOut
'
'    'Write the standard uuencode file header
'    Put #FNumOut, , "begin 644 " + StripFilename(InputFilename) + vbCrLf
'
'    'TotalParts equals the number of full sized chunks of "PortionSize" bytes.
'    TotalParts = LOF(1) \ PortionSize
'
'    'Remainder equals how many remaining bytes are at the end of the file.
'    Remain = LOF(1) Mod PortionSize
'
'    'CurrentPos is the current file position
'    CurrentPos = 1
'
'    'Set up InString as a buffer the size of PortionSize
'    InString = String(PortionSize, 0)
'
'    'For loop to read the portions one by one
'    For i = 1 To TotalParts
'        Get #FNumIn, CurrentPos, InString
'        'use the ENC() for standard uuencode compatibility, pad "M"
'        Put #FNumOut, , ENC(PortionSize) + EncodeString(InString) + vbCrLf
'        CurrentPos = CurrentPos + PortionSize
'        RaiseEvent EncodeProgress(CurrentPos / LOF(1), CurrentPos)
'    Next
'
'    'Set up InString again as a buffer the size of the remaining bytes
'    InString = String(Remain, 0)
'
'    'get the remaining bytes toward end of the file
'    Get #FNumIn, CurrentPos, InString
'
'    'get the remaining bytes size and calculate ENC() for the last line
'    Put #FNumOut, , ENC(LOF(1) - CurrentPos + 1) + EncodeString(InString) + vbCrLf
'
'    'put "end" for standard uuencode compatibility
'    Put #FNumOut, , ENC(0) + vbCrLf + "end" + vbCrLf
'
'    'All Done!
'    Close #FNumIn
'    Close #FNumOut
'    UUEncodeFile = True
'    Exit Function
'
'ErrorHandler:
'    Close #FNumIn
'    Close #FNumOut
'    UUEncodeFile = False
'End Function
'
'
'Private Function EncodeFile(sFileInput As String, dType As DecodeType, sFileOutput As String, Optional PrgBar As ProgressBar) As Byte()
'    Dim pint As Integer
'    Select Case dType
'        Case uu
'            Call UUEncodeFile(sFileInput, sFileOutput)
'        Case base64
'            EncodeFile = Base64EncodeFile(sFileInput, sFileOutput, PrgBar)
'    End Select
'End Function
'
'Private Function GetTmpFile() As String
'    Dim pstrTmpPath As String
'    Dim pstrTmpFileName As String
'
'    pstrTmpPath = String(100, Chr$(0))
'    Call GetTempPath(100, pstrTmpPath)
'    pstrTmpPath = Left$(pstrTmpPath, InStr(pstrTmpPath, Chr$(0)) - 1)
'    pstrTmpFileName = String(260, 0)
'    Call GetTempFileName(pstrTmpPath, "NPD", 0, pstrTmpFileName)
'
'    GetTmpFile = Left$(pstrTmpFileName, InStr(pstrTmpFileName, Chr$(0)) - 1)
'
'End Function
'
'
'Public Sub AddAttachment(Filename As String, Optional PrgBar As ProgressBar)
'    Static pintC As Integer
'    Dim pintFF As Integer
'    Dim pstrTMPFile As String
'    Dim pstrBuffer() As Byte
'    Dim pstrTotal As String
'    Dim pstrHeader As String
'
'    pintC = pintC + 1
'    ReDim Preserve gstrFileNames(pintC)
'    gstrFileNames(pintC) = Filename
'    gblnAttachments = True
'    pintFF = FreeFile
'    ReDim Preserve FileCode(pintC)
'    pstrTMPFile = GetTmpFile
'
'    pstrBuffer = EncodeFile(gstrFileNames(pintC), base64, pstrTMPFile, PrgBar)
'
'
'    pstrTotal = StrConv(pstrBuffer(), vbUnicode)
'
'    pstrHeader = MIMEFileHeader(conBoundary, gstrFileNames(pintC))
'    FileCode(pintC).Code = pstrHeader _
'            & vbCrLf & pstrTotal & vbCrLf
'    FileCode(pintC).Filename = gstrFileNames(pintC)
'    FileCode(pintC).Send = True
'    Kill pstrTMPFile
'
'End Sub
'
'
'Private Function EncodeArr(bInput() As Byte, Optional PrgBar As ProgressBar) As Byte()
'   Dim bOutput() As Byte
'   Dim k As Long
'   Dim l As Long
'   Dim i As Long
'   Dim evenBound As Long
'   Dim CurrentOut As Long
'   Dim b As Byte
'   Dim c As Byte
'   Dim d As Byte
'   Dim LineLength As Long
'
'   k = LBound(bInput)
'   l = UBound(bInput)
'
'   'Calculate the input size
'   i = l - k + 1
'
'   'Calculate the output size
'   Select Case i Mod 3
'      Case 0:
'         i = (i \ 3) * 4
'         evenBound = l
'      Case 1:
'         i = ((i \ 3) * 4) + 4
'         evenBound = l - 1
'      Case 2:
'         i = ((i \ 3) * 4) + 4
'         evenBound = l - 2
'      Case 3:
'         i = ((i \ 3) * 4) + 4
'         evenBound = l - 3
'   End Select
'
'   'Add in the line feeds.
'   If i Mod MAX_LINELENGTH = 0 Then
'      i = i + (i \ MAX_LINELENGTH) * 2 - 2
'   Else
'      i = i + (i \ MAX_LINELENGTH) * 2
'   End If
'
'   'Size the output array
'   ReDim bOutput(0 To i - 1)
'
'   CurrentOut = 0
'   LineLength = 0
'
'   On Error Resume Next
'   PrgBar.Max = evenBound
'   On Error GoTo 0
'
'   For i = k To evenBound Step 3
'      b = bInput(i)
'      c = bInput(i + 1)
'      d = bInput(i + 2)
'      bOutput(CurrentOut) = m_Index1(b And &HFC)
'      bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
'      bOutput(CurrentOut + 2) = m_Index3((c And &HF) Or (d And &HC0))
'      bOutput(CurrentOut + 3) = m_Index4(d And &H3F)
'      CurrentOut = CurrentOut + 4
'      LineLength = LineLength + 4
'
'      If LineLength >= MAX_LINELENGTH Then
'         If i <> l - 2 Then  ' If this is the last line, don't add crlf
'            bOutput(CurrentOut) = CHAR_CR
'            bOutput(CurrentOut + 1) = CHAR_LF
'         End If
'         CurrentOut = CurrentOut + 2
'         LineLength = 0
'      End If
'      On Error Resume Next
'      PrgBar.Value = i
'      On Error GoTo 0
'   Next i
'
'   Select Case l - i
'      Case 1:
'         b = bInput(i)
'         c = bInput(i + 1)
'         d = 0
'         bOutput(CurrentOut) = m_Index1(b And &HFC)
'         bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
'         bOutput(CurrentOut + 2) = m_Index3((c And &HF) Or (d And &HC0))
'         bOutput(CurrentOut + 3) = CHAR_EQUAL
'         CurrentOut = CurrentOut + 4
'         LineLength = LineLength + 4
'      Case 0:
'         b = bInput(i)
'         c = 0
'         bOutput(CurrentOut) = m_Index1(b And &HFC)
'         bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
'         bOutput(CurrentOut + 2) = CHAR_EQUAL
'         bOutput(CurrentOut + 3) = CHAR_EQUAL
'         CurrentOut = CurrentOut + 4
'         LineLength = LineLength + 4
'   End Select
'
'   EncodeArr = bOutput
'End Function
'
'
'
'' Encode a string to a string.
'Private Function Encode(sInput As String) As String
'   Dim bTemp() As Byte
'
'   'Convert to a byte array then convert.
'   'This is faster the repetitive calls to asc() or chr$()
'   bTemp = StrConv(sInput, vbFromUnicode)
'   Encode = StrConv(EncodeArr(bTemp), vbUnicode)
'End Function
'
'Public Function EncodeStringBase64(pstrString) As String
'
'    EncodeStringBase64 = Encode(Trim(pstrString))
'
'End Function
'
'
'Private Function Base64String(ByVal pstrCode As String) As String
'
'    Base64String = Decode(Trim(pstrCode))
'
'End Function
'
'Private Function DecodeArr(bInput() As Byte) As Byte()
'   Dim bOutput() As Byte
'   Dim OutLength As Long
'   Dim CurrentOut As Long
'
'   Dim k As Long
'   Dim l As Long
'   Dim i As Long
'   Dim j As Long
'
'   Dim b As Byte
'   Dim c As Byte
'   Dim d As Byte
'   Dim e As Byte
'
'   On Error GoTo ErrorHandler
'
'   k = LBound(bInput)
'   l = UBound(bInput)
'
'   'Calculate the length of the input
'   i = l - k + 1
'
'   'Calculate the expected length of the output
'   'It should be no more (but may possible be less)
'   j = i Mod (MAX_LINELENGTH + 2)
'   If j = 0 Then
'      OutLength = (i \ (MAX_LINELENGTH + 2)) * (MAX_LINELENGTH \ 4) * 3
'   Else
'      j = (j / 4) * 3
'      If bInput(l) = CHAR_EQUAL Then j = j - 1
'      If bInput(l - 1) = CHAR_EQUAL Then j = j - 1
'      OutLength = (i \ (MAX_LINELENGTH + 2)) * (MAX_LINELENGTH \ 4) * 3 + j
'   End If
'
'   'Allocate the output
'   ReDim bOutput(0 To OutLength - 1)
'
'   CurrentOut = 0
'
'   For i = k To l
'      Select Case bInput(i)
'         Case CHAR_CR
'            'Do nothing
'         Case CHAR_LF
'            'Do nothing
'         Case Else
'            If l - i >= 3 Then
'               b = bInput(i)
'               c = bInput(i + 1)
'               d = bInput(i + 2)
'               e = bInput(i + 3)
'
'               If e <> CHAR_EQUAL Then
'
'                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
'                  bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
'                  bOutput(CurrentOut + 2) = m_ReverseIndex3(d, 1) Or m_ReverseIndex4(e)
'                  CurrentOut = CurrentOut + 3
'                  i = i + 3
'               ElseIf d <> CHAR_EQUAL Then
'                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
'                  bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
'                  CurrentOut = CurrentOut + 2
'                  i = i + 3
'               Else
'                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
'                  CurrentOut = CurrentOut + 1
'                  i = i + 3
'               End If
'
'            Else
'               'Possible input code error, but may also be
'               'an extra CrLf, so we will ignore it.
'            End If
'      End Select
'   Next i
'
'   'On properly formed input we should have to do this.
'   If OutLength <> CurrentOut + 1 Then
'      ReDim Preserve bOutput(0 To CurrentOut - 1)
'   End If
'
'   DecodeArr = bOutput
'
'   Exit Function
'ErrorHandler:
'    If Err.Number = 9 Then
'        Exit Function
'    End If
'End Function
'
'
'
''Decode a string to a string.
'Private Function Decode(sInput As String) As String
'   Dim bTemp() As Byte
'
'   'Convert to a byte array then convert.
'   'This is faster the repetitive calls to asc() or chr$()
'   bTemp = StrConv(sInput, vbFromUnicode)
'   Decode = StrConv(DecodeArr(bTemp), vbUnicode)
'End Function
'
'
'Private Sub SocketSend(Socket As Winsock, Data As String)
'    If gblnSSL Then
'        Call SSLSend(Socket, Data)
'    Else
'        Socket.SendData Data
'    End If
'End Sub
'
'Private Sub CertificateToPublicKey()
'
'    'Create CryptoAPI Blob from Certificate
'    Const lPbkLen As Long = 1024
'    Dim lOffset As Long
'    Dim lStart As Long
'    Dim sBlkLen As String
'    Dim sRevKey As String
'    Dim ASNStart As Long
'    Dim ASNKEY As String
'
'    lOffset = CLng(lPbkLen \ 8)
'    lStart = 5 + (lOffset \ 128) * 2
'
'    ASNStart = InStr(1, ENCODED_CERT, Chr(48) & Chr(129) & Chr(137) & Chr(2) & Chr(129) & Chr(129) & Chr(0)) + lStart
'    ASNKEY = Mid(ENCODED_CERT, ASNStart, 128)
'
'    sRevKey = ReverseString(ASNKEY)
'
'    sBlkLen = CStr(Hex(lPbkLen \ 256))
'    If Len(sBlkLen) = 1 Then sBlkLen = "0" & sBlkLen
'
'    PUBLIC_KEY = (HexToBin( _
'            "06020000" & _
'            "00A40000" & _
'            "52534131" & _
'            "00" & sBlkLen & "0000" & _
'            "01000100") & sRevKey)
'
'End Sub
'
'
'
'Private Function VerifyMAC(ByVal DecryptedRecord As String) As Boolean
'
'    'Verify the Message Authentication Code
'    Dim PrependedMAC As String
'    Dim RecordData As String
'    Dim CalculatedMAC As String
'
'    PrependedMAC = Mid(DecryptedRecord, 1, 16)
'    RecordData = Mid(DecryptedRecord, 17)
'
'    CalculatedMAC = MD5_Hash(CLIENT_READ_KEY & RecordData & RecvSequence)
'
'    Call IncrementRecv
'
'    If CalculatedMAC = PrependedMAC Then
'        VerifyMAC = True
'    Else
'        VerifyMAC = False
'    End If
'
'End Function
'
'Private Function SendSequence() As String
'    'Convert Send Counter to a String
'    Dim TempString As String
'    Dim TempSequence As Double
'    Dim TempByte As Double
'    Dim i As Integer
'
'    TempSequence = SEND_SEQUENCE_NUMBER
'
'    For i = 1 To 4
'        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
'        TempSequence = Int(TempSequence / 256)
'        TempString = Chr(TempByte) & TempString
'    Next
'
'    SendSequence = TempString
'
'End Function
'
'Private Function RecvSequence() As String
'Dim i As Integer
'
'    'Convert Receive Counter to a String
'    Dim TempString As String
'    Dim TempSequence As Double
'    Dim TempByte As Double
'
'    TempSequence = RECV_SEQUENCE_NUMBER
'
'    For i = 1 To 4
'        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
'        TempSequence = Int(TempSequence / 256)
'        TempString = Chr(TempByte) & TempString
'    Next
'
'    RecvSequence = TempString
'
'End Function
'
'Private Sub SendClientHello(ByRef Socket As Winsock)
'
'    'Send Client Hello
'    Layer = 0
'
'    Call GenerateRandomBytes(16, CHALLENGE_DATA)
'
'    SEND_SEQUENCE_NUMBER = 0
'    RECV_SEQUENCE_NUMBER = 0
'
'    CLIENT_HELLO = Chr(1) & _
'                    Chr(0) & Chr(2) & _
'                    Chr(0) & Chr(3) & _
'                    Chr(0) & Chr(0) & _
'                    Chr(0) & Chr(Len(CHALLENGE_DATA)) & _
'                    Chr(1) & Chr(0) & Chr(128) & _
'                    CHALLENGE_DATA
'
'    If Socket.State = 7 Then Socket.SendData AddRecordHeader(CLIENT_HELLO)
'
'End Sub
'
'Private Sub SendMasterKey(ByRef Socket As Winsock)
'
'    'Send Master Key
'    Layer = 1
'
'    Call GenerateRandomBytes(32, MASTER_KEY)
'
'    Call CertificateToPublicKey
'
'    Socket.SendData AddRecordHeader(Chr(2) & _
'                                    Chr(1) & Chr(0) & Chr(128) & _
'                                    Chr(0) & Chr(0) & _
'                                    Chr(0) & Chr(128) & _
'                                    Chr(0) & Chr(0) & _
'                                    ExportKeyBlob(MASTER_KEY, CLIENT_READ_KEY, CLIENT_WRITE_KEY, CHALLENGE_DATA, CONNECTION_ID, PUBLIC_KEY))
'
'End Sub
'
'Private Sub SendClientFinish(ByRef Socket As Winsock)
'
'    'Send ClientFinished Message
'    Layer = 2
'    Call SSLSend(Socket, Chr(3) & CONNECTION_ID)
'
'End Sub
'
'Private Sub SSLSend(ByRef Socket As Winsock, ByVal Plaintext As String)
'
'    'Send Plaintext as an Encrypted SSL Record
'    Dim SSLRecord As String
'    Dim OtherPart As String
'    Dim SendAnother As Boolean
'
'    If Len(Plaintext) > 32751 Then
'        SendAnother = True
'        Plaintext = Mid(Plaintext, 1, 32751)
'        OtherPart = Mid(Plaintext, 32752)
'    Else
'        SendAnother = False
'    End If
'
'    SSLRecord = AddMACData(Plaintext)
'    SSLRecord = RC4_Encrypt(SSLRecord)
'    SSLRecord = AddRecordHeader(SSLRecord)
'
'    Socket.SendData SSLRecord
'
'    If SendAnother = True Then
'        Call SSLSend(Socket, OtherPart)
'    End If
'
'End Sub
'
'Private Function AddMACData(ByVal Plaintext As String) As String
'
'    'Prepend MAC Data to the Plaintext
'    AddMACData = MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext
'
'End Function
'
'Private Function AddRecordHeader(ByVal RecordData As String) As String
'
'    'Prepend SLL Record Header to the Data Record
'    Dim FirstChar As String
'    Dim LastChar As String
'    Dim TheLen As Long
'
'    TheLen = Len(RecordData)
'
'    FirstChar = Chr(128 + (TheLen \ 256))
'    LastChar = Chr(TheLen Mod 256)
'
'    AddRecordHeader = FirstChar & LastChar & RecordData
'
'    Call IncrementSend
'
'End Function
'
'Private Sub IncrementSend()
'
'    'Increment Counter for Each Record Sent
'    SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
'    If SEND_SEQUENCE_NUMBER = 4294967296# Then SEND_SEQUENCE_NUMBER = 0
'
'End Sub
'
'Private Sub IncrementRecv()
'
'    'Increment Counter for Each Record Received
'    RECV_SEQUENCE_NUMBER = RECV_SEQUENCE_NUMBER + 1
'    If RECV_SEQUENCE_NUMBER = 4294967296# Then RECV_SEQUENCE_NUMBER = 0
'
'End Sub
'
'Private Function BytesToLen(ByVal TwoBytes As String) As Long
'
'    'Convert Byte Pair to Packet Length
'    Dim FirstByteVal As Long
'    FirstByteVal = Asc(Left(TwoBytes, 1))
'    If FirstByteVal >= 128 Then FirstByteVal = FirstByteVal - 128
'
'    BytesToLen = 256 * FirstByteVal + Asc(Right(TwoBytes, 1))
'
'End Function
'
'Private Function HexToBin(ByVal HexString As String) As String
'
'    'Convert a Hexadecimal String to characters
'    Dim BinString As String
'    Dim i As Integer
'
'    For i = 1 To Len(HexString) Step 2
'        BinString = BinString & Chr(Val("&H" & Mid(HexString, i, 2)))
'    Next i
'
'    HexToBin = BinString
'
'End Function
'
'Private Function ReverseString(ByVal TheString As String) As String
'
'    'Reverse String
'    Dim Reversed As String
'    Dim i As Integer
'
'    For i = Len(TheString) To 1 Step -1
'        Reversed = Reversed & Mid(TheString, i, 1)
'    Next i
'    ReverseString = Reversed
'
'End Function
'
'Private Sub ProcessData(ByVal TheData As String, Socket As Winsock)
'    Dim pblnStatus As Boolean
'
'    Debug.Print TheData
'
'    Select Case State
'        Case Connect
'            pblnStatus = CheckResponse(TheData, "220", Socket)
'            Call SocketSend(Socket, "HELO" & gstrDomainName & vbCrLf)
'            If pblnStatus = True Then
'                State = State + 1
'            Else
'                Socket.Close
'                State = 0
'                gblnDone = True
'            End If
'        Case helo
'            pblnStatus = CheckResponse(TheData, "250", Socket)
'            If pblnStatus = True Then
'                Call SocketSend(Socket, "MAIL FROM: " & gstrFromAddress & vbCrLf)
'                State = State + 1
'            Else
'                Socket.Close
'                State = 0
'                gblnDone = True
'                Exit Sub
'            End If
'        Case MailFrom
'            pblnStatus = CheckResponse(TheData, "250", Socket)
'            If pblnStatus = True Then
'                If gintSent > UBound(gstrAllAddresses) And gblnBCCMode = False And gstrAllAddresses(0) <> "" Then
'                    gintsentTo = 0
'                    gintsentCC = 0
'                    State = State + 1
'                    Call ProcessData("250 ", Socket)
'                Else
'                    If gstrAllAddresses(0) = "" Then
'                        gblnDone = True
'                        Exit Sub
'                    End If
'                    If gblnBCCMode = True Then
'                        pintSend = 0
'                        gintSent = 0
'                    End If
'                    Call SocketSend(Socket, "RCPT TO:" & gstrAllAddresses(gintSent) & vbCrLf)
'                End If
'                If gblnBCCMode Then
'                    State = State + 1
'                    gintSent = 0
'                Else
'                    gintSent = gintSent + 1
'                End If
'            Else
'                If Left(TheData, 3) = "530" Then
'                    Call SocketSend(Socket, "AUTH LOGIN" & vbCrLf)
'                ElseIf Left(TheData, 3) = "334" Then
'                    Dim pstrIn As String
'                    Dim pstrAnswer As String
'                    pstrIn = Mid(TheData, 4)
'                    pstrIn = Left(pstrIn, Len(pstrIn) - 2)
'                    pstrIn = Base64String(Trim(pstrIn))
'                    If InStr(LCase(pstrIn), "user") Then
'                        pstrAnswer = gstrUserID
'                        pstrAnswer = EncodeStringBase64(gstrUserID)
'                    ElseIf InStr(LCase(pstrIn), "pass") Then
'                        pstrAnswer = gstrUserPass
'                        pstrAnswer = EncodeStringBase64(pstrAnswer)
'                    End If
'                    Call SocketSend(Socket, pstrAnswer & vbCrLf)
'                    ElseIf Left(TheData, 3) = "501" Then
'                        gblnDone = True
'                    ElseIf Left(TheData, 3) = "235" Then
'                        Call SocketSend(Socket, "MAIL FROM: " & gstrFromAddress & vbCrLf)
'                        State = MailFrom
'                    End If
'                End If
'        Case sendto
'            Call SocketSend(Socket, "DATA" & vbCrLf)
'            State = State + 1
'        Case Data
'            Call RTBX.LoadFile(gstrMessageFileName, rtfText)
'            If gblnBCCMode = True Then
'                Dim pintLine As Integer
'                Dim pintPos As Integer
'                Dim pintStringStart As Integer
'                Dim pintStringEnd As Integer
'                With RTBX
'                    pintPos = .Find("CC: ")
'                    pintLine = .GetLineFromChar(pintPos)
'                    pintStringEnd = InStr(pintPos, .Text, vbCrLf)
'                    .SelStart = pintPos
'                    .SelLength = pintStringEnd - pintPos
'                    .SelText = "CC: "
'                    pintPos = .Find("TO: ")
'                    pintLine = .GetLineFromChar(pintPos)
'                    pintStringEnd = InStr(pintPos, .Text, vbCrLf)
'                    .SelStart = pintPos
'                    .SelLength = pintStringEnd - pintPos
'                    .SelText = "TO: " & gstrAllAddresses(0)
'                End With
'            End If
'            gstrData = Split(RTBX.Text, vbCrLf)
'            If gstrBCC(0) <> "" Then
'                ProgBar.Max = (UBound(gstrData) * (UBound(gstrBCC) + 1))
'            Else
'                ProgBar.Max = UBound(gstrData)
'            End If
'            For pintSend = 0 To UBound(gstrData)
'                Call SocketSend(Socket, gstrData(pintSend) & vbCrLf)
'                On Error Resume Next
'                ProgBar.Value = gintTransfer
'                gintTransfer = gintTransfer + 1
'                On Error GoTo 0
'                DoEvents
'            Next pintSend
'            State = State + 1
'            Call SocketSend(Socket, vbCrLf & "." & vbCrLf)
'        Case MessageData
'            gblnDone = True
'            State = State + 1
'
'    End Select
'
'End Sub
'
'Private Function CheckResponse(pstrGot, PstrExpect, Socket As Winsock) As Boolean
'    Dim pstrResp As String
'
'    pstrResp = Mid(pstrGot, 1, 3)
'
'    If StrComp(pstrResp, PstrExpect) = 0 Then
'        CheckResponse = True
'    Else
'        CheckResponse = False
'    End If
'
'End Function
'
'Private Function RC4_Encrypt(ByVal Plaintext As String) As String
'
'    'Encrypt with Client Write Key
'    Dim lngLength As Long
'    Dim lngReturnValue As Long
'
'    lngLength = Len(Plaintext)
'    lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)
'
'    RC4_Encrypt = Plaintext
'
'End Function
'
'
'Private Function RC4_Decrypt(ByVal Ciphertext As String) As String
'
'    'Decrypt with Client Read Key
'    Dim lngLength As Long
'    Dim lngReturnValue As Long
'
'    lngLength = Len(Ciphertext)
'    lngReturnValue = CryptDecrypt(hClientReadKey, 0, False, 0, Ciphertext, lngLength)
'
'    RC4_Decrypt = Ciphertext
'
'End Function
'
'
'Private Function MD5_Hash(ByVal TheString As String) As String
'
'    'Digest a String using MD5
'    Dim lngReturnValue As Long
'    Dim strHash As String
'    Dim hHash As Long
'    Dim lngHashLen As Long
'
'    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
'    lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
'    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
'    strHash = String(lngHashLen, vbNullChar)
'    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)
'
'    If hHash <> 0 Then CryptDestroyHash hHash
'
'    MD5_Hash = strHash
'
'End Function
'
'Private Function GenerateRandomBytes(ByVal Length As Long, ByRef TheString As String) As Boolean
'
'    'Generate Random Bytes
'    Dim i As Integer
'
'    Randomize
'    TheString = ""
'    For i = 1 To Length
'        TheString = TheString & Chr(Int(Rnd * 256))
'    Next
'
'    GenerateRandomBytes = CryptGenRandom(hCryptProv, Length, TheString)
'
'End Function
'
'
'Private Function ExportKeyBlob(ByRef StrMasterKey As String, ByRef StrReadKey As String, ByRef StrWriteKey As String, ByVal StrChallenge As String, ByVal StrConnectionID As String, ByVal StrPublicKey As String) As String
'
'    'Create Keys and Return PKCS Block
'    Dim lngReturnValue As Long
'    Dim lngLength As Long
'    Dim rgbBlob As String
'    Dim hPublicKey As Long
'
'    Call CreateKey(hMasterKey, StrMasterKey)
'    StrMasterKey = MD5_Hash(StrMasterKey)
'
'    Call CreateKey(hClientReadKey, StrMasterKey & "0" & StrChallenge & StrConnectionID)
'    Call CreateKey(hClientWriteKey, StrMasterKey & "1" & StrChallenge & StrConnectionID)
'
'    StrReadKey = MD5_Hash(StrMasterKey & "0" & StrChallenge & StrConnectionID)
'    StrWriteKey = MD5_Hash(StrMasterKey & "1" & StrChallenge & StrConnectionID)
'
'    lngReturnValue = CryptImportKey(hCryptProv, StrPublicKey, Len(StrPublicKey), 0, 0, hPublicKey)
'
'    lngReturnValue = CryptExportKey(hMasterKey, hPublicKey, SIMPLEBLOB, 0, vbNull, lngLength)
'    rgbBlob = String(lngLength, 0)
'    lngReturnValue = CryptExportKey(hMasterKey, hPublicKey, SIMPLEBLOB, 0, rgbBlob, lngLength)
'
'    If hPublicKey <> 0 Then CryptDestroyKey hPublicKey
'    If hMasterKey <> 0 Then CryptDestroyKey hMasterKey
'
'    ExportKeyBlob = StrReverse(Right(rgbBlob, 128))
'
'End Function
'
'
'Private Sub CreateKey(ByRef KeyName As Long, ByVal HashData As String)
'
'    'Create a Session Key from a Hash
'    Dim lngParams As Long
'    Dim lngReturnValue As Long
'    Dim lngHashLen As Long
'    Dim hHash As Long
'
'    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
'    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
'
'    lngReturnValue = CryptHashData(hHash, HashData, Len(HashData), 0)
'    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
'
'    lngParams = GEN_KEY_BITS Or CRYPT_EXPORTABLE
'    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, KeyName)
'    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a session key (CryptDeriveKey API)"
'
'    If hHash <> 0 Then CryptDestroyHash hHash
'
'End Sub
'
'Private Sub Class_Initialize()
'
'    'Initiate Secure Session
'    Dim lngReturnValue As Long
'    Dim TheAnswer As Long
'
'    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container
'
'    If lngReturnValue = 0 Then
'        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
'        If lngReturnValue = 0 Then TheAnswer = MsgBox("GCN has detected that you do not have the required High Encryption Pack installed." & vbCrLf & "Would like to download this pack from Microsoft's website?", 16 + vbYesNo)
'    End If
'
''    If TheAnswer = vbYes Then
''        Call Shell("START http://www.microsoft.com/windows/ie/downloads/recommended/128bit/default.asp", vbHide)
''        SocketSSL_Close
''    End If
''
''    If TheAnswer = vbNo Then
''        SocketSSL_Close
''    End If
'
'   Dim i As Long
'
'   'Setup the encodeing and decoding lookup arrays.
'   'Essentially we speed up the routine by pre-shifting
'   'the data so it only needs combined with And and Or.
'   m_Index4(0) = 65 'Asc("A")
'   m_Index4(1) = 66 'Asc("B")
'   m_Index4(2) = 67 'Asc("C")
'   m_Index4(3) = 68 'Asc("D")
'   m_Index4(4) = 69 'Asc("E")
'   m_Index4(5) = 70 'Asc("F")
'   m_Index4(6) = 71 'Asc("G")
'   m_Index4(7) = 72 'Asc("H")
'   m_Index4(8) = 73 'Asc("I")
'   m_Index4(9) = 74 'Asc("J")
'   m_Index4(10) = 75 'Asc("K")
'   m_Index4(11) = 76 'Asc("L")
'   m_Index4(12) = 77 'Asc("M")
'   m_Index4(13) = 78 'Asc("N")
'   m_Index4(14) = 79 'Asc("O")
'   m_Index4(15) = 80 'Asc("P")
'   m_Index4(16) = 81 'Asc("Q")
'   m_Index4(17) = 82 'Asc("R")
'   m_Index4(18) = 83 'Asc("S")
'   m_Index4(19) = 84 'Asc("T")
'   m_Index4(20) = 85 'Asc("U")
'   m_Index4(21) = 86 'Asc("V")
'   m_Index4(22) = 87 'Asc("W")
'   m_Index4(23) = 88 'Asc("X")
'   m_Index4(24) = 89 'Asc("Y")
'   m_Index4(25) = 90 'Asc("Z")
'   m_Index4(26) = 97 'Asc("a")
'   m_Index4(27) = 98 'Asc("b")
'   m_Index4(28) = 99 'Asc("c")
'   m_Index4(29) = 100 'Asc("d")
'   m_Index4(30) = 101 'Asc("e")
'   m_Index4(31) = 102 'Asc("f")
'   m_Index4(32) = 103 'Asc("g")
'   m_Index4(33) = 104 'Asc("h")
'   m_Index4(34) = 105 'Asc("i")
'   m_Index4(35) = 106 'Asc("j")
'   m_Index4(36) = 107 'Asc("k")
'   m_Index4(37) = 108 'Asc("l")
'   m_Index4(38) = 109 'Asc("m")
'   m_Index4(39) = 110 'Asc("n")
'   m_Index4(40) = 111 'Asc("o")
'   m_Index4(41) = 112 'Asc("p")
'   m_Index4(42) = 113 'Asc("q")
'   m_Index4(43) = 114 'Asc("r")
'   m_Index4(44) = 115 'Asc("s")
'   m_Index4(45) = 116 'Asc("t")
'   m_Index4(46) = 117 'Asc("u")
'   m_Index4(47) = 118 'Asc("v")
'   m_Index4(48) = 119 'Asc("w")
'   m_Index4(49) = 120 'Asc("x")
'   m_Index4(50) = 121 'Asc("y")
'   m_Index4(51) = 122 'Asc("z")
'   m_Index4(52) = 48 'Asc("0")
'   m_Index4(53) = 49 'Asc("1")
'   m_Index4(54) = 50 'Asc("2")
'   m_Index4(55) = 51 'Asc("3")
'   m_Index4(56) = 52 'Asc("4")
'   m_Index4(57) = 53 'Asc("5")
'   m_Index4(58) = 54 'Asc("6")
'   m_Index4(59) = 55 'Asc("7")
'   m_Index4(60) = 56 'Asc("8")
'   m_Index4(61) = 57 'Asc("9")
'   m_Index4(62) = 43 'Asc("+")
'   m_Index4(63) = 47 'Asc("/")
'
'   'Calculate the other Arrays
'   For i = 0 To 63
'      m_Index1((i * 4) And &HFC) = m_Index4(i)
'      m_Index2(((i And &HF) * 16) Or ((i And &H30) \ 16)) = m_Index4(i)
'      m_Index3((i \ 4 And &HF) Or ((i And &H3) * 64)) = m_Index4(i)
'   Next i
'
'
'   m_ReverseIndex4(65) = 0 'Asc("A")
'   m_ReverseIndex4(66) = 1 'Asc("B")
'   m_ReverseIndex4(67) = 2 'Asc("C")
'   m_ReverseIndex4(68) = 3 'Asc("D")
'   m_ReverseIndex4(69) = 4 'Asc("E")
'   m_ReverseIndex4(70) = 5 'Asc("F")
'   m_ReverseIndex4(71) = 6 'Asc("G")
'   m_ReverseIndex4(72) = 7 'Asc("H")
'   m_ReverseIndex4(73) = 8 'Asc("I")
'   m_ReverseIndex4(74) = 9 'Asc("J")
'   m_ReverseIndex4(75) = 10 'Asc("K")
'   m_ReverseIndex4(76) = 11 'Asc("L")
'   m_ReverseIndex4(77) = 12 'Asc("M")
'   m_ReverseIndex4(78) = 13 'Asc("N")
'   m_ReverseIndex4(79) = 14 'Asc("O")
'   m_ReverseIndex4(80) = 15 'Asc("P")
'   m_ReverseIndex4(81) = 16 'Asc("Q")
'   m_ReverseIndex4(82) = 17 'Asc("R")
'   m_ReverseIndex4(83) = 18 'Asc("S")
'   m_ReverseIndex4(84) = 19 'Asc("T")
'   m_ReverseIndex4(85) = 20 'Asc("U")
'   m_ReverseIndex4(86) = 21 'Asc("V")
'   m_ReverseIndex4(87) = 22 'Asc("W")
'   m_ReverseIndex4(88) = 23 'Asc("X")
'   m_ReverseIndex4(89) = 24 'Asc("Y")
'   m_ReverseIndex4(90) = 25 'Asc("Z")
'   m_ReverseIndex4(97) = 26 'Asc("a")
'   m_ReverseIndex4(98) = 27 'Asc("b")
'   m_ReverseIndex4(99) = 28 'Asc("c")
'   m_ReverseIndex4(100) = 29 'Asc("d")
'   m_ReverseIndex4(101) = 30 'Asc("e")
'   m_ReverseIndex4(102) = 31 'Asc("f")
'   m_ReverseIndex4(103) = 32 'Asc("g")
'   m_ReverseIndex4(104) = 33 'Asc("h")
'   m_ReverseIndex4(105) = 34 'Asc("i")
'   m_ReverseIndex4(106) = 35 'Asc("j")
'   m_ReverseIndex4(107) = 36 'Asc("k")
'   m_ReverseIndex4(108) = 37 'Asc("l")
'   m_ReverseIndex4(109) = 38 'Asc("m")
'   m_ReverseIndex4(110) = 39 'Asc("n")
'   m_ReverseIndex4(111) = 40 'Asc("o")
'   m_ReverseIndex4(112) = 41 'Asc("p")
'   m_ReverseIndex4(113) = 42 'Asc("q")
'   m_ReverseIndex4(114) = 43 'Asc("r")
'   m_ReverseIndex4(115) = 44 'Asc("s")
'   m_ReverseIndex4(116) = 45 'Asc("t")
'   m_ReverseIndex4(117) = 46 'Asc("u")
'   m_ReverseIndex4(118) = 47 'Asc("v")
'   m_ReverseIndex4(119) = 48 'Asc("w")
'   m_ReverseIndex4(120) = 49 'Asc("x")
'   m_ReverseIndex4(121) = 50 'Asc("y")
'   m_ReverseIndex4(122) = 51 'Asc("z")
'   m_ReverseIndex4(48) = 52 'Asc("0")
'   m_ReverseIndex4(49) = 53 'Asc("1")
'   m_ReverseIndex4(50) = 54 'Asc("2")
'   m_ReverseIndex4(51) = 55 'Asc("3")
'   m_ReverseIndex4(52) = 56 'Asc("4")
'   m_ReverseIndex4(53) = 57 'Asc("5")
'   m_ReverseIndex4(54) = 58 'Asc("6")
'   m_ReverseIndex4(55) = 59 'Asc("7")
'   m_ReverseIndex4(56) = 60 'Asc("8")
'   m_ReverseIndex4(57) = 61 'Asc("9")
'   m_ReverseIndex4(43) = 62 'Asc("+")
'   m_ReverseIndex4(47) = 63 'Asc("/")
'
'   'Calculate the other arrays.
'   For i = 0 To 255
'      If m_ReverseIndex4(i) <> 0 Then
'         m_ReverseIndex1(i) = m_ReverseIndex4(i) * 4
'
'         m_ReverseIndex2(i, 0) = m_ReverseIndex4(i) \ 16
'         m_ReverseIndex2(i, 1) = (m_ReverseIndex4(i) And &HF) * 16
'
'         m_ReverseIndex3(i, 0) = m_ReverseIndex4(i) \ 4
'         m_ReverseIndex3(i, 1) = (m_ReverseIndex4(i) And &H3) * 64
'      End If
'   Next i
'
'End Sub
'
'Private Sub Class_Terminate()
'
'    'Free up Memory
'    If hClientWriteKey <> 0 Then CryptDestroyKey hClientWriteKey
'    If hClientReadKey <> 0 Then CryptDestroyKey hClientReadKey
'    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
'
'    Set SocketSSL = Nothing
'    Set RTBX = Nothing
'
'End Sub
'
'
'Private Sub SocketSSL_Close()
'
'    'Close Socket
'    Layer = 0
'    SocketSSL.Close
'    gblnConnected = False
'
'End Sub
'
'Private Sub SocketSSL_Connect()
'Dim strHello As String
'
'    State = Connect
'
'' Added by GioRock
''--------------------------------------------------------------'
'    If InStr(gstrUserID, "@") <> 0 Then
'        strHello = Left$(gstrUserID, InStr(gstrUserID, "@") - 1)
'    Else
'        strHello = gstrUserID
'    End If
'
'    SocketSSL.SendData "EHLO " + strHello + vbCrLf
''--------------------------------------------------------------'
'
'    gblnConnected = True
'
'End Sub
'
'Public Function ConnectSSLSocket() As Integer
'    Dim pintC As Integer
'
'    With SocketSSL
'        .Close
'        .protocol = sckTCPProtocol
'        .Connect gstrServerName, glngPort
'    End With
'
'    Do
'        DoEvents
'        If gblnConnected = True Then Exit Do
'    Loop
'
'    Do
'        DoEvents
'        If gblnDone = True Then Exit Do
'    Loop
'
'    ReDim gstrCC(0)
'    ReDim gstrToAddress(0)
'    ReDim gstrAllAddresses(0)
'
'    pintC = 0
'
'    On Error Resume Next
'
'    On Error GoTo 0
'
'    Do Until pintC > UBound(gstrBCC)
'        gblnBCCMode = True
'        If gstrBCC(pintC) = "" Or gstrBCC(pintC) = "<>" Then
'            Exit Do
'        End If
'        gstrAllAddresses(0) = gstrBCC(pintC)
'        gblnDone = False
'        State = helo
'        Call ProcessData("250", SocketSSL)
'        Do
'            DoEvents
'        Loop Until gblnDone = True
'        pintC = pintC + 1
'    Loop
'
'    SocketSSL.Close
'
'    gblnBCCMode = False
'
'End Function
'
'Private Sub SocketSSL_DataArrival(ByVal bytesTotal As Long)
'
'    'Parse each SSL Record
'    Dim TheData As String
'    Static lCMDSequence As Long
'
'' Added by Giorock
''---------------------------------------------------------'
'    ' Starting a Secure Socket Layer connection
'    If lCMDSequence <> 3 Then
'        Select Case lCMDSequence
'            Case 0
'                SocketSSL.GetData TheData, vbString
''                wsSSL.SendData "EHLO giorock" + vbCrLf
''                Send on Connect Event
'            Case 1
'                SocketSSL.GetData TheData, vbString
'                SocketSSL.SendData "STARTTLS" + vbCrLf
'            Case 2
'                SocketSSL.GetData TheData, vbString
'                SendClientHello SocketSSL
'        End Select
'        lCMDSequence = lCMDSequence + 1
'        Debug.Print TheData
'        Exit Sub
'    End If
''---------------------------------------------------------'
'
'    Do
'
'        If SeekLen = 0 Then
'            If bytesTotal >= 2 Then
'                SocketSSL.GetData TheData, vbString, 2
'                SeekLen = BytesToLen(TheData)
'                bytesTotal = bytesTotal - 2
'            Else
'                Exit Sub
'            End If
'        End If
'
'        If bytesTotal >= SeekLen Then
'            SocketSSL.GetData TheData, vbString, SeekLen
'            bytesTotal = bytesTotal - SeekLen
'        Else
'            Exit Sub
'        End If
'
'        Debug.Print "Layer -> " & Layer
'
'        Select Case Layer
'            Case 0:
'                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
'                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
'                Call IncrementRecv
'                Call SendMasterKey(SocketSSL)
'            Case 1:
'                TheData = RC4_Decrypt(TheData)
'                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
'                    If VerifyMAC(TheData) Then Call SendClientFinish(SocketSSL)
'                Else
'                    SocketSSL.Close
'                End If
'             Case 2:
'                TheData = RC4_Decrypt(TheData)
'                If VerifyMAC(TheData) = False Then
'                    SocketSSL.Close
'                Else
'' Added by Giorock
''-----------------------------------------------------------'
'                    State = MailFrom
'                    SSLSend SocketSSL, "AUTH LOGIN" + vbCrLf
''-----------------------------------------------------------'
'                End If
'                Layer = 3
'             Case 3
'                TheData = RC4_Decrypt(TheData)
'                If VerifyMAC(TheData) Then
'                    Call ProcessData(Mid(TheData, 17), SocketSSL)
'                End If
'        End Select
'
'        SeekLen = 0
'
'    Loop Until bytesTotal = 0
'
'End Sub
'
'Private Sub SocketSSL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    MsgBox Number & " " & Description, vbCritical
'    SocketSSL_Close
'End Sub
'
