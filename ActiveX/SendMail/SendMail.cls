VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private WithEvents sckMail      As Winsock          ' project must include the Winsock control
Attribute sckMail.VB_VarHelpID = -1
Private WithEvents SocketSSL      As Winsock          ' project must include the Winsock control
Attribute SocketSSL.VB_VarHelpID = -1


'Variables for Parsing
Private Layer As Integer
Private InBuffer As String
Private Processing As Boolean
Private SeekLen As Integer

Dim RSA_KEY_LEN As Long

Private RTBX As Object 'RichTextBox
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_CONNECTION_LAN = &H2 'Local system uses a LAN to connect to the Internet.
Private Const INTERNET_CONNECTION_MODEM = &H1 'Local system uses a modem to connect to the Internet.
Private Const INTERNET_CONNECTION_PROXY = &H4 'Local system uses a proxy server to connect to the Internet.
Private Const INTERNET_RAS_INSTALLED = &H10 'Local system has RAS installed.
'Declaration Used For InternetDialUp.
Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Const INTERNET_DIAL_SHOW_OFFLINE = &H4000
Private Const INTERNET_DIAL_FORCE_PROMPT = &H2000
''SENDING VARIABLES
Private gstrServerName As String
Private glngPort As Long
Private gblnSSL As Boolean
Private gstrDomainName As String
Private gstrUserID As String
Private gstrUserPass As String
Private gintSent As Integer
'Server Attributes
Private Private_KEY As String
Private ENCODED_CERT As String
Private CONNECTION_ID As String
'Hand Shake Variables
Private CLIENT_HELLO As String
Private CHALLENGE_DATA As String
Private PUBLIC_KEY As String

Private gintsentTo As Integer
Private gintsentCC As Integer
Private pintSend As Integer
Private pintSent As Integer

Private Enum EState
    Connect2
    helo
    MailFrom
    sendto
    Data
    MessageData
    EndMessage
End Enum
Private State As EState
'CryptoAPI Constants
Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const ALG_CLASS_HASH = (4 * 2 ^ 13)
Private Const ALG_SID_MD5 = 3
Private Const ALG_TYPE_ANY = 0
Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const ALG_CLASS_DATA_ENCRYPT = (3 * 2 ^ 13)
Private Const ALG_SID_RC4 = 1
Private Const ALG_TYPE_STREAM = (4 * 2 ^ 9)
Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
Private Const HP_HASHVAL As Long = 2
Private Const SIMPLEBLOB As Long = 1
Private Const GEN_KEY_BITS As Long = &H800000



Private m_Index1(0 To 255) As Byte
Private m_Index2(0 To 255) As Byte
Private m_Index3(0 To 255) As Byte
Private m_Index4(0 To 63) As Byte
Private m_ReverseIndex1(0 To 255) As Byte
Private m_ReverseIndex2(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex3(0 To 255, 0 To 1) As Byte
Private m_ReverseIndex4(0 To 255) As Byte


Private Const MAX_LINELENGTH As Long = 76 ' Must be a multiple of 4
Private Const CHAR_EQUAL As Byte = 61
Private Const CHAR_CR As Byte = 13
Private Const CHAR_LF As Byte = 10

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal lpsBuffer As String, ByVal dwNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, ByVal lpszConnectoid As String, ByVal dwFlags As Long, ByRef lpdwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal lpdwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (dst As Any, src As Any, ByVal cb As Long)

Public Event DecodeProgress(ByVal Percent As Single, ByVal Total As Long)
Public Event EncodeProgress(ByVal Percent As Single, ByVal Total As Long)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Type FileInfo
    Filename As String
    Code As String
    Send As Boolean
End Type
Private FileCode() As FileInfo
Private Const conBoundary = "SendNextMIME_121_32"

Private Enum DecodeType
    base64
    uu
End Enum

Private Const PortionSize As Integer = 45
Private Const BufLen As Integer = 1024


' API Constants
Private Const REG_SZ = 1&
Private Const ERROR_SUCCESS     As Long = 0
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const WS_VERSION_REQD   As Long = &H101
Private Const MIN_SOCKETS_REQD  As Long = 1
Private Const DATA_SIZE = 32
Private Const MAX_WSAD = 256
Private Const MAX_WSAS = 128
Private Const PING_TIMEOUT = 255

Private Const TIME_ZONE_ID_UNKNOWN  As Long = 1
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Private Const TIME_ZONE_ID_INVALID  As Long = &HFFFFFFFF

' Winsock API Type defs...
Private Type ICMP_OPTIONS
    Ttl                         As Byte
    Tos                         As Byte
    Flags                       As Byte
    OptionsSize                 As Byte
    OptionsData                 As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address                     As Long
    Status                      As Long
    RoundTripTime               As Long
    DataSize                    As Long
    DataPointer                 As Long
    options                     As ICMP_OPTIONS
    Data                        As String * 250
End Type

Private Type HostEnt
    hName                       As Long
    hAliases                    As Long
    hAddrType                   As Integer
    hLen                        As Integer
    hAddrList                   As Long
End Type

Private Type WSADATA
    wVersion                    As Integer
    wHighVersion                As Integer
    szDescription(MAX_WSAD)     As Byte
    szSystemStatus(MAX_WSAS)    As Byte
    wMaxSockets                 As Integer
    wMaxUDPDG                   As Integer
    dwVendorInfo                As Long
End Type

' SystemTime and TimeZone API Type defs...
Private Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias                        As Long
    StandardName(63)            As Byte
    StandardDate                As SYSTEMTIME
    StandardBias                As Long
    DaylightName(63)            As Byte
    DaylightDate                As SYSTEMTIME
    DaylightBias                As Long
End Type

' Class Enum for host name string validation
Public Enum VALIDATE_HOST_METHOD
    VALIDATE_HOST_NONE = 0
    VALIDATE_HOST_SYNTAX = 1
    VALIDATE_HOST_PING = 2
    VALIDATE_HOST_DNS = 3
End Enum

' Class Enum for email address string validation
Public Enum VALIDATE_METHOD
    VALIDATE_NONE = 0
    VALIDATE_SYNTAX = 1
End Enum

' Class Enum for email encoding method
Public Enum ENCODE_METHOD
    MIME_ENCODE = 0
    UU_ENCODE = 1
End Enum

' Class Enum for mail priority
Public Enum MAIL_PRIORITY
    HIGH_PRIORITY = 1
    NORMAL_PRIORITY = 3
    LOW_PRIORITY = 5
End Enum

' Structure to hold mail elements
Private Type MAIL_DATA
    sToAddr()                   As String           ' To: email address
    sToDisplayName()            As String           ' To: display name
    sCcAddr()                   As String           ' Cc: email address
    sCcDisplayName()            As String           ' Cc: display name
    sBccAddr()                  As String           ' Bcc: email address
    sFromAddr                   As String           ' From: email address
    sFromDisplayName            As String           ' From: display name
    sReplyToAddr                As String           ' ReplyTo: email address
    sSubject                    As String           ' Subject
    sMailMessage                As String           ' email message body
    sAttachment()               As String           ' attachment path\filename
    sAttachNameOnly()           As String           ' attachment name only
    bAttachCID()                As Boolean          ' attachment has an assigned CID in an HTML document
    lAttachNameSize             As Long             ' sum of the lenght of all attachment names
    lAttachFileSize             As Long             ' sum of all file lenghts
    lAttachCount                As Long             ' number of attachments
End Type

' Class Property var's
Private utMail                  As MAIL_DATA        ' see above type def
Private etPriority              As MAIL_PRIORITY    ' mail priority, Normal - High - Low
Private psDelimiter             As String           ' string to delimit multiple entries
Private psSMTPHost              As String           ' remote host name or IP number
Private plSMTPPort              As Long             ' remote host port number
Private pbUseAuthentication     As Boolean          ' flag, use login authentication with host
Private pbHtmlText              As Boolean          ' flag, send plain text / html text
Private psContentBase           As String           ' Content base for HTML text
Private plConnectTimeout        As Long             ' timeout value for connection attempts
Private plConnectRetry          As Long             ' number of times to attempt a connection
Private plMessageTimeOut        As Long             ' timeout value for sending a message
Private pbPersistentSettings    As Long             ' flag, persistent/non-persistent settings
Private etEncodeType            As ENCODE_METHOD    ' MIME / UUEncode flag
Private etEmailValidation       As VALIDATE_METHOD  ' type of email address validation to use
Private etSMTPHostValidation    As VALIDATE_METHOD  ' type of Host validation to use
Private pbReceipt               As Boolean          ' flag, request a return receipt
Private plMaxRecipients         As Long             ' maximun recipient count before raising an error

' Class local var's
Private psTimeZoneBias          As String           ' time zone offset bias
Private pColErrors              As Collection       ' errors collection
Private pbBase64Byt(0 To 63)    As Byte             ' base 64 encoder byte array
Private psUUEncodeChr(0 To 63)  As String           ' UU encoder string array
Private pb8BitMail              As Boolean          ' flag, 7/8 bit message body
Private pbExitImmediately       As Boolean          ' flag - unrecoverable error
Private pbConnected             As Boolean          ' flag, connection to host established
Private pbManualDisconnect      As Boolean          ' flag, stay connected until 'Disconnect' called
Private pbRequestAccepted       As Boolean          ' flag, host accepted request
Private pbDataOK                As Boolean          ' flag, received "OK" from host
Private pbAuthLoginSupported    As Boolean          ' flag, host supports auth login
Private pbAuthMailFromOK        As Boolean          ' flag, host accepts 'mail from' auth
Private pbAuthLoginSuccess      As Boolean          ' flag, Auth login accepted by remote host
Private plBytesSent             As Long             ' running total of bytes sent
Private plBytesRemaining        As Long             ' bytes remaining to be sent in sock send buffer
Private pbSendProgress          As Boolean          ' flag indicating that the send progress event has fired
Private plMailSize              As Long             ' total size of email session
Private psUserName              As String           ' Auth username - optional, not supported by all servers
Private psPassword              As String           ' Auth password - optional, not supported by all servers
Private psPriority              As String           ' string version of priority Property for MSMail
Private plPop3Status            As Long             ' POP3 connection status
Private pbUsePopAuthentication  As Boolean          ' server requires Pop authorization (before SMTP)
Private pbPopAuthOk             As Boolean          ' POP3 auth OK
Private psPop3Host              As String           ' POP3 server name
                                                    ' or a reference to the mswinsck.ocx
Private psDay()                 As String           ' day name array
Private psMonth()               As String           ' month name array


' Class Constants

' base 64 encoder string
Private Const BASE64CHR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

' error strings used with 'pColErrors' collection to report errors to the user
Private Const ERR_INVALID_HOST = "Invalid or Missing SMTP Host Name"
Private Const ERR_INVALID_POP_HOST = "Invalid or Missing POP3 Host Name"
Private Const ERR_INVALID_PORT = "Invalid Remote Port"
Private Const ERR_INVALID_REC_EMAIL = "Missing or Invalid Recipient E-mail Address"
Private Const ERR_NO_REC_EMAIL = "No Recipient E-mail Address Specified"
Private Const ERR_INVALID_CC_EMAIL = "Invalid Cc: Recipient E-mail Address"
Private Const ERR_INVALID_BCC_EMAIL = "Invalid Bcc: Recipient E-mail Address"
Private Const ERR_INVALID_SND_EMAIL = "Missing or Invalid Sender E-mail Address"
Private Const ERR_TIMEOUT = "Timeout occurred: The SMTP Host did not respond to the request"
Private Const ERR_FILE_NOT_EXIST = "The file you tried to attach does not exist"
Private Const ERR_RECIPIENT_COUNT = "Too many recipients"
Private Const ERR_HTML_REQUIRES_MIME = "Sending HTML requires MIME encoding"

' misc startup defaults
Private Const CONNECT_TIMEOUT = 40                  ' seconds to wait before giving up
Private Const CONNECT_RETRY = 4                     ' number of times to try before giving up
Private Const MSG_TIMEOUT = 60                      ' seconds before timing out on message transmission
Private Const REG_KEY = "vbSendMail"                ' registry key
Private Const SETTINGS_KEY = "Settings"             ' registry sub key
Private Const DEFAULT_PORT As Long = 25             ' default socket port for SMTP
Private Const POP3_PORT As Long = 110               ' default socket port for POP3

Private Const Q_CODE_HDR    As String = "=?ISO-8859-1?Q?"
Private Const B_CODE_HDR    As String = "=?ISO-8859-1?B?"
Private Const CODE_END      As String = "?="
Private Const CHAR_SET      As String = "iso-8859-1"


' maximums per RFC 821...
Private Const MAX_TEXTLINE_LEN = 1000               ' maximum total lenght of a text line
Private Const MAX_RECIPIENTS = 100                  ' maximum number of recipients that must be buffered

' list of top level Domains, obtained from www.IANA.com.
' Can and will change, used in host name syntax checking
Private Const TOP_DOMAINS = "AERO BIZ COM COOP EDU GOV INFO INT MIL MUSEUM NAME NET ORG PRO " & _
        "AC AD AE AF AG AI AL AM AN AO AQ AR AS AT AU AW AZ BA BB BD " & _
        "BE BF BG BH BI BJ BM BN BO BR BS BT BV BW BY BZ CA CC CD CF " & _
        "CG CH CI CK CL CM CN CO CR CU CV CX CY CZ DE DJ DK DM DO DZ " & _
        "EC EE EG EH ER ES ET FI FJ FK FM FO FR GA GD GE GF GG GH GI " & _
        "GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL " & _
        "IM IN IO IQ IR IS IT JE JM JO JP KE KG KH KI KM KN KP KR KW " & _
        "KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD MG MH MK ML " & _
        "MM MN MO MP MQ MR MS MT MU MV MW MX MY MZ NA NC NE NF NG NI " & _
        "NL NO NP NR NU NZ OM PA PE PF PG PH PK PL PM PN PR PS PT PW " & _
        "PY QA RE RO RU RW SA SB SC SD SE SG SH SI SJ SK SL SM SN SO " & _
        "SR ST SV SY SZ TC TD TF TG TH TJ TK TM TN TO TP TR TT TV TW " & _
        "TZ UA UG UK UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT YU " & _
        "ZA ZM ZW"

' Class Events
Public Event SendSuccesful()
Public Event SendFailed(Explanation As String)
Public Event Status(Status As String)
Public Event Progress(PercentComplete As Long)

' API prototypes...
' winsock
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function GetHostName Lib "wsock32.dll" Alias "gethostname" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long

' registry
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal RESERVED As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
        
' misc
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'CryptoAPI Functions
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long

'CryptoAPI Constants
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

'Encryption Keys
Private MASTER_KEY As String
Private CLIENT_READ_KEY As String
Private CLIENT_WRITE_KEY As String

'Counters
Private SEND_SEQUENCE_NUMBER As Double
Private RECV_SEQUENCE_NUMBER As Double

''MESSAGE VARIABLES
Private gblnBCCMode As Boolean
Private gstrData() As String

'Class Variables
Private hCryptProv As Long
Private hClientWriteKey As Long
Private hClientReadKey As Long
Private hMasterKey As Long
'Private lngType As Long

Private gstrFromName As String
Private gstrSubject As String
Private gstrMessage As String
Private gstrToAddress() As String
Private gstrCC() As String
Private gstrBCC() As String
Private gblnDone As Boolean
Private gstrAllAddresses() As String
'Private gblnBCCMode As Boolean
Private gintTransfer As Integer
Private gstrFileNames() As String
Private gstrFromAddress As String
Private gblnAttachments As Boolean
'Private conBoundary As String
Private gstrMessageFileName  As String

Private mvarSendSuccesful  As Boolean
Private mvarStrStatus      As String
Private mvarSSL            As Boolean
Private ProgBar As ProgressBar
Public Property Let SSL(ByVal vData As Boolean)
    mvarSSL = vData
End Property
Public Property Get SSL() As Boolean
    SSL = mvarSSL
End Property
Public Property Let StrStatus(ByVal vData As String)
    mvarStrStatus = vData
End Property
Public Property Get StrStatus() As String
    StrStatus = mvarStrStatus
End Property
Public Property Let SendSuccesful(ByVal vData As Boolean)
    mvarSendSuccesful = vData
End Property
Public Property Get SendSuccesful() As Boolean
    SendSuccesful = mvarSendSuccesful
End Property

Private Sub Class_Initialize()

    Dim iPtr        As Integer                      ' loop counter
    Dim utTZ        As TIME_ZONE_INFORMATION        ' api time zone type
    Dim dwBias      As Long
    
    ' instantiate the Error collection
    Set pColErrors = New Collection
    
    ' instantiate the Winsock Control
    Set sckMail = frmSck.WinSck
    
    ' alternate method of instantiating Winsock without a Form.
    ' use a project Reference instead of the included frmSck & Winsock control
    ' *** currently has unresolved deployment issues ***
    'Set sckMail = New Winsock
    
    ' initialize default values...
    pbPersistentSettings = CLng(RegGet("PersistentSettings", "0"))
    If pbPersistentSettings Then
        ' load defaults from the registry
        utMail.sFromAddr = RegGet("From", "")
        utMail.sFromDisplayName = RegGet("FromDisplayName", "")
        psPop3Host = RegGet("Pop3Host", "")
        psSMTPHost = RegGet("RemoteHost", "")
        plSMTPPort = CLng(RegGet("RemotePort", DEFAULT_PORT))
        etSMTPHostValidation = RegGet("SMTPHostValidation", VALIDATE_HOST_DNS)
        etEmailValidation = CLng(RegGet("EmailValidation", VALIDATE_SYNTAX))
        plConnectTimeout = CLng(RegGet("ConnectTimeout", CONNECT_TIMEOUT))
        plMessageTimeOut = CLng(RegGet("MessageTimeout", MSG_TIMEOUT))
        plConnectRetry = CLng(RegGet("ConnectRetry", CONNECT_RETRY))
        etEncodeType = RegGet("EncodeType", MIME_ENCODE)
        psUserName = RegGet("Username", "")
        pbUseAuthentication = RegGet("UseAuthentication", False)
        pbUsePopAuthentication = RegGet("UsePopAuthentication", False)
        plMaxRecipients = CLng(RegGet("MaxRecipients", MAX_RECIPIENTS))
        
        plConnectTimeout = IIf(plConnectTimeout < 3, 3, plConnectTimeout)
        plMessageTimeOut = IIf(plMessageTimeOut < 3, 3, plMessageTimeOut)
    Else
        ' load standard defaults
        plSMTPPort = DEFAULT_PORT
        etSMTPHostValidation = VALIDATE_HOST_DNS
        etEmailValidation = VALIDATE_SYNTAX
        plConnectTimeout = CONNECT_TIMEOUT
        plMessageTimeOut = MSG_TIMEOUT
        plConnectRetry = CONNECT_RETRY
        etEncodeType = MIME_ENCODE
        pbHtmlText = False
        plMaxRecipients = MAX_RECIPIENTS
    End If

    If plMaxRecipients <= 0 Then plMaxRecipients = MAX_RECIPIENTS

    ' initialize the arrays for base64 & uu encoders
    For iPtr = 0 To 63
        pbBase64Byt(iPtr) = Asc(Mid$(BASE64CHR, iPtr + 1, 1))
        psUUEncodeChr(iPtr) = Chr$(iPtr + &H20)
    Next
    psUUEncodeChr(0) = Chr$(&H60)

    ' calculate the time zone offset bias
    Select Case GetTimeZoneInformation(utTZ)
        Case TIME_ZONE_ID_DAYLIGHT:
            dwBias = utTZ.Bias + utTZ.DaylightBias
        Case Else
            dwBias = utTZ.Bias + utTZ.StandardBias
    End Select
    psTimeZoneBias = Format$(-dwBias \ 60, "00") & Format$(Abs(dwBias - (dwBias \ 60) * 60), "00")
    If InStr(psTimeZoneBias, "-") = 0 Then psTimeZoneBias = "+" & psTimeZoneBias
    
    ' init mail recipient arrays (sets Ubound to -1)
    utMail.sToAddr = Split("")
    utMail.sToDisplayName = utMail.sToAddr
    utMail.sCcAddr = utMail.sToAddr
    utMail.sCcDisplayName = utMail.sToAddr
    utMail.sBccAddr = utMail.sToAddr
    utMail.sAttachment = utMail.sToAddr

    ' set default delimiter
    psDelimiter = ";"

    ' set default priority
    etPriority = NORMAL_PRIORITY

    ' initialize the day/month arrays needed to support non-English systems.
    ' some email clients/servers will not accept non-English words in the
    ' date field so we need to guarantee that the day & month are English.
    ' These arrays are used in the Send Sub to format the current time/date.
    psDay() = Split(",Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")
    psMonth() = Split(",Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ",")

   Call Class_Initialize2
End Sub

Private Sub Class_Terminate()

    ' make sure sckMail is closed
    If sckMail.State <> sckClosed Then
        DisconnectFromHost
    End If
    
    ' release memory
    Set sckMail = Nothing
    Set pColErrors = Nothing
   Class_Terminate2
End Sub

' ******************************************************************************
' *      Class Properties                                                      *
' ******************************************************************************

Public Property Get AsHTML() As Boolean

    ' return the Property value
    AsHTML = pbHtmlText

End Property

Public Property Let AsHTML(ByVal NewValue As Boolean)

    ' save the new Property value
    pbHtmlText = NewValue

End Property

Public Property Get Attachment() As String

    ' return the Property value
    Attachment = Join(utMail.sAttachment, psDelimiter)

End Property

Public Property Let Attachment(ByVal NewValue As String)

    Dim sNameOnly()     As String
    Dim lPtr            As Long

    ' save the new Property value
    utMail.sAttachment = Split(NewValue, psDelimiter)
    
    ' reset the counters
    utMail.lAttachCount = UBound(utMail.sAttachment) + 1
    utMail.lAttachFileSize = 0
    utMail.lAttachNameSize = 0

    RemoveError ERR_FILE_NOT_EXIST

    ' process all of the file names
    If utMail.lAttachCount Then
        ReDim utMail.sAttachNameOnly(utMail.lAttachCount - 1)
        ReDim utMail.bAttachCID(utMail.lAttachCount - 1)
        ' verify each entry...
        For lPtr = 0 To UBound(utMail.sAttachment)
            ' check that the file exists
            If Dir(utMail.sAttachment(lPtr)) = "" Then
                AddError ERR_FILE_NOT_EXIST
                Exit For
            End If
            ' extract the file name
            sNameOnly = Split(utMail.sAttachment(lPtr), "\")
            utMail.sAttachNameOnly(lPtr) = sNameOnly(UBound(sNameOnly))
            ' add up the file sizes and name lengths for later...
            utMail.lAttachFileSize = utMail.lAttachFileSize + FileLen(utMail.sAttachment(lPtr))
            utMail.lAttachNameSize = utMail.lAttachNameSize + Len(utMail.sAttachNameOnly(lPtr))
        Next
    End If
    
    Dir (App.Path)
    
End Property
Public Property Get BCcRecipient() As String
    ' return the Property value
    BCcRecipient = Join(utMail.sBccAddr, psDelimiter)
End Property
Public Property Let BCcRecipient(ByVal NewValue As String)
    ' save the new Property value
    utMail.sBccAddr = Split(NewValue, psDelimiter)
    TrimWhiteSpace utMail.sBccAddr
    ValidateAddress NewValue, ERR_INVALID_BCC_EMAIL
End Property
Public Property Get CcDisplayName() As String
    ' return the Property value
    CcDisplayName = Join(utMail.sCcDisplayName, psDelimiter)
End Property
Public Property Let CcDisplayName(ByVal NewValue As String)
    ' save the new Property value
    utMail.sCcDisplayName = Split(NewValue, psDelimiter)
    TrimWhiteSpace utMail.sCcDisplayName
End Property
Public Property Get CcRecipient() As String
    ' return the Property value
    CcRecipient = Join(utMail.sCcAddr, psDelimiter)
End Property
Public Property Let CcRecipient(ByVal NewValue As String)
    ' save the new Property value
    utMail.sCcAddr = Split(NewValue, psDelimiter)
    TrimWhiteSpace utMail.sCcAddr
    ValidateAddress NewValue, ERR_INVALID_CC_EMAIL
End Property

Public Property Get ConnectRetry() As Long

    ' return the Property value
    ConnectRetry = plConnectRetry

End Property

Public Property Let ConnectRetry(ByVal NewValue As Long)

    ' save the new Property value
    If NewValue > 0 And NewValue <= 20 Then plConnectRetry = NewValue
    RegSave "ConnectRetry", Str$(NewValue)

End Property

Public Property Get ConnectTimeout() As Long

    ' return the Property value
    ConnectTimeout = plConnectTimeout

End Property

Public Property Let ConnectTimeout(ByVal NewValue As Long)

    ' save the new Property value
    If NewValue > 0 And NewValue <= 120 Then plConnectTimeout = NewValue
    RegSave "ConnectTimeout", Str$(NewValue)

End Property

Public Property Get ContentBase() As String

    ' return the Property value
    ContentBase = psContentBase

End Property

Public Property Let ContentBase(ByVal NewValue As String)

    ' save the new Property value
    ' fix some common mistakes...
    If Len(NewValue) Then
        Replace$ NewValue, "\", "/"
        If InStr(1, NewValue, "http://", vbTextCompare) = 0 Then NewValue = "http://" & NewValue
        If Right$(NewValue, 1) <> "/" Then NewValue = NewValue & "/"
    End If
    psContentBase = NewValue

End Property

Public Property Get Delimiter() As String

    ' return the Property value
    Delimiter = psDelimiter

End Property

Public Property Let Delimiter(ByVal NewValue As String)

    ' save the new Property value
    psDelimiter = Left$(NewValue, 1)

End Property

Public Property Get EmailAddressValidation() As VALIDATE_METHOD

    ' return the Property value
    EmailAddressValidation = etEmailValidation

End Property

Public Property Let EmailAddressValidation(ByVal NewValue As VALIDATE_METHOD)

    ' save the new Property value
    etEmailValidation = NewValue
    RegSave "EmailValidation", Str$(NewValue)

End Property

Public Property Get EncodeType() As ENCODE_METHOD

    ' return the Property value
    EncodeType = etEncodeType

End Property

Public Property Let EncodeType(ByVal NewValue As ENCODE_METHOD)
    
    ' save the new Property value
    etEncodeType = NewValue
    RegSave "EncodeType", Str$(NewValue)

End Property

Public Property Get from() As String

    ' return the Property value
    from = utMail.sFromAddr

End Property

Public Property Let from(ByVal NewValue As String)
    
    ' save the new Property value
    utMail.sFromAddr = Trim$(NewValue)
    ValidateAddress NewValue, ERR_INVALID_SND_EMAIL
    RegSave "From", NewValue

End Property

Public Property Get FromDisplayName() As String

    ' return the Property value
    FromDisplayName = utMail.sFromDisplayName

End Property

Public Property Let FromDisplayName(ByVal NewValue As String)
    
    ' save the new Property value
    utMail.sFromDisplayName = Trim$(NewValue)
    RegSave "FromDisplayName", NewValue

End Property


Public Property Get MaxRecipients() As Long

    ' return the Property value
    MaxRecipients = plMaxRecipients

End Property

Public Property Let MaxRecipients(ByVal NewValue As Long)
    
    ' save the new Property value
    plMaxRecipients = Abs(NewValue)
    RegSave "MaxRecipients", CStr(plMaxRecipients)

End Property


Public Property Get Message() As String

    ' return the Property value
    Message = utMail.sMailMessage

End Property

Public Property Let Message(ByVal NewValue As String)

    Dim lPtr        As Long
    Dim bytTmp()    As Byte

    ' save the new Property value
    utMail.sMailMessage = FormatMail(NewValue)

    ' check for any 8 bit characters
    pb8BitMail = False
    bytTmp() = StrConv(utMail.sMailMessage, vbFromUnicode)

    For lPtr = 0 To UBound(bytTmp)
        If bytTmp(lPtr) > 126 Then
            pb8BitMail = True
            Exit For
        End If
    Next

End Property

Public Property Get MessageTimeout() As Long

    ' return the Property value
    MessageTimeout = plMessageTimeOut

End Property

Public Property Let MessageTimeout(ByVal NewValue As Long)

    ' save the new Property value
    plMessageTimeOut = Abs(NewValue)
    RegSave "MessageTimeout", Str$(NewValue)

End Property

Public Property Get Password() As String

    ' return the Property value
    Password = psPassword

End Property

Public Property Let Password(ByVal NewValue As String)

    ' save the new Property value
    psPassword = NewValue

End Property

Public Property Get PersistentSettings() As Boolean

    ' return the Property value
    PersistentSettings = pbPersistentSettings

End Property

Public Property Let PersistentSettings(ByVal NewValue As Boolean)

    ' save the new Property value
    pbPersistentSettings = NewValue
    RegSave "PersistentSettings", CStr(CLng(NewValue))

End Property

Public Property Get Priority() As MAIL_PRIORITY

    ' return the Property value
    Priority = etPriority

End Property

Public Property Let Priority(ByVal NewValue As MAIL_PRIORITY)

    ' save the new Property value
    etPriority = NewValue

    ' set the string version to match
    Select Case etPriority
        
        Case NORMAL_PRIORITY
            psPriority = "Normal"
        
        Case HIGH_PRIORITY
            psPriority = "High"
        
        Case LOW_PRIORITY
            psPriority = "Low"
    
    End Select

End Property

Public Property Get Receipt() As Boolean

    ' return the Property value
    Receipt = pbReceipt

End Property

Public Property Let Receipt(ByVal NewValue As Boolean)

    ' save the new Property value
    pbReceipt = NewValue

End Property

Public Property Get Recipient() As String

    ' return the Property value
    Recipient = Join(utMail.sToAddr, psDelimiter)

End Property

Public Property Let Recipient(ByVal NewValue As String)
    
    ' save the new Property value
    utMail.sToAddr = Split(NewValue, psDelimiter)
    TrimWhiteSpace utMail.sToAddr
    ValidateAddress NewValue, ERR_INVALID_REC_EMAIL

End Property

Public Property Get RecipientDisplayName() As String

    ' return the Property value
    RecipientDisplayName = Join(utMail.sToDisplayName, psDelimiter)

End Property

Public Property Let RecipientDisplayName(ByVal NewValue As String)

    ' save the new Property value
    utMail.sToDisplayName = Split(NewValue, psDelimiter)
    TrimWhiteSpace utMail.sToDisplayName

End Property

Public Property Get ReplyToAddress() As String

    ' return the Property value
    ReplyToAddress = utMail.sReplyToAddr

End Property

Public Property Let ReplyToAddress(ByVal NewValue As String)

    ' save the new Property value
    utMail.sReplyToAddr = Trim$(NewValue)

End Property

Public Property Get POP3Host() As String

    ' return the Property value
    POP3Host = psPop3Host

End Property

Public Property Let POP3Host(NewValue As String)

    Dim bValid      As Boolean
    
    NewValue = Replace(NewValue, vbNullChar, vbNullString)
    
    ' validate the new host name...
    If Len(NewValue) Then
        Select Case etSMTPHostValidation
        
            Case VALIDATE_HOST_SYNTAX
                bValid = IsValidIPHost(NewValue)
        
            Case VALIDATE_HOST_PING
                bValid = Ping(NewValue)
        
            Case VALIDATE_HOST_DNS
                If GetIPAddress(NewValue) <> "" Then bValid = True
        
            Case Else
                bValid = True
    
        End Select
    Else
        bValid = True
    End If

    ' save the new Property value
    If bValid Then
        RegSave "Pop3Host", NewValue
        RemoveError ERR_INVALID_POP_HOST
        psPop3Host = NewValue
    Else
        AddError ERR_INVALID_POP_HOST
    End If

End Property

Public Property Get SMTPHost() As String

    ' return the Property value
    SMTPHost = psSMTPHost

End Property

Public Property Let SMTPHost(NewValue As String)

    Dim bValid      As Boolean
    
    NewValue = Replace(NewValue, vbNullChar, vbNullString)
    
    ' validate the new host name...
    If Len(NewValue) Then
        Select Case etSMTPHostValidation
        
            Case VALIDATE_HOST_SYNTAX
                bValid = IsValidIPHost(NewValue)
        
            Case VALIDATE_HOST_PING
                bValid = Ping(NewValue)
        
            Case VALIDATE_HOST_DNS
                If GetIPAddress(NewValue) <> "" Then bValid = True
        
            Case Else
                bValid = True
    
        End Select
    Else
        bValid = True
    End If

    ' save the new Property value
    If bValid Then
        RegSave "RemoteHost", NewValue
        RemoveError ERR_INVALID_HOST
        psSMTPHost = NewValue
    Else
        AddError ERR_INVALID_HOST
    End If

End Property

Public Property Get SMTPHostValidation() As VALIDATE_HOST_METHOD

    ' return the Property value
    SMTPHostValidation = etSMTPHostValidation

End Property

Public Property Let SMTPHostValidation(ByVal NewValue As VALIDATE_HOST_METHOD)

    ' save the new Property value
    etSMTPHostValidation = NewValue
    RegSave "SMTPHostValidation", Str$(NewValue)

    ' in case this is set after the host value is set
    If psSMTPHost <> "" Then SMTPHost = psSMTPHost

End Property

Public Property Get SMTPPort() As Long

    ' return the Property value
    SMTPPort = plSMTPPort

End Property

Public Property Let SMTPPort(ByVal NewValue As Long)

    ' save the new Property value
    If NewValue < 1 Or NewValue > 65535 Then
        AddError ERR_INVALID_PORT
    Else
        plSMTPPort = NewValue
        RegSave "RemotePort", Str$(NewValue)
        RemoveError ERR_INVALID_PORT
    End If

End Property

Public Property Get Subject() As String

    ' return the Property value
    Subject = utMail.sSubject

End Property

Public Property Let Subject(ByVal NewValue As String)

    ' save the new Property value
    utMail.sSubject = NewValue

End Property

Public Property Get UseAuthentication() As Boolean

    ' return the Property value
    UseAuthentication = pbUseAuthentication

End Property

Public Property Let UseAuthentication(ByVal NewValue As Boolean)

    ' save the new Property value
    pbUseAuthentication = NewValue
    RegSave "UseAuthentication", CStr(CLng(NewValue))

End Property

Public Property Get UsePopAuthentication() As Boolean

    ' return the Property value
    UsePopAuthentication = pbUsePopAuthentication

End Property

Public Property Let UsePopAuthentication(ByVal NewValue As Boolean)

    ' save the new Property value
    pbUsePopAuthentication = NewValue
    RegSave "UsePopAuthentication", CStr(CLng(NewValue))

End Property

Public Property Get Username() As String

    ' return the Property value
    Username = psUserName

End Property

Public Property Let Username(ByVal NewValue As String)

    ' save the new Property value
    psUserName = NewValue
    RegSave "Username", NewValue

End Property

' ******************************************************************************
' *      Class Methods                                                         *
' ******************************************************************************

Public Function Connect() As Boolean

    ' public version of ConnectToHost
    ' sets pbManualDisconnect flag so Send Sub
    ' will not disconnect when finished....
    pbManualDisconnect = True
    Connect = ConnectToHost

End Function

Public Sub Disconnect()

    ' public version of DisconnectFromHost
    ' clears pbManualDisconnect flag
    pbManualDisconnect = False
    DisconnectFromHost

End Sub

Public Function GetContentType(ByVal strFile As String, Optional strDefault As String = "application/octet-stream") As String

' ******************************************************************************
'
' Synopsis:     Get the Content Type from the Registry.
'
' Parameters:   strFile     - The filename to get the Content Type for
'               strDefault  - The default data to return if nothing is found
'
' Return:       The Content Type string
'
' Description:
' The Content Type string for registered file extensions is located in
' the system registry, in the root key HKEY_CLASSES_ROOT. Open the registry
' key for the given file extension and read the 'Content Type' value. If the
' key and/or value are not found, assign a default value of
' 'application/octet-stream'
'
' ******************************************************************************

    Dim hKey                As Long                 ' key handle
    Dim strBuff             As String               ' buffer for API to write to
    Dim lBuffLen            As Long                 ' lenght of API return string
    Dim lRet                As Long                 ' API return code
    Dim lValueType          As Long                 ' data type for retun value
    Dim iPtr                As Integer              ' scratch pointer
    Dim strValueName        As String               ' registry 'value name
    Dim strKeyName          As String               ' registry 'key name

    If bInEXE Then On Local Error GoTo ERR_GetContentType
    
    GetContentType = strDefault

    ' registry value name
    strValueName = "Content Type"

    ' get the passed in key name. We only want
    ' the file extension here e.g. .exe, .doc, etc.
    ' if an extension is not found, assign default
    ' value and return
    iPtr = InStrRev(strFile, ".")
    If iPtr Then
        strKeyName = Mid$(strFile, iPtr)
    Else
        Exit Function
    End If

    ' open the Registry key, if key not found, return the defaut value
    lRet = RegOpenKey(HKEY_CLASSES_ROOT, strKeyName, hKey)
    If lRet <> ERROR_SUCCESS Then Exit Function

    ' query the key value to get it's data type & length
    lRet = RegQueryValueEx(hKey, strValueName, 0&, lValueType, 0&, lBuffLen)

    ' should be type string...
    If lValueType = REG_SZ Then
        ' create a buffer & call the API again
        strBuff = String$(lBuffLen, " ")
        lRet = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal strBuff, lBuffLen)
        ' get the string value, drop the trailing '0'
        If lRet = ERROR_SUCCESS Then GetContentType = Left$(strBuff, lBuffLen - 1)
    End If

    ' close the key
    If hKey Then lRet = RegCloseKey(hKey)

    Exit Function

ERR_GetContentType:

    If hKey Then lRet = RegCloseKey(hKey)
    GetContentType = strDefault

End Function

Public Function GetIPAddress(sHostName As String) As String

    ' Resolves host-name to an IP address (DNS)
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim lpHost          As Long
    Dim HOST            As HostEnt
    Dim dwIPAddr        As Long
    Dim tmpIPAddr()     As Byte
    Dim i               As Integer
    Dim sIPAddr         As String
    
    ' init winsock api
    If Not SocketsInitialize() Then
        GetIPAddress = ""
        Exit Function
    End If
    
    ' if no name given, use local host
    If sHostName = "" Then sHostName = GetIPHost
    sHostName = Trim$(sHostName) & Chr$(0)
    
    ' call api
    lpHost = gethostbyname(sHostName)

    If lpHost Then
        ' extract the data...
        CopyMemory HOST, ByVal lpHost, Len(HOST)
        CopyMemory dwIPAddr, ByVal HOST.hAddrList, 4
        ReDim tmpIPAddr(1 To HOST.hLen)
        CopyMemory tmpIPAddr(1), ByVal dwIPAddr, HOST.hLen

        ' convert format
        For i = 1 To HOST.hLen
            sIPAddr = sIPAddr & tmpIPAddr(i) & "."
        Next

        ' set the return value
        GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    
    Else
        WSAGetLastError
        GetIPAddress = ""
    End If
    
    SocketsCleanup

End Function

Public Function GetIPHost() As String

    ' Resolves the local host name
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim sHostName   As String
    Dim iPtr        As Integer

    ' create a buffer
    sHostName = String$(256, Chr$(0))

    ' init winsock api
    If Not SocketsInitialize() Then Exit Function

    ' get the loacal hosts name
    If GetHostName(sHostName, Len(sHostName)) = ERROR_SUCCESS Then
        iPtr = InStr(sHostName, Chr$(0))
        If iPtr > 1 Then GetIPHost = Mid$(sHostName, 1, iPtr - 1)
    End If

    SocketsCleanup

End Function

Public Function IsValidEmailAddress(AddressString As String)  ' As Boolean

    Dim sTmp()      As String

    ' assume failure
    IsValidEmailAddress = False

    ' sould have one "@"
    sTmp = Split(AddressString, "@")
    If UBound(sTmp) <> 1 Then Exit Function

    IsValidEmailAddress = IsValidIPHost(sTmp(1))

End Function

Public Function MXQuery(Optional IPDomain As String = "") As String

    Dim sDomain     As String
    
    ' return the best server found in an MX Query
  
    If bInEXE Then On Local Error GoTo Err_MXQuery
    
    sDomain = Trim$(IPDomain)
    
    If Len(sDomain) Then
        mvarStrStatus = mvarStrStatus & vbNewLine & "Performing MX Query, Domain: " & sDomain
        RaiseEvent Status("Performing MX Query, Domain: " & sDomain)
    Else
        mvarStrStatus = mvarStrStatus & vbNewLine & "Performing MX Query"
        RaiseEvent Status("Performing MX Query")
    End If
    
    MXQuery = MX_Query(sDomain)
    
    Exit Function
    
Err_MXQuery:
    
   MXQuery = vbNullString
   mvarStrStatus = mvarStrStatus & vbNewLine & Err.Description
   RaiseEvent Status(Err.Description)
    
End Function

Public Function Ping(Address As String, _
        Optional RoundTripTime As String = "", _
        Optional DataSize As String = "", _
        Optional DataMatch As Boolean = False) As Boolean

    ' Ping a remote host
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim ECHO            As ICMP_ECHO_REPLY
    Dim iPtr            As Integer
    Dim Dt              As String
    Dim sAddress        As String
    Dim hPort           As Long
    Dim lAddress        As Long
    Dim bytAddr(3)      As Byte
    
    If bInEXE Then On Local Error GoTo DPErr

    ' assume failure
    Ping = False

    ' if passed a name, get the IP address
    If Not IsDottedQuad(Address) Then
        sAddress = GetIPAddress(Address)
    Else
        sAddress = Address
    End If

    If sAddress = "" Then Exit Function

    If SocketsInitialize Then

        ' build string of random characters
        For iPtr = 1 To DATA_SIZE
            Dt = Dt & Chr$(Rnd() * 254 + 1)
        Next

        ' ping an ip address, passing the
        ' address and the ECHO structure
        lAddress = AddressStringToLong(sAddress)
        hPort = IcmpCreateFile()
        IcmpSendEcho hPort, lAddress, Dt, Len(Dt), 0, ECHO, Len(ECHO), PING_TIMEOUT
        IcmpCloseHandle hPort

        ' get the results from the ECHO structure
        RoundTripTime = ECHO.RoundTripTime
        CopyMemory bytAddr(0), ECHO.Address, 4
        Address = CStr(bytAddr(0)) & "." & _
                  CStr(bytAddr(1)) & "." & _
                  CStr(bytAddr(2)) & "." & _
                  CStr(bytAddr(3))

        DataSize = ECHO.DataSize & " bytes"

        iPtr = InStr(ECHO.Data, Chr$(0))
        If iPtr > 1 Then DataMatch = (Left$(ECHO.Data, iPtr - 1) = Dt)
        If ECHO.Status = 0 And ECHO.Address = lAddress Then Ping = True

        SocketsCleanup

    End If

    Exit Function

DPErr:

End Function

Public Sub Send()
    Dim sSenderName         As String
    Dim sToHeader           As String
    Dim sCcHeader           As String
    Dim iCtr                As Integer
    Dim sAuth               As String
    Dim sTxt                As String
    Dim strBoundry          As String
    Dim bMimeMultiPart      As Boolean
    Dim fStart              As Single
    Dim fTimeOut            As Single
    Dim lSendBuffSize       As Long
    Dim bRelatedLinks       As Boolean
    Dim st                  As SYSTEMTIME
    
   mvarSendSuccesful = False
   mvarStrStatus = ""

    
    ' general catch all error handler only
    ' works when running in stand alone EXE
    If bInEXE Then On Local Error GoTo Err_Send
    
    ' check for multipart MIME
    If etEncodeType = MIME_ENCODE And utMail.lAttachCount > 0 Then
        bMimeMultiPart = True
    Else
        bMimeMultiPart = False
    End If

    ' check sender
    If Len(utMail.sFromAddr) = 0 Then AddError ERR_INVALID_SND_EMAIL

    ' HTML & UU Encode are mutually exclusive
    If pbHtmlText = True And etEncodeType = UU_ENCODE Then AddError ERR_HTML_REQUIRES_MIME

    ' check recipient count
    If UBound(utMail.sToAddr) + UBound(utMail.sCcAddr) + UBound(utMail.sBccAddr) = -3 Then AddError ERR_NO_REC_EMAIL
    If UBound(utMail.sToAddr) + UBound(utMail.sCcAddr) + UBound(utMail.sBccAddr) + 3 > plMaxRecipients Then AddError ERR_RECIPIENT_COUNT

    ' resize the display name arrays to match the recipient arrays
    iCtr = UBound(utMail.sToAddr)
    If iCtr >= 0 Then ReDim Preserve utMail.sToDisplayName(iCtr)
    iCtr = UBound(utMail.sCcAddr)
    If iCtr >= 0 Then ReDim Preserve utMail.sCcDisplayName(iCtr)
    
    ' we won't try to send if there's already an error
    If pColErrors.Count > 0 Then
        SendFail
        Exit Sub
    End If

    ' get the Content-Location for any linked objects
    If utMail.lAttachCount Then bRelatedLinks = GetAttachCID

    ' get the mail size
    plMailSize = EstimateMailSize

    ' this flag gets set when a socket error occurs or the host cannot process an
    ' input command, see 'SendFail', 'sckMail_DataArrival' & 'WaitUntilTrue' Subs
    pbExitImmediately = False

    With sckMail

        ' if not already conected then connect to the remote host
        If .State <> sckConnected Then
            If Not ConnectToHost Then
               Exit Sub
            End If
        End If

        ' reset the progress counter
        plBytesSent = 0

        ' tell the host who the mail is 'From
        mvarStrStatus = mvarStrStatus & vbNewLine & "Sending Sender Information..."
        RaiseEvent Status("Sending Sender Information...")
        pbRequestAccepted = False
        If pbAuthMailFromOK Then sAuth = " AUTH=" & utMail.sFromAddr Else sAuth = vbNullString
        '.SendData "MAIL FROM: <" & utMail.sFromAddr & ">" & sAuth & vbCrLf
        Call SocketSend(sckMail, "MAIL FROM: <" & utMail.sFromAddr & ">" & sAuth & vbCrLf)
        
        
        If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
        If pbExitImmediately Then Exit Sub
        
        ' tell the host who the recipients are
        ' build the 'To:' header string 'sToHeader' too
        mvarStrStatus = mvarStrStatus & vbNewLine & "Sending Recipient Information..."
        RaiseEvent Status("Sending Recipient Information...")
        For iCtr = 0 To UBound(utMail.sToAddr)
            ' send the recipient address & wait for a reply
            pbRequestAccepted = False
            '.SendData "RCPT TO: <" & utMail.sToAddr(iCtr) & ">" & vbCrLf
            Call SocketSend(sckMail, "RCPT TO: <" & utMail.sToAddr(iCtr) & ">" & vbCrLf)
            If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
            If pbExitImmediately Then Exit Sub
            
            ' build the 'To:' header string for later...
            If Len(utMail.sToDisplayName(iCtr)) Then
                sToHeader = sToHeader & CText(utMail.sToDisplayName(iCtr), True)
            Else
                sToHeader = sToHeader & """" & Trim$(utMail.sToAddr(iCtr)) & """"
            End If
            sToHeader = sToHeader & " <" & utMail.sToAddr(iCtr) & ">"
            If iCtr < UBound(utMail.sToAddr) Then sToHeader = sToHeader & ", " & vbCrLf & vbTab
        Next

        ' send Cc: recipient addresses (just more 'RCPT TO' addresses)
        ' build the 'Cc:' header string too
        For iCtr = 0 To UBound(utMail.sCcAddr)
            ' send the recipient address & wait for a reply
            pbRequestAccepted = False
            '.SendData "RCPT TO: <" & utMail.sCcAddr(iCtr) & ">" & vbCrLf
            Call SocketSend(sckMail, "RCPT TO: <" & utMail.sCcAddr(iCtr) & ">" & vbCrLf)
            If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
            If pbExitImmediately Then Exit Sub
            
            ' build the 'Cc:' header string for later...
            If Len(utMail.sCcDisplayName(iCtr)) Then
                sCcHeader = sCcHeader & CText(utMail.sCcDisplayName(iCtr), True)
            Else
                sCcHeader = sCcHeader & """" & Trim$(utMail.sCcAddr(iCtr)) & """"
            End If

            sCcHeader = sCcHeader & " <" & utMail.sCcAddr(iCtr) & ">"
            If iCtr < UBound(utMail.sCcAddr) Then sCcHeader = sCcHeader & ", " & vbCrLf & vbTab
        Next

        ' send Bcc: recipient addresses (more of the same)
        ' no display headers here, these are blind
        For iCtr = 0 To UBound(utMail.sBccAddr)
            ' send the recipient address & wait for a reply
            pbRequestAccepted = False
            '.SendData "RCPT TO: <" & Trim$(utMail.sBccAddr(iCtr)) & ">" & vbCrLf
            Call SocketSend(sckMail, "RCPT TO: <" & Trim$(utMail.sBccAddr(iCtr)) & ">" & vbCrLf)
            If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
            If pbExitImmediately Then Exit Sub
        Next

        ' tell the remote host we're ready to send data
        mvarStrStatus = mvarStrStatus & vbNewLine & "Sending Message..."
        RaiseEvent Status("Sending Message...")
        pbDataOK = False
        '.SendData "DATA" & vbCrLf
        Call SocketSend(sckMail, "DATA" & vbCrLf)
        If Not WaitUntilTrue(pbDataOK, plMessageTimeOut, True) Then Exit Sub
        If pbExitImmediately Then Exit Sub
        
        ' OK, the host is ready for data, this is where the mail message starts
        ' Send the mail headers (the ones displayed on the target email client)
        pbRequestAccepted = False
        
        '.SendData "Message-ID: <" & MessageID & ">" & vbCrLf
        Call SocketSend(sckMail, "Message-ID: <" & MessageID & ">" & vbCrLf)

        ' from, to, cc & subject headers..
        If Len(Trim$(utMail.sFromDisplayName)) Then
            sSenderName = CText(utMail.sFromDisplayName, True)
        Else
            sSenderName = """" & utMail.sFromAddr & """"
        End If
        sSenderName = sSenderName & " <" & utMail.sFromAddr & ">"
        '.SendData "From: " & sSenderName & vbCrLf
        Call SocketSend(sckMail, "From: " & sSenderName & vbCrLf)
        '.SendData "To: " & sToHeader & vbCrLf
        Call SocketSend(sckMail, "To: " & sToHeader & vbCrLf)
        If Len(sCcHeader) Then
           '.SendData "Cc: " & sCcHeader & vbCrLf
           Call SocketSend(sckMail, "Cc: " & sCcHeader & vbCrLf)
        End If
        
        '.SendData "Subject: " & CText(utMail.sSubject) & vbCrLf
        Call SocketSend(sckMail, "Subject: " & CText(utMail.sSubject) & vbCrLf)
        If Len(utMail.sReplyToAddr) Then
           '.SendData "Reply-to: <" & utMail.sReplyToAddr & ">" & vbCrLf
           Call SocketSend(sckMail, "Reply-to: <" & utMail.sReplyToAddr & ">" & vbCrLf)
        End If
        ' send English formated date/time string
        '.SendData "Date: " & GetTimeDate & vbCrLf
        Call SocketSend(sckMail, "Date: " & GetTimeDate & vbCrLf)

        ' MIME headers...
        If etEncodeType = MIME_ENCODE Then
            ' create a Unique-Boundary string for multi-part MIME encoding
            strBoundry = "----_=_NextPart_000_" & Right$("00000000" & Hex$(Date), 8) & "." & Right$("00000000" & Hex$(CLng(Time * 10 ^ 8)), 8)
            '.SendData "MIME-Version: 1.0" & vbCrLf
            Call SocketSend(sckMail, "MIME-Version: 1.0" & vbCrLf)
            If etPriority <> NORMAL_PRIORITY Then
                '.SendData "X-Priority: " & Trim$(Str$(etPriority)) & vbCrLf
                '.SendData "X-MSMail-Priority: " & psPriority & vbCrLf
                Call SocketSend(sckMail, "X-Priority: " & Trim$(Str$(etPriority)) & vbCrLf)
                Call SocketSend(sckMail, "X-MSMail-Priority: " & psPriority & vbCrLf)
            End If
            If pbReceipt Then
               '.SendData "Disposition-Notification-To: " & sSenderName & vbCrLf
               Call SocketSend(sckMail, "Disposition-Notification-To: " & sSenderName & vbCrLf)
            End If
            ' if it's multi part send the boundry info
            If bMimeMultiPart Then
                If bRelatedLinks Then
                    '.SendData "Content-Type: multipart/related;" & vbCrLf
                    Call SocketSend(sckMail, "Content-Type: multipart/related;" & vbCrLf)
                Else
                    '.SendData "Content-Type: multipart/mixed;" & vbCrLf
                    Call SocketSend(sckMail, "Content-Type: multipart/mixed;" & vbCrLf)
                End If
                '.SendData vbTab & "boundary=" & """" & strBoundry & """" & vbCrLf & vbCrLf
                '.SendData "This is a multi-part message in MIME format." & vbCrLf & vbCrLf
                '' send the MIME boundry and content headers for the message body
                '.SendData "--" & strBoundry & vbCrLf
                Call SocketSend(sckMail, vbTab & "boundary=" & """" & strBoundry & """" & vbCrLf & vbCrLf)
                Call SocketSend(sckMail, "This is a multi-part message in MIME format." & vbCrLf & vbCrLf)
                ' send the MIME boundry and content headers for the message body
                Call SocketSend(sckMail, "--" & strBoundry & vbCrLf)
            End If
            ' plain or html text...
            If pbHtmlText Then sTxt = "text/html;" Else sTxt = "text/plain;"
            '.SendData "Content-Type: " & sTxt & vbCrLf
            '.SendData vbTab & "charset=" & """" & CHAR_SET & """" & vbCrLf
            Call SocketSend(sckMail, "Content-Type: " & sTxt & vbCrLf)
            Call SocketSend(sckMail, vbTab & "charset=" & """" & CHAR_SET & """" & vbCrLf)
            If pb8BitMail Then sTxt = "8bit" Else sTxt = "7bit"
            '.SendData "Content-Transfer-Encoding: " & sTxt & vbCrLf
            Call SocketSend(sckMail, "Content-Transfer-Encoding: " & sTxt & vbCrLf)
            ' if we're sending html & the user supplied the content base then send it too
            If pbHtmlText Then
               If Len(psContentBase) Then
                  '.SendData "Content-Base: " & """" & psContentBase & """" & vbCrLf
                  Call SocketSend(sckMail, "Content-Base: " & """" & psContentBase & """" & vbCrLf)
               End If
            End If
        End If

        '.SendData vbCrLf
         Call SocketSend(sckMail, vbCrLf)
         
        ' Send the message body
        '.SendData utMail.sMailMessage & vbCrLf ''& vbCrLf & vbCrLf
        Call SocketSend(sckMail, utMail.sMailMessage & vbCrLf)

        ' Send attachments, if any...
        If utMail.lAttachCount Then
           '.SendData vbCrLf
           Call SocketSend(sckMail, vbCrLf)
        End If
        For iCtr = 0 To utMail.lAttachCount - 1
            If utMail.bAttachCID(iCtr) Then
                mvarStrStatus = mvarStrStatus & vbNewLine & "Sending Embedded File, " & utMail.sAttachNameOnly(iCtr) & "..."
                RaiseEvent Status("Sending Embedded File, " & utMail.sAttachNameOnly(iCtr) & "...")
            Else
                mvarStrStatus = mvarStrStatus & vbNewLine & "Sending Attachment, " & utMail.sAttachNameOnly(iCtr) & "..."
                RaiseEvent Status("Sending Attachment, " & utMail.sAttachNameOnly(iCtr) & "...")
            End If
            If etEncodeType = MIME_ENCODE Then
                ' send the next MIME boundry & content headers
                '.SendData "--" & strBoundry & vbCrLf
                '.SendData "Content-Type: " & GetContentType(utMail.sAttachNameOnly(iCtr)) & ";" & vbCrLf
                '.SendData vbTab & "name=" & """" & utMail.sAttachNameOnly(iCtr) & """" & vbCrLf
                '.SendData "Content-Transfer-Encoding: base64" & vbCrLf
                '.SendData "Content-Disposition: attachment;" & vbCrLf
                '.SendData vbTab & "filename=" & """" & utMail.sAttachNameOnly(iCtr) & """" & vbCrLf
                Call SocketSend(sckMail, "--" & strBoundry & vbCrLf)
                Call SocketSend(sckMail, "Content-Type: " & GetContentType(utMail.sAttachNameOnly(iCtr)) & ";" & vbCrLf)
                Call SocketSend(sckMail, vbTab & "name=" & """" & utMail.sAttachNameOnly(iCtr) & """" & vbCrLf)
                Call SocketSend(sckMail, "Content-Transfer-Encoding: base64" & vbCrLf)
                Call SocketSend(sckMail, "Content-Disposition: attachment;" & vbCrLf)
                Call SocketSend(sckMail, vbTab & "filename=" & """" & utMail.sAttachNameOnly(iCtr) & """" & vbCrLf)
                If (bRelatedLinks And utMail.bAttachCID(iCtr)) Then
                     '.SendData "Content-ID: <" & utMail.sAttachNameOnly(iCtr) & ">" & vbCrLf
                     Call SocketSend(sckMail, "Content-ID: <" & utMail.sAttachNameOnly(iCtr) & ">" & vbCrLf)
                End If
                '.SendData vbCrLf
                Call SocketSend(sckMail, vbCrLf)
                ' send the encoded file
                EncodeAndSendFile utMail.sAttachment(iCtr), MIME_ENCODE
                If pbExitImmediately Then Exit Sub
                '.SendData vbCrLf
                Call SocketSend(sckMail, vbCrLf)
            Else
                ' start a UUEncode session
                '.SendData "begin 600 " & utMail.sAttachNameOnly(iCtr) & vbCrLf
                Call SocketSend(sckMail, "begin 600 " & utMail.sAttachNameOnly(iCtr) & vbCrLf)
                ' send the encoded file
                EncodeAndSendFile utMail.sAttachment(iCtr), UU_ENCODE
                If pbExitImmediately Then Exit Sub
                ' send the ending sequence
                '.SendData "end" & vbCrLf
                Call SocketSend(sckMail, "end" & vbCrLf)
            End If

            ' the sckMail Send buffer now holds the current file
            ' if its a large file, wait here for the buffer to
            ' empty before loading the next one
            Do While plBytesRemaining > 4096
                ' timeout code...
                fStart = Timer
                ' Deal with timer being reset at Midnight
                If fStart + plMessageTimeOut < 86400 Then
                    fTimeOut = fStart + plMessageTimeOut
                Else
                    fTimeOut = (fStart - 86400) + plMessageTimeOut
                End If
                ' wait for a change in the send buffer
                ' if it's changing, everything is OK
                lSendBuffSize = plBytesRemaining
                Do Until lSendBuffSize <> plBytesRemaining
                    If plBytesRemaining < 4096 Then Exit Do
                    If Timer >= fTimeOut Then
                        TimeOut
                        Exit Sub
                    End If
                    Sleep (10)
                    DoEvents
                Loop
            Loop
        Next iCtr

        If bMimeMultiPart = True Then
            ' send the MIME closing boundry header
            '.SendData "--" & strBoundry & "--" & vbCrLf
            Call SocketSend(sckMail, "--" & strBoundry & "--" & vbCrLf)
        End If
        
        ' Send the 'end of mail' string
        pbRequestAccepted = False
        '.SendData "." & vbCrLf
        Call SocketSend(sckMail, "." & vbCrLf)
        If Not WaitUntilTrue(pbRequestAccepted, plMessageTimeOut, True) Then Exit Sub
        
    End With

    ' send completion notifications...
    RaiseEvent Progress(100)
    mvarStrStatus = mvarStrStatus & vbNewLine & "Transmission Complete..."
    RaiseEvent Status("Transmission Complete...")

    ' if the Public Function Connect() was called,
    ' stay connected to the host, otherwise disconnect
    If Not pbManualDisconnect Then DisconnectFromHost

    SendSuccesful = True
    RaiseEvent SendSuccesful
    

    Exit Sub

Err_Send:

    ' add the error to the error collection
    AddError Err.Description
    SendFail

End Sub

Public Sub Shutdown()

' stub function, here to maintain binary
' compatibility with previous versions.

End Sub


' ******************************************************************************
' *      Private Class Functions                                               *
' ******************************************************************************

Private Function GetTimeDate() As String

    GetTimeDate = psDay(Weekday(Now)) & ", " & Day(Now) & " " & psMonth(Month(Now)) & " " & Format$(Now, "YYYY") & " " & _
                  Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00") & " " & psTimeZoneBias

End Function

Private Sub AddError(ByVal ErrStr As String)
   ' add error string to the error collection
   On Local Error Resume Next
   pColErrors.Add ErrStr, ErrStr
End Sub

Private Function AddressStringToLong(ByVal tmp As String) As Long

    ' convert an ip address string to a long value
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim sParts()    As String

    sParts = Split(tmp, ".")

    If UBound(sParts) <> 3 Then
        AddressStringToLong = 0
        Exit Function
    End If

    ' build the long value out of the
    ' hex of the extracted strings
    AddressStringToLong = Val("&H" & Right$("00" & Hex$(sParts(3)), 2) & _
            Right$("00" & Hex$(sParts(2)), 2) & _
            Right$("00" & Hex$(sParts(1)), 2) & _
            Right$("00" & Hex$(sParts(0)), 2))

End Function

Private Function bInEXE() As Boolean

' ******************************************************************************
'
' Synopsis:     Check if application is running in the VB IDE or stand alone EXE.
'
' Parameters:   none
'
' Return:       True if running in EXE, False if running in IDE
'
' Description:
'
' Debug.print 1/0 will error produce a divide by zero error if running in IDE.
' If running in exe debug.print statement will be ignored
'
' ******************************************************************************

' modified version of Brian Gillham's code
' sample available at www.freevbcode.com

    On Local Error GoTo ErrorHandler

    Debug.Print 1 / 0                               ' this line will fail in the IDE
    bInEXE = True                                   ' this line will execute only in EXE or dll

    Exit Function

ErrorHandler:

    bInEXE = False

End Function

Private Function ConnectToHost() As Boolean
   Dim iCtr            As Integer
   Dim sHello          As String

   If bInEXE Then On Local Error GoTo Connect_Error

   ' already connected?
   If sckMail.State = sckConnected Then
      ConnectToHost = True
      Exit Function
   ElseIf sckMail.State <> sckClosed Then
      sckMail.Close
      DoEvents
   End If
   
   ' check the SMTP host
   If Len(psSMTPHost) = 0 Then
      psSMTPHost = MXQuery
      If Len(psSMTPHost) = 0 Then
         AddError ERR_INVALID_HOST
         Exit Function
      End If
   End If

   ' Pop3 Authentication first?
   If pbUsePopAuthentication Then
      mvarStrStatus = mvarStrStatus & vbNewLine & "Connecting to POP3 Server (" & Me.POP3Host & ")..."
      RaiseEvent Status("Connecting to POP3 Server (" & Me.POP3Host & ")...")
      pbExitImmediately = False
      pbConnected = False
      pbPopAuthOk = False
      plPop3Status = 0
      If Len(psPop3Host) = 0 Then
         AddError ERR_INVALID_POP_HOST
         SendFail
         Exit Function
      End If
      ' open POP3 connection
      With sckMail
         .RemoteHost = psPop3Host
         .RemotePort = POP3_PORT
         For iCtr = 1 To plConnectRetry
            If .State <> sckConnected Then
               If .State = sckClosed Then
                  .Connect
               End If
               If WaitUntilTrue(pbConnected, plConnectTimeout, False) Then Exit For
               If pbExitImmediately Then Exit Function
               If .State = sckError Then .Close
            Else
               pbConnected = True
               Exit For
            End If
         Next
         ' data arival event responds automatically
         WaitUntilTrue pbPopAuthOk, plConnectTimeout, False
         .Close
      End With
      DoEvents
      If pbExitImmediately Then Exit Function
      mvarStrStatus = mvarStrStatus & vbNewLine & "POP3 Authentication Successful..."
      RaiseEvent Status("POP3 Authentication Successful...")
   End If

   ' reset var's
   pbRequestAccepted = False
   pbDataOK = False
   pbAuthLoginSupported = False
   pbAuthMailFromOK = False
   pbAuthLoginSuccess = False
   pbExitImmediately = False
   ConnectToHost = False
   pbConnected = False
   
   ' open an SMTP session...
   With sckMail

       ' setup the port
       If .State <> sckClosed Then
           .Close
           DoEvents
       End If
       .RemoteHost = psSMTPHost
       .RemotePort = plSMTPPort
       DoEvents
       
       ' open a connection with the remote host
       ' try 'plConnectRetry' times before giving up
       mvarStrStatus = mvarStrStatus & vbNewLine & "Connecting to SMTP Server (" & Me.SMTPHost & ")..."
       RaiseEvent Status("Connecting to SMTP Server (" & Me.SMTPHost & ")...")
       For iCtr = 1 To plConnectRetry
           If .State <> sckConnected Then
               Debug.Print .State
               pbConnected = False
               If .State = sckClosed Then
                  If mvarSSL Then
                     .protocol = sckTCPProtocol
                  End If
                  .Connect 'psSMTPHost, plSMTPPort
               End If
               If pbUsePopAuthentication Then
                  If WaitUntilTrue(pbConnected, plConnectTimeout, False) Then
                     Exit For
                  End If
               Else
                  If WaitUntilTrue(pbConnected, 0, False) Then
                     Exit For
                  End If
               End If
               'If WaitUntilTrue(pbConnected, IIf(pbUsePopAuthentication, plConnectTimeout, 0), False) Then Exit For
               'If WaitUntilTrue(pbConnected, plConnectTimeout, False) Then Exit For
               If pbExitImmediately Then
                  Exit Function
               End If
               If .State = sckError Then
                  .Close
               End If
           Else
               pbConnected = True
               Exit For
           End If
           DoEvents
       Next
       DoEvents
       
       ' if the connect attempt failed, exit
       If WaitUntilTrue(pbRequestAccepted, plConnectTimeout, False) = False Or _
               pbConnected = False Or _
               pbExitImmediately = True Or _
               .State <> sckConnected Then
           
           TimeOut
           Exit Function
       End If

        ' once a connection is established, say 'hello
       mvarStrStatus = mvarStrStatus & vbNewLine & "Initializing Communications..."
       RaiseEvent Status("Initializing Communications...")
       pbRequestAccepted = False
       ' EHLO is the extended (ESMTP) hello command, HELO is the standard hello command
       If pbUseAuthentication Then sHello = "EHLO " Else sHello = "HELO "
       .SendData sHello & Mid$(utMail.sFromAddr, InStr(utMail.sFromAddr, "@") + 1) & vbCrLf
       If Not WaitUntilTrue(pbRequestAccepted, plConnectTimeout, True) Then Exit Function

       ' Login Authentication ...
       ' the 'EHLO" command will cause the host to send a list of supported extensions
       ' via a series of 250 replies, wait to see if 'Auth Logon' is listed. The Sub
       ' sckMail_DataArrival will set pbUseAuthentication = True if Auth Login is
       ' supported by the remote host. If it is supported, Sub sckMail_DataArrival will
       ' respond to the host's Username & Password requests (psUserName, psPassword).
       If pbUseAuthentication = True Then
           If WaitUntilTrue(pbAuthLoginSupported, 5, False) Then
               mvarStrStatus = mvarStrStatus & vbNewLine & "Sending Login Authentication..."
               RaiseEvent Status("Sending Login Authentication...")
               .SendData "AUTH Login" & vbCrLf
               If WaitUntilTrue(pbAuthLoginSuccess, 5, False) Then
                   mvarStrStatus = mvarStrStatus & vbNewLine & "Host Login OK!"
                   RaiseEvent Status("Host Login OK!")
               Else
                   mvarStrStatus = mvarStrStatus & vbNewLine & "Host Login Failed!"
                   RaiseEvent Status("Host Login Failed!")
                   pbExitImmediately = True
                   Exit Function
               End If
               If pbExitImmediately Then Exit Function
           Else
               mvarStrStatus = mvarStrStatus & vbNewLine & "Login Not Supported by Host, Continuing..."
               RaiseEvent Status("Login Not Supported by Host, Continuing...")
           End If
       End If

   End With

   ConnectToHost = True

   Exit Function


Connect_Error:

   mvarStrStatus = mvarStrStatus & vbNewLine & "Connect error: " & Err.Description
   RaiseEvent Status("Connect error: " & Err.Description)


End Function

Private Function CText(sIn As String, Optional bAddQuotesIfNotConverted As Boolean = False) As String

'   'B' or 'Q' encode an ASCII string, defined in RFC 2047...
'   The "B" encoding is identical to the "BASE64" encoding defined by RFC 1521.
'   The "Q" encoding is similar to the "Quoted-Printable" content-
'   transfer-encoding defined in RFC 1521.  It is designed to allow text
'   containing mostly ASCII characters to be decipherable on an ASCII
'   terminal without decoding.
    
'   perform both & return the smaller of the two

    Dim iPtr            As Integer
    Dim bNeedsEncoding  As Boolean
    Dim iMax            As Integer
    Dim sChr            As String
    Dim sLine           As String
    Dim sQCode          As String
    Dim sBCode          As String
    Dim bytTmp()        As Byte
    
    
    If bInEXE Then On Local Error GoTo Err_Qtext
    

    ' scan for 8bit characters
    bytTmp() = StrConv(sIn, vbFromUnicode)

    For iPtr = 0 To UBound(bytTmp)
        If bytTmp(iPtr) > 126 Then
             bNeedsEncoding = True
            Exit For
        End If
    Next

    If Not bNeedsEncoding Then
        If bAddQuotesIfNotConverted Then
            ' if its part of an address string it needs
            ' to be quoted if it's returned as plain text
            CText = """" & sIn & """"
        Else
            CText = sIn
        End If
        Exit Function
    End If
        
'    ' Q encode
'    iMax = 54
'    For iPtr = 1 To Len(sIn)
'        sChr = Mid$(sIn, iPtr, 1)
'        Select Case Asc(sChr)
'            ' pass printable ascii as is, except "=" "?" "_" " "
'            Case 33 To 60, 62, 64 To 94, 96 To 126
'                sLine = sLine & sChr
'            ' convert space to underscore (for readability)
'            Case 32
'                sLine = sLine & "_"
'            ' Q Code everything else
'            Case Is > 255, Is < 0
'                sLine = sLine & "=" & Left$(Hex$(Asc(sChr)), 2) & _
'                                "=" & Right$(Hex$(Asc(sChr)), 2)
'            Case Else
'                sLine = sLine & "=" & Right$("00" & Hex$(Asc(sChr)), 2)
'
'        End Select
'        If Len(sLine) >= iMax Then
'            sQCode = sQCode & Q_CODE_HDR & sLine & CODE_END
'            If iPtr < Len(sIn) Then sQCode = sQCode & vbCrLf & vbTab
'            sLine = ""
'        End If
'    Next
'    sQCode = sQCode & Q_CODE_HDR & sLine & CODE_END
    
    
    ' B encode
    iMax = 42
    sLine = sIn
    Do While Len(sLine)
        ' encode a line, maximun lenght is 76 characters
        ' <header><base64encoded text><end><CrLf>
        sBCode = sBCode & B_CODE_HDR & EncodeBase64String(Mid$(sLine, 1, iMax))
        ' strip off the CrLf & add END_CODE , CrLF & Tab
        sBCode = Mid$(sBCode, 1, Len(sBCode) - 2) & CODE_END
        ' get ready for the next line
        sLine = Mid$(sLine, iMax + 1)
        If Len(sLine) Then sBCode = sBCode & vbCrLf & vbTab
    Loop
'
'    If Len(sQCode) < Len(sBCode) Then
'        CText = sQCode
'    Else
        CText = sBCode
'    End If
   
    Exit Function
    
Err_Qtext:
    
    CText = sIn
    
End Function

Public Function DecodeBase64String(ByVal str2Decode As String) As String

' ******************************************************************************
'
' Synopsis:     Decode a Base 64 string
'
' Parameters:   str2Decode  - The base 64 encoded input string
'
' Return:       decoded string
'
' Description:
' Coerce 4 base 64 encoded bytes into 3 decoded bytes by converting 4, 6 bit
' values (0 to 63) into 3, 8 bit values. Transform the 8 bit value into its
' ascii character equivalent. Stop converting at the end of the input string
' or when the first '=' (equal sign) is encountered.
'
' ******************************************************************************

    Dim lPtr            As Long
    Dim iValue          As Integer
    Dim iLen            As Integer
    Dim iCtr            As Integer
    Dim Bits(1 To 4)    As Byte
    Dim strDecode       As String

    ' for each 4 character group....
    For lPtr = 1 To Len(str2Decode) Step 4
        iLen = 4
        For iCtr = 0 To 3
            ' retrive the base 64 value, 4 at a time
            iValue = InStr(1, BASE64CHR, Mid$(str2Decode, lPtr + iCtr, 1), vbBinaryCompare)
            Select Case iValue
                ' A~Za~z0~9+/
                Case 1 To 64: Bits(iCtr + 1) = iValue - 1
                ' =
                Case 65
                    iLen = iCtr
                    Exit For
                ' not found
                Case 0: Exit Function
            End Select
        Next

        ' convert the 4, 6 bit values into 3, 8 bit values
        Bits(1) = Bits(1) * &H4 + (Bits(2) And &H30) \ &H10
        Bits(2) = (Bits(2) And &HF) * &H10 + (Bits(3) And &H3C) \ &H4
        Bits(3) = (Bits(3) And &H3) * &H40 + Bits(4)

        ' add the three new characters to the output string
        For iCtr = 1 To iLen - 1
            strDecode = strDecode & Chr$(Bits(iCtr))
        Next

    Next

    DecodeBase64String = strDecode

End Function

Private Sub DisconnectFromHost()

    Dim lCtr        As Long
    
    On Local Error Resume Next
    
    With sckMail
        ' notify the user
        If .State <> sckClosed Then
           mvarStrStatus = mvarStrStatus & vbNewLine & "Closing Connection..."
           RaiseEvent Status("Closing Connection...")
        End If

        ' tell the host we're closing the connection...
        If .State = sckConnected Then
            pbRequestAccepted = False
            .SendData "QUIT" & vbCrLf
            WaitUntilTrue pbRequestAccepted, 2, False
        End If

        ' close the connection
        .Close
         DoEvents
    End With

End Sub

Private Sub EncodeAndSendFile(ByVal strFile As String, ByVal Encode As ENCODE_METHOD)

' ******************************************************************************
'
' Synopsis:     Send a file attachment via an open socket
'
' Parameters:   strFile  - The input file name
'               Encode   -  type of encoding to use; MIME or UU
'
' Return:       nothing
'
' Description:
' Open the file & read characters in. Send the characters through the
' appropriate encoder, either MIME (Base64) or UUEncode, before
' tranmission via an open socket.
'
' ******************************************************************************

    Dim hFile               As Integer              ' file handle
    Dim sValue              As String               ' temp string buffer
    Dim bInFile()           As Byte                 ' byte array file buffer
    Dim lEventCtr           As Long                 ' counter
    Dim lChunkSize          As Long                 ' number of bytes to get
    Dim lNumBytes           As Long                 ' file pointer

    ' in case there's a file io error
    If bInEXE Then On Local Error GoTo File_Error

    ' open the file
    hFile = FreeFile
    Open strFile For Binary Access Read As #hFile
    
    ' bytes to read
    lNumBytes = LOF(hFile)

    If Encode = MIME_ENCODE Then

        Do While lNumBytes
            ' set input buffer size, MUST be a multiple of 57
            lChunkSize = IIf(lNumBytes > 11400, 11400, lNumBytes)

            ' set to true in sckMail.SendProgress Event
            pbSendProgress = False
            
            ' read & Base 64 encode a group of characters
            ' changed from 'InputB' to 'Get' to improve performance
            ' on Netware servers/clients, thanks to Richard Gatewood.
            'bInFile = InputB(lChunkSize, #hFile)       ' nw change (remove)
            ReDim bInFile(lChunkSize - 1)               ' nw change (add)
            Get #hFile, , bInFile()                     ' nw change (add)
            If sckMail.State = sckConnected Then
                sckMail.SendData EncodeBase64Byte(bInFile)
            Else
                Err.Raise 0, , "Socket not Open"
            End If
            ' adjust file pointer
            lNumBytes = lNumBytes - lChunkSize

            'DoEvents
            ' wait for sckMail.SendProgress Event to fire
            ' suggested by David Hill to fix an issue with a very fast machine
            WaitUntilTrue pbSendProgress, 2, False
        Loop

    ElseIf Encode = UU_ENCODE Then

        Do While lNumBytes
            ' set input buffer size, MUST be 45
            lChunkSize = IIf(lNumBytes > 45, 45, lNumBytes)

            ' read & UU encode a line of characters
            sValue = Input(lChunkSize, #hFile)
            If sckMail.State = sckConnected Then
                sckMail.SendData UUEncodeString(sValue) & vbCrLf
            Else
                Err.Raise 0, , "Socket not Open"
            End If
            
            ' adjust file pointer
            lNumBytes = lNumBytes - lChunkSize
            
            ' DoEvents (occasionally)
            lEventCtr = lEventCtr + 1
            If lEventCtr Mod 50 = 0 Then DoEvents
        Loop

    End If

File_Done:

    Close #hFile
    Exit Sub

File_Error:

    AddError Err.Description
    SendFail
    pbExitImmediately = True
    Resume File_Done
    
End Sub

Private Function EncodeBase64Byte(InArray() As Byte) As Byte()

'******************************************************************************
'
' Synopsis:     Base 64 encode a byte array
'
' Parameters:   InArray  - The input byte array
'
' Return:       encoded byte array
'
' Description:
'   Convert a byte array to a Base 64 encoded byte array. Coerce 3 bytes into
'   4 by converting 3, 8 bit bytes into 4, 6 bit values. Each 6 bit value
'   (0 to 63) is then used as a pointer into a base64 byte array to derive a
'   character.
'
'******************************************************************************

Dim lInPtr              As Long         ' pointer into input array
Dim lOutPtr             As Long         ' pointer into output array
Dim OutArray()          As Byte         ' output byte array buffer
Dim lLen                As Long         ' number of extra bytes past 3 byte boundry
Dim iNewLine            As Long         ' line counter

' if size of input array is not a multiple of 3,
' increase it to the next multiple of 3
lLen = (UBound(InArray) - LBound(InArray) + 1) Mod 3
If lLen Then
    lLen = 3 - lLen
    ReDim Preserve InArray(UBound(InArray) + lLen)
End If

' create an output buffer
ReDim OutArray(UBound(InArray) * 2 + 100)

' step through the input array, 3 bytes at a time
For lInPtr = 0 To UBound(InArray) Step 3
    
    ' add CrLf as required
    If iNewLine = 19 Then
        OutArray(lOutPtr) = 13
        OutArray(lOutPtr + 1) = 10
        lOutPtr = lOutPtr + 2
        iNewLine = 0
    End If
    
    ' convert 3 bytes into 4 base 64 encoded bytes
    OutArray(lOutPtr) = pbBase64Byt((InArray(lInPtr) And &HFC) \ 4)
    OutArray(lOutPtr + 1) = pbBase64Byt((InArray(lInPtr) And &H3) * &H10 + (InArray(lInPtr + 1) And &HF0) \ &H10)
    OutArray(lOutPtr + 2) = pbBase64Byt((InArray(lInPtr + 1) And &HF) * 4 + (InArray(lInPtr + 2) And &HC0) \ &H40)
    OutArray(lOutPtr + 3) = pbBase64Byt(InArray(lInPtr + 2) And &H3F)
    
    ' update pointers
    lOutPtr = lOutPtr + 4
    iNewLine = iNewLine + 1
Next

' add terminator '=' as required
Select Case lLen
    Case 1: OutArray(lOutPtr - 1) = 61
    Case 2: OutArray(lOutPtr - 1) = 61: OutArray(lOutPtr - 2) = 61
End Select

' add CrLf if not already there
If OutArray(lOutPtr - 2) <> 13 Then
    OutArray(lOutPtr) = 13
    OutArray(lOutPtr + 1) = 10
    lOutPtr = lOutPtr + 2
End If

' resize output buffer and return
ReDim Preserve OutArray(lOutPtr - 1)
EncodeBase64Byte = OutArray

End Function

Private Function EncodeBase64String(ByRef str2Encode As String) As String

' ******************************************************************************
'
' Synopsis:     Base 64 encode a string
'
' Parameters:   str2Encode  - The input string
'
' Return:       encoded string
'
' Description:
' Convert a string to a byte array and pass to EncodeBase64Byte function (above)
' for Base64 conversion. Convert byte array back to a string and return.
'
' ******************************************************************************

    Dim tmpByte()   As Byte

    If Len(str2Encode) Then
    
        ' convert string to byte array
        tmpByte = StrConv(str2Encode, vbFromUnicode)

        ' pass to the byte array encoder
        tmpByte = EncodeBase64Byte(tmpByte)

        ' convert back to string & return
        EncodeBase64String = StrConv(tmpByte, vbUnicode)

    End If
    
End Function

Private Function EstimateMailSize() As Long

' ******************************************************************************
'
' Synopsis:     Estimate the size (number of bytes) of the mail message
'
' Parameters:   none
'
' Return:       long - number of bytes
'
' Description:
' Estimate the size in bytes of the mail message being sent. Include the
' message body, headers, attachments, etc. Account for type of encoding.
' The result is used to calculate send progress.
'
' ******************************************************************************

    Dim lNumBytes       As Long
    Dim iCtr            As Integer

    lNumBytes = 93

    ' Mail From
    lNumBytes = lNumBytes + Len(utMail.sFromAddr)

    ' login authentication
    If pbUseAuthentication Then
        lNumBytes = lNumBytes + 25 + Len(utMail.sFromAddr)
        If Len(psUserName) > 0 Then lNumBytes = lNumBytes + (Len(psUserName) * 4 \ 3)
        If Len(psPassword) > 0 Then lNumBytes = lNumBytes + (Len(psPassword) * 4 \ 3)
    End If

    ' To: recipients
    For iCtr = 0 To UBound(utMail.sToAddr)
        lNumBytes = lNumBytes + 15 + Len(utMail.sToAddr(iCtr)) * 2  ' sent twice, RCPT & 'To:' header
        If iCtr > 0 Then lNumBytes = lNumBytes + 6
    Next

    ' To Display
    For iCtr = 0 To UBound(utMail.sToDisplayName)
        lNumBytes = lNumBytes + Len(utMail.sToDisplayName(iCtr)) + 11
    Next

    ' Cc: recipients
    For iCtr = 0 To UBound(utMail.sCcAddr)
        lNumBytes = lNumBytes + 15 + Len(utMail.sCcAddr(iCtr)) * 2  ' sent twice, RCPT & 'Cc:' header
        If iCtr > 0 Then lNumBytes = lNumBytes + 6                  ' header
    Next

    ' Cc Display
    For iCtr = 0 To UBound(utMail.sCcDisplayName)
        lNumBytes = lNumBytes + Len(utMail.sCcDisplayName(iCtr)) + 11
    Next

    ' Bcc: recipients
    For iCtr = 0 To UBound(utMail.sBccAddr)
        lNumBytes = lNumBytes + 15 + Len(utMail.sBccAddr(iCtr))  ' RCPT & 'Bcc:' header
        If iCtr > 0 Then lNumBytes = lNumBytes + 6               ' header
    Next

    ' From:
    If Len(utMail.sFromDisplayName) Then lNumBytes = lNumBytes + Len(utMail.sFromDisplayName) + 3
    lNumBytes = lNumBytes + Len(utMail.sFromAddr)

    ' ReplyTo
    If Len(utMail.sReplyToAddr) Then lNumBytes = lNumBytes + Len(utMail.sReplyToAddr) + 14

    ' Subject
    lNumBytes = lNumBytes + Len(utMail.sSubject)

    ' Message body
    lNumBytes = lNumBytes + Len(utMail.sMailMessage)

    ' MIME headers....
    If etEncodeType = MIME_ENCODE Then
        lNumBytes = lNumBytes + 64
        If pbHtmlText = True And Len(psContentBase) > 0 Then lNumBytes = lNumBytes + 18 + Len(psContentBase)
        If pbReceipt Then lNumBytes = lNumBytes + 36 + Len(utMail.sFromDisplayName) + Len(utMail.sFromAddr)
    End If

    ' attachments
    If utMail.lAttachCount > 0 Then
        If etEncodeType = MIME_ENCODE Then
            lNumBytes = lNumBytes + utMail.lAttachFileSize * 4 \ 3 + 42  ' length of encoded file
            lNumBytes = lNumBytes + (utMail.lAttachFileSize \ 57) * 2    ' add CrLf for each line
            lNumBytes = lNumBytes + utMail.lAttachNameSize * 2           ' add file name twice
            lNumBytes = lNumBytes + (utMail.lAttachCount * 182)          ' attachment header per file
            lNumBytes = lNumBytes + 290                                  ' additional MIME headers
        Else
            lNumBytes = lNumBytes + utMail.lAttachFileSize * 4 \ 3       ' length of encoded file
            lNumBytes = lNumBytes + (utMail.lAttachFileSize \ 45) * 3    ' add length char + CrLf for each line
            lNumBytes = lNumBytes + utMail.lAttachNameSize               ' add file name once
            lNumBytes = lNumBytes + (utMail.lAttachCount * 20)           ' attachment header per file
        End If
    End If

    EstimateMailSize = lNumBytes

End Function

Private Function FormatMail(ByVal strIn As String) As String

' ******************************************************************************
'
' Synopsis:     Re-format text lines per RFC 821
'
' Parameters:   strIn   - The input string to be formated
'
' Return:       re-formated string
'
' Description:
' RFC 821 places the following restrictions on user text:
' 1) Before sending a line of mail text begining with a '.
' the sender will add an additional '.
'
' 2) The receiver checks each line of mail text, if a line is single '.
' it is the end of the mail message. If the first character is
' a '.' and there are other characters on the line, the first '.
' is deleted.
'
' 3) The maximum line lenght will not exceed 1000 characters
'
' ******************************************************************************

    Dim sTextLine()     As String
    Dim sRemainder      As String
    Dim sNewLine        As String
    Dim sDelimiter      As String
    Dim lPtr            As Long
    Dim lSplit          As Long

    If Len(strIn) = 0 Then Exit Function

    ' Select the correct delimiter character
    If InStr(strIn, vbCrLf) Then
        sDelimiter = vbCrLf
    ElseIf InStr(strIn, vbCr) Then
        sDelimiter = vbCr
    Else
        sDelimiter = vbNullString
    End If

    ' split the text into seperate lines
    sTextLine() = Split(strIn, sDelimiter)

    ' process each line
    For lPtr = 0 To UBound(sTextLine)
        ' check for lines starting with a '.
        ' when found, add a second '.
        If Left$(sTextLine(lPtr), 1) = "." Then sTextLine(lPtr) = "." & sTextLine(lPtr)

        ' check that the line is not too long (account for 2 extra characters - vbCrLf)
        ' break into smaller elements as required
        If Len(sTextLine(lPtr)) > MAX_TEXTLINE_LEN - 2 Then
            sRemainder = sTextLine(lPtr)
            sNewLine = vbNullString
            If sDelimiter = vbNullString Then sDelimiter = vbCrLf
            Do While Len(sRemainder) > MAX_TEXTLINE_LEN - 2
                ' try to split at a space character, if not then split at MAX_TEXTLINE_LEN - 2
                lSplit = InStrRev(sRemainder, " ", MAX_TEXTLINE_LEN - 2)
                If lSplit = 0 Then lSplit = MAX_TEXTLINE_LEN - 2
                ' insert a vbCrLf at the split point
                sNewLine = sNewLine & Mid$(sRemainder, 1, lSplit) & sDelimiter
                sRemainder = Mid$(sRemainder, lSplit + 1)
            Loop
            sTextLine(lPtr) = sNewLine & sRemainder
        End If
    Next

    FormatMail = Join(sTextLine, sDelimiter)

End Function

Private Function GetAttachCID() As Boolean

    ' search the email body for tags with filenames that match the list of attached
    ' filenames, replace the path with a 'cid' and flag the array as having a valid CID
    ' example: <IMG SRC="/images/somefile.jpg"> is replaced with <IMG SRC="CID:somefile.jpg">

    Dim iCtr            As Integer
    Dim lPtr            As Long
    Dim lEndFirstPart   As Long
    Dim lStartLastPart  As Long
    Dim lQuotePos       As Long
    Dim lEqualPos       As Long
    Dim lNextPos        As Long
    Dim lGtPos          As Long
    Dim lLtPos          As Long
        
    If utMail.lAttachCount < 1 Then Exit Function
    If Not pbHtmlText Then Exit Function
    
    ' for each attached file
    For iCtr = 0 To utMail.lAttachCount - 1
        ' find the first occurance
        lPtr = InStr(1, utMail.sMailMessage, utMail.sAttachNameOnly(iCtr), vbTextCompare)
        Do While lPtr
            ' found an occurance of the file name,
            ' is it part of a tag?
            lLtPos = InStrRev(utMail.sMailMessage, "<", lPtr)
            lGtPos = IIf(lLtPos > 0, InStr(lLtPos, utMail.sMailMessage, ">"), 0)
            If lLtPos > 0 And lGtPos > 0 And lGtPos > lPtr And lLtPos < lPtr Then
                ' yes it's part of an HTML tag
                ' find the equal sign & quote if any exists
                lEqualPos = InStrRev(utMail.sMailMessage, "=", lPtr)
                lQuotePos = InStr(lEqualPos, utMail.sMailMessage, """")
                ' first part
                If lQuotePos > 0 And lQuotePos < lPtr Then
                    lEndFirstPart = lQuotePos
                Else
                    lEndFirstPart = lEqualPos
                End If
                ' last part
                lStartLastPart = lPtr + Len(utMail.sAttachNameOnly(iCtr))
                ' replace with "CID:somefile.jpg"
                utMail.sMailMessage = Mid$(utMail.sMailMessage, 1, lEndFirstPart) & _
                                      "cid:" & utMail.sAttachNameOnly(iCtr) & _
                                      Mid$(utMail.sMailMessage, lStartLastPart)
                utMail.bAttachCID(iCtr) = True
                GetAttachCID = True
                lNextPos = lEndFirstPart + Len(utMail.sAttachNameOnly(iCtr)) + 4
            Else
                lNextPos = lPtr + Len(utMail.sAttachNameOnly(iCtr))
            End If
            ' find the next one
            lPtr = InStr(lNextPos, utMail.sMailMessage, utMail.sAttachNameOnly(iCtr), vbTextCompare)
        Loop
    Next

End Function

Private Function MessageID() As String
    
    Dim iCtr            As Integer
    Dim iPtr            As Integer
    Dim sTmp            As String
    Dim ut              As SYSTEMTIME
    
    On Local Error Resume Next
    
    Randomize
    
    GetSystemTime ut
    MessageID = CStr(ut.wYear) & Format$(ut.wMonth, "00") & Format$(ut.wDay, "00") & _
                Format$(ut.wHour, "00") & Format$(ut.wMinute, "00") & Format$(ut.wSecond, "00") & _
                Format$(ut.wMilliseconds, "000") & "."
    
    For iCtr = 1 To 20
        iPtr = Int(Rnd * 61 + 1)
        MessageID = MessageID & Mid(BASE64CHR, iPtr, 1)
    Next
    
    MessageID = MessageID & "@"
     
    sTmp = GetRemoteHostName(sckMail.RemoteHostIP)
    If Len(sTmp) = 0 Then sTmp = sckMail.RemoteHost
    
    MessageID = MessageID & sTmp
    
End Function


Private Function IsDottedQuad(ByVal HostString As String) As Boolean

    ' verify that a string is 'xxx.xxx.xxx.xxx' format
    
    Dim sSplit()        As String
    Dim iCtr            As Integer

    ' split at the "."
    sSplit = Split(HostString, ".")

    ' should be 4 elements
    If UBound(sSplit) <> 3 Then Exit Function

    ' check each element
    For iCtr = 0 To 3
        ' should be numeric
        If Not IsNumeric(sSplit(iCtr)) Then Exit Function

        ' range check
        If iCtr = 0 Then
            If Val(sSplit(iCtr)) > 239 Then Exit Function
        Else
            If Val(sSplit(iCtr)) > 255 Then Exit Function
        End If
    Next
    
    IsDottedQuad = True

End Function

Private Function IsValidIPHost(ByVal HostString As String) As Boolean

    ' validate a host string

    Dim sHost               As String
    Dim sSplit()            As String

    sHost = UCase$(Trim$(HostString))

    If Len(sHost) = 0 Then Exit Function
    
    ' if it's a dotted quad it's OK
    If IsDottedQuad(sHost) Then
        IsValidIPHost = True
        Exit Function
    End If

    sSplit = Split(sHost, ".")

    ' it's not dotted quad, top level domain?
    If UBound(sSplit) > 0 And InStr(TOP_DOMAINS, sSplit(UBound(sSplit))) > 0 Then
        IsValidIPHost = True
        Exit Function
    End If

End Function

Private Function RegGet(ByVal sSettingName As String, ByVal sDefaultValue As String) As String

    If bInEXE Then On Local Error GoTo ERR_RegGet

    ' get registry setting
    RegGet = GetSetting(REG_KEY, SETTINGS_KEY, sSettingName, sDefaultValue)

    Exit Function

ERR_RegGet:

     RegGet = sDefaultValue
     
End Function

Private Sub RegSave(ByVal sSettingName As String, ByVal sNewValue As String)

    If bInEXE Then On Local Error GoTo ERR_RegSave

    ' save registry setting
    If pbPersistentSettings Then SaveSetting REG_KEY, SETTINGS_KEY, sSettingName, sNewValue

ERR_RegSave:

End Sub

Private Sub RemoveError(ByVal ErrStr As String)
   'remove an error string from the error collection
   Dim i   As Long
   
   On Local Error Resume Next
   ' walk the collection looking for the string to remove
   For i = 1 To pColErrors.Count
     If pColErrors(i) = ErrStr Then pColErrors.Remove ErrStr
   Next
End Sub

Private Sub SendFail()
   Dim iCtr            As Integer
   Dim sErrorString    As String

   ' report all errors to the user
   For iCtr = 1 To pColErrors.Count
       sErrorString = sErrorString & pColErrors(iCtr) & vbCrLf
   Next

   RaiseEvent SendFailed(sErrorString)
   SendSuccesful = False

   ' close the connection with the remote host
   If sckMail.State <> sckClosed Then DisconnectFromHost

   ' set flag to exit 'Send' Sub without further processing
   pbExitImmediately = True

    ' clear all errors
    Set pColErrors = New Collection
End Sub

Private Sub SocketsCleanup()

    ' Cleanup Windows sockets
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    WSACleanup

End Sub

Private Function SocketsInitialize() As Boolean

    ' Initialize Windows sockets
    '
    ' THIS CODE IS BASED ON FUNCTIONS
    ' WITHIN RICHARD DEEMING'S IP UTILITIES:
    ' http://www.freevbcode.com

    Dim WSAD            As WSADATA

    SocketsInitialize = False

    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then Exit Function
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then Exit Function

    SocketsInitialize = True

End Function

Private Sub TrimWhiteSpace(sInArray() As String)

    Dim i   As Long
    
    For i = LBound(sInArray) To UBound(sInArray)
        sInArray(i) = Trim$(sInArray(i))
        sInArray(i) = Replace(sInArray(i), vbCrLf, vbNullString)
        sInArray(i) = Replace(sInArray(i), vbTab, vbNullString)
    Next

End Sub

Private Sub TimeOut()

    ' time out occured, add the 'Timeout' error
    ' to the error collection
    AddError ERR_TIMEOUT
    SendFail

End Sub

Private Function UUEncodeString(ByRef str2UUEncode As String) As String

' ******************************************************************************
'
' Synopsis:     UUEncode a string
'
' Parameters:   str2UUEncode  - The input string
'
' Return:       encoded string
'
' Description:
' UU Encode a string. Coerce 3 bytes into 4 by converting 3, 8 bit bytes into
' 4, 6 bit values. Each 6 bit value (0 to 63) is then used as a pointer into
' the UUEncode string array to derive the correct character. The string will
' be a multiple of 4 bytes in lenght after conversion, padded with '=' as
' required. The line length will be encoded as a leading character
' (same 0 to 63 encoding) in the return string.
'
' ******************************************************************************

    Dim sValue              As String
    Dim lPtr                As Long
    Dim lCtr                As Long
    Dim lLen                As Long
    Dim lLineLen            As Long
    Dim sEncoded            As String
    Dim Bits8(1 To 3)       As Byte
    Dim Bits6(1 To 4)       As Byte

    lLineLen = Len(str2UUEncode)
    ' lines are limited to 63
    If lLineLen > 63 Then Exit Function

    For lCtr = 1 To Len(str2UUEncode) Step 3

        ' Get 3 characters
        sValue = Mid$(str2UUEncode, lCtr, 3)
        lLen = Len(sValue)

        ' Move string data into a byte array, then
        ' swap bits to create 4, 6 bit values (0-63)
        If lLen < 3 Then Erase Bits8
        CopyMemory Bits8(1), ByVal sValue, lLen
        Bits6(1) = (Bits8(1) And &HFC) \ &H4
        Bits6(2) = (Bits8(1) And &H3) * &H10 + (Bits8(2) And &HF0) \ &H10
        Bits6(3) = (Bits8(2) And &HF) * &H4 + (Bits8(3) And &HC0) \ &H40
        Bits6(4) = Bits8(3) And &H3F

        ' Encode new 4 byte string by selecting a character from
        ' the array. Length is determined by 'lLen' to make sure
        ' the file attachment is the right length
        For lPtr = 1 To lLen + 1
            sEncoded = sEncoded & psUUEncodeChr(Bits6(lPtr))
        Next

    Next

    ' add the line length character
    sEncoded = psUUEncodeChr(lLineLen) & sEncoded

    ' The decoder expects the size to be a multiple of 4 bytes.
    ' Possible sizes for the last packet are: 2, 3 & 4.
    Select Case lLen + 1
        Case 2: sEncoded = sEncoded & "=="          ' send two pad characters
        Case 3: sEncoded = sEncoded & "="           ' send one pad character
    ' no pad characers needed
    End Select

    UUEncodeString = sEncoded

End Function

Private Sub ValidateAddress(ByVal sRecip As String, ByVal sError As String)
    
    ' Validate Recipient, Cc and Bcc email address
    ' Appropriate validation methods for are:
    ' VALIDATE_NONE, VALIDATE_SYNTAX

    Dim iPtr            As Integer
    Dim sRecipArray()   As String

    RemoveError sError

    ' if VALIDATE_SYNTAX...
    If etEmailValidation = VALIDATE_SYNTAX Then
        ' split components into an array
        sRecipArray = Split(sRecip, psDelimiter)
        For iPtr = 0 To UBound(sRecipArray)
            ' validate address...
            If IsValidEmailAddress(sRecipArray(iPtr)) = False Then
                AddError sError
                Exit For
            End If
        Next
    End If

End Sub

Private Function WaitUntilTrue(ByRef Flag As Boolean, ByVal SecondsToWait As Long, Optional ByVal bRaiseTimeOutError As Boolean = True) As Boolean

    Dim fStart              As Single
    Dim fTimetoQuit         As Single

    If SecondsToWait < 1 Then Exit Function

    fStart = Timer

    ' Deal with timer being reset at Midnight
    If fStart + SecondsToWait < 86400 Then
        fTimetoQuit = fStart + SecondsToWait
    Else
        fTimetoQuit = (fStart - 86400) + SecondsToWait
    End If

    Do Until Flag = True
        If Timer >= fTimetoQuit Then
            If bRaiseTimeOutError Then TimeOut
            Exit Function
        End If
        If pbExitImmediately Then Exit Function
        DoEvents
        Sleep (10)                                  ' added to reduce CPU load during wait periods
    Loop

    WaitUntilTrue = Flag

End Function


' ******************************************************************************
' *      Private Winsock OCX Events                                            *
' ******************************************************************************

Private Sub sckMail_Close()
    'Close Socket
    ' keep track of connection state
    Layer = 0
    'sckMail.Close
    pbConnected = False
End Sub

Private Sub sckMail_Connect()

    ' keep track of connection state
    pbConnected = True

End Sub

Private Sub sckMail_DataArrival(ByVal bytesTotal As Long)

' ********************************************************
' SMTP Reply codes, outlined in RFC 821
' ********************************************************
' 211 - System status/help reply
' 214 - Help message
' 220 - <domain> Service ready
' 221 - <domain> Service closing channel
' 250 - OK: action completed
' 251 - User not local, will forward to <domain>
' 354 - OK: Start mail input, end with <CrLf>.<CrLf>
' 421 - <domain> Service not available, closing channel
' 450 - Mailbox busy, action not taken
' 451 - Requested action aborted: error in processing
' 452 - Requested action not taken: insufficient system storage
' 500 - Syntax error, command unrecognized
' 501 - Syntax error in parameters or arguments
' 502 - Command not implimented
' 503 - Bad sequence of commands
' 504 - Command parameter not implimented
' 550 - Mailbox unavailable, action not taken
' 553 - Requested action not taken: mailbox name not allowed / invalid
' 551 - User not local, please try <forward-path>
' 552 - Requested action not taken: exceeds storage allocation
' 554 - Transaction failed

' ********************************************************
' ESMTP AUTHentication extensions, outlined in RFC 2554
' ********************************************************
' 235 - Authentication successful
' 334 - Server challenge / ready response
' 432 - A password transition is needed
' 454 - Temporary authentication failure
' 530 - Authentication required
' 534 - Authentication mechanism is too weak
' 535 - Server rejected authentication
' 538 - Encryption required for requested authentication mechanism

' ********************************************************
'  POP3 Command Summary, outlined in RFC 1939
' ********************************************************
'  USER name               valid in the AUTHORIZATION state
'  PASS string
'  QUIT
'
'  STAT                    valid in the TRANSACTION state
'  List [msg]
'  RETR msg
'  DELE msg
'  NOOP
'  RSET
'  QUIT
'
'  Optional POP3 Commands:
'  APOP name digest        valid in the AUTHORIZATION state
'  TOP msg n               valid in the TRANSACTION state
'  UIDL [msg]
'
'  POP3 Replies:
'  +OK
'  -ERR
'
'  Note that with the exception of the STAT, LIST, and UIDL commands,
'  the reply given by the POP3 server to any command is significant
'  only to "+OK" and "-ERR".  Any text occurring after this reply
'  may be ignored by the client.

   
    Dim strAns          As String
    Dim sMsg            As String

    On Local Error Resume Next

    If sckMail.State <> sckConnected Then Exit Sub

    sckMail.GetData strAns, vbString

    Debug.Print strAns

    Select Case Left$(strAns, 3)

        ' Ready
        Case "220"
            pbRequestAccepted = True

        ' OK
        Case "221", "251"
            pbRequestAccepted = True


        ' OK, check for authentication support
        Case "250"
            pbRequestAccepted = True
            If InStr(1, strAns, "auth", vbTextCompare) > 0 And InStr(1, strAns, "login", vbTextCompare) > 0 Then
                pbAuthLoginSupported = True
            End If
            
            
        ' Auth Login OK
        Case "235"
            pbAuthLoginSuccess = True

        ' mail host 'AUTH' challenge
        Case "334"

            ' clean up the message portion
            sMsg = Trim$(Mid$(strAns, 4))
            sMsg = Replace(sMsg, vbCrLf, vbNullString)
            
            ' username requested
            If InStr(1, DecodeBase64String(sMsg), "username", vbTextCompare) Then
                sckMail.SendData EncodeBase64String(psUserName)

            ' password requested
            ElseIf InStr(1, DecodeBase64String(sMsg), "password", vbTextCompare) Then
                sckMail.SendData EncodeBase64String(psPassword)

            ' unexpected or unsupported challenge, cancel Auth request
            ' which will result in a 501 error reply from the host
            Else
                sckMail.SendData vbCrLf & "*" & vbCrLf
            End If

        ' OK, send data
        Case "354"
            pbDataOK = True

        ' do nothing
        Case "211", "214"

        ' POP3 success
        Case "+OK"
            Select Case plPop3Status
                Case 0
                    sckMail.SendData "USER " & psUserName & vbCrLf
                    plPop3Status = plPop3Status + 1
                
                Case 1
                    sckMail.SendData "PASS " & psPassword & vbCrLf
                    plPop3Status = plPop3Status + 1
                
                Case 2
                     sckMail.SendData "QUIT" & vbCrLf
                     plPop3Status = plPop3Status + 1
                
                Case 3
                     pbPopAuthOk = True
                     plPop3Status = 0
                     
            End Select
            
            
        ' POP3 error
        Case "-ER"
            pbPopAuthOk = False
            plPop3Status = 0
            pbExitImmediately = True
            AddError strAns
            SendFail
        
        ' host didn't like what we sent or couldn't process it
        Case Else
            AddError strAns  ''sMsg
            SendFail

    End Select

End Sub
Private Sub sckMail_Error(ByVal Number As Integer, Description As String, _
        ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, _
        ByVal HelpContext As Long, CancelDisplay As Boolean)

    ' socket error, add the error to the error collection
    AddError Description
    SendFail

    'MsgBox Number & " " & Description, vbCritical
    'SocketSSL_Close
End Sub
Private Sub sckMail_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

    Dim lNewValue           As Long
    Static lProgressLast    As Long

    pbSendProgress = True
    
    ' add up sent bytes
    plBytesSent = plBytesSent + bytesSent

    ' calculate the percentage of the total
    If plMailSize > 0 Then lNewValue = CLng(CSng(plBytesSent / plMailSize) * 100)
    If lNewValue > 100 Then lNewValue = 100

    ' update if the value changed
    If lNewValue <> lProgressLast Then
        lProgressLast = lNewValue
        If sckMail.State = sckConnected Then RaiseEvent Progress(lNewValue)
    End If

    ' keep track of what's left
    plBytesRemaining = bytesRemaining

End Sub
'=======================================================
'=======================================================
'=======================================================
'=======================================================
Private Function GetDomainName(ByVal pstrFromAddress As String) As String
    Dim pintInstr As Integer

    pintInstr = InStr(pstrFromAddress, "@")
    GetDomainName = Mid(pstrFromAddress, pintInstr + 1)
End Function

Private Function CheckAddressFrom(pstrAddy) As String
    Dim pintC As Integer

    If Left(pstrAddy, 1) <> "<" Then
        pstrAddy = "<" & pstrAddy
    End If

    If Right(pstrAddy, 1) <> ">" Then
        pstrAddy = pstrAddy & ">"
    End If

    CheckAddressFrom = pstrAddy
End Function
Private Function GetMIME(MIMEBoundary As String) As String
   Dim strMIMEMessageContent As String
   Dim strMIMEVersion As String
   Dim strMIMEContent As String

   strMIMEContent = "Content-Type: multipart/mixed; " _
                     & "boundary=" & """" & MIMEBoundary & """"

   strMIMEMessageContent = "--" & MIMEBoundary & vbCrLf & _
                            "Content-Type: text/plain; " & "charset=" & """" & "iso-8859-1" & """" & vbCrLf & _
                            "Content-Transfer-Encoding: 8bit"
   GetMIME = strMIMEContent & vbCrLf & vbCrLf & vbCrLf & strMIMEMessageContent
End Function



'=======================================================
'=======================================================
'=======================================================
'=======================================================
Private Sub SocketSend(Socket As Winsock, Data As String)
    If mvarSSL Then
        Call SSLSend(Socket, Data)
    Else
        Socket.SendData Data
    End If
End Sub
Private Function SendSequence() As String
    'Convert Send Counter to a String
    Dim TempString As String
    Dim TempSequence As Double
    Dim TempByte As Double
    Dim i As Integer

    TempSequence = SEND_SEQUENCE_NUMBER

    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next

    SendSequence = TempString

End Function
Private Sub SSLSend(ByRef Socket As Winsock, ByVal Plaintext As String)

    'Send Plaintext as an Encrypted SSL Record
    Dim SSLRecord As String
    Dim OtherPart As String
    Dim SendAnother As Boolean

    If Len(Plaintext) > 32751 Then
        SendAnother = True
        Plaintext = Mid(Plaintext, 1, 32751)
        OtherPart = Mid(Plaintext, 32752)
    Else
        SendAnother = False
    End If

    SSLRecord = AddMACData(Plaintext)
    SSLRecord = RC4_Encrypt(SSLRecord)
    SSLRecord = AddRecordHeader(SSLRecord)

    Socket.SendData SSLRecord

    If SendAnother = True Then
        Call SSLSend(Socket, OtherPart)
    End If

End Sub
Private Function AddMACData(ByVal Plaintext As String) As String
    'Prepend MAC Data to the Plaintext
    AddMACData = MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext
End Function
'
Private Function AddRecordHeader(ByVal RecordData As String) As String

    'Prepend SLL Record Header to the Data Record
    Dim FirstChar As String
    Dim LastChar As String
    Dim TheLen As Long

    TheLen = Len(RecordData)

    FirstChar = Chr(128 + (TheLen \ 256))
    LastChar = Chr(TheLen Mod 256)

    AddRecordHeader = FirstChar & LastChar & RecordData

    Call IncrementSend

End Function

Private Sub IncrementSend()

    'Increment Counter for Each Record Sent
    SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
    If SEND_SEQUENCE_NUMBER = 4294967296# Then SEND_SEQUENCE_NUMBER = 0

End Sub
Private Function RC4_Encrypt(ByVal Plaintext As String) As String

    'Encrypt with Client Write Key
    Dim lngLength As Long
    Dim lngReturnValue As Long

    lngLength = Len(Plaintext)
    lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)

    RC4_Encrypt = Plaintext

End Function
Private Function MD5_Hash(ByVal TheString As String) As String

    'Digest a String using MD5
    Dim lngReturnValue As Long
    Dim strHash As String
    Dim hHash As Long
    Dim lngHashLen As Long

    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNullString, lngHashLen, 0)
    strHash = String(lngHashLen, vbNullChar)
    lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)

    If hHash <> 0 Then CryptDestroyHash hHash

    MD5_Hash = strHash

End Function
'================================================================================================
'================================================================================================
'================================================================================================
'================================================================================================
'================================================================================================

Private Function ProcessHeader(sTo As String, sFrom As String, sCC As String, sSubject As String) As String

    Dim strMIMEMessage          As String
    Dim strMIMEContent          As String
    Dim strMIMEMessageContent   As String
    Dim strMIMEClient           As String
    Dim strMIMEMailInfo         As String
    Dim strMIMEVersion          As String

    strMIMEMailInfo = "DATE: " & Format(Now, "dd mmm yy ttttt") & vbCrLf & _
                      "FROM: " & Trim$(sFrom) & vbCrLf & _
                      "TO: " & Trim$(sTo) & vbCrLf & _
                      "CC: " & Trim$(sCC) & vbCrLf & _
                      "SUBJECT: " & sSubject

    strMIMEClient = "X-Mailer: " & App.ProductName & vbCrLf & _
                    "X-Version: " & App.Major & "." & App.Minor & vbCrLf & _
                    "X-CompanyName: " & App.CompanyName

    strMIMEVersion = "MIME-Version: 1.0"

    ProcessHeader = strMIMEMailInfo & vbCrLf & _
                    strMIMEVersion & vbCrLf & _
                    strMIMEClient



End Function


Private Function GetTo(pstrGroup() As String) As String
    Dim pintC

    For pintC = 0 To UBound(pstrGroup)
        GetTo = GetTo & ", " & pstrGroup(pintC)
    Next pintC

    GetTo = Mid(GetTo, 3)
End Function


Private Sub CreateDocToSend()
    Dim pstrHeader As String
    Dim pstrTMPFile As String
    Dim pintFF As Integer
    Dim pstrMIMEInfo As String
    Dim pstrBuffer As String
    Dim pstrTo As String
    Dim pstrCC As String
    Dim pstrFrom As String
    Dim pstrTotal As String
Dim FileCode
    pstrTo = Join(utMail.sToAddr, psDelimiter) 'pbReceipt 'GetTo(gstrToAddress)
    pstrCC = GetTo(gstrCC)
    If pstrTo = "" Then
        pstrTo = "Undisclosed"
    Else
        pstrTo = Replace(pstrTo, "<", "")
        pstrTo = Replace(pstrTo, ">", "")
    End If

    If pstrCC <> "" Then
        pstrCC = Replace(pstrCC, "<", "")
        pstrCC = Replace(pstrCC, ">", "")
    End If

    pstrFrom = gstrFromAddress
    pstrHeader = ProcessHeader(pstrTo, Chr(34) & gstrFromName & Chr(34) & " " & pstrFrom, pstrCC, gstrSubject)
    pstrTMPFile = GetTmpFile
    pintFF = FreeFile
    pstrHeader = pstrHeader & vbCrLf

    If gblnAttachments = True Then
        pstrMIMEInfo = GetMIME(conBoundary)
        pstrHeader = pstrHeader & pstrMIMEInfo & vbCrLf
    End If

    ''WRITE
    Open pstrTMPFile For Output As #pintFF
        Print #pintFF, pstrHeader
        Print #pintFF, gstrMessage
        If gblnAttachments = True Then
            Dim pintC As Integer
            For pintC = 0 To UBound(FileCode)
                If FileCode(pintC).Send = True Then
                    Print #pintFF, FileCode(pintC).Code
                    If pintC = UBound(FileCode) Then
                        Print #pintFF, "--" & conBoundary & "--"
                    End If
                End If
            Next pintC
        End If

    Close #pintFF
    gstrMessageFileName = pstrTMPFile
End Sub


Private Sub CheckAddress()
    Dim pintC As Integer

    On Error Resume Next

    For pintC = 0 To UBound(gstrToAddress)
        If Left(gstrToAddress(pintC), 1) <> "<" Then
            gstrToAddress(pintC) = "<" & gstrToAddress(pintC)
        End If

        If Right(gstrToAddress(pintC), 1) <> ">" Then
            gstrToAddress(pintC) = gstrToAddress(pintC) & ">"
        End If
    Next pintC

    For pintC = 0 To UBound(gstrCC)
        If Left(gstrCC(pintC), 1) <> "<" Then
            gstrCC(pintC) = "<" & gstrCC(pintC)
        End If

        If Right(gstrCC(pintC), 1) <> ">" Then
            gstrCC(pintC) = gstrCC(pintC) & ">"
        End If
    Next pintC

    For pintC = 0 To UBound(gstrBCC)
        If Left(gstrBCC(pintC), 1) <> "<" Then
            gstrBCC(pintC) = "<" & gstrBCC(pintC)
        End If

        If Right(gstrBCC(pintC), 1) <> ">" Then
            gstrBCC(pintC) = gstrBCC(pintC) & ">"
        End If
    Next pintC

End Sub


Private Function CheckConnectionPing() As Boolean
   Dim sTmp As String
   Dim hInet As Long
   Dim hUrl As Long
   Dim Flags As Long
   Dim url As Variant
   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
   If hInet Then
      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
      hUrl = InternetOpenUrl(hInet, "http://www.yahoo.com", vbNullString, 0, Flags, 0)
      If hUrl Then
         CheckConnectionPing = True
         Call InternetCloseHandle(hUrl)
    Else
         hUrl = InternetOpenUrl(hInet, "http://www.google.com", vbNullString, 0, Flags, 0)
         If hUrl Then
            CheckConnectionPing = True
            Call InternetCloseHandle(hUrl)
         End If
      End If
   End If
   Call InternetCloseHandle(hInet)
End Function


Private Function StartConnection() As Long
    Dim pintRes As Integer
    pintRes = InternetDial(0, "", INTERNET_DIAL_FORCE_PROMPT, 0, 0)
    StartConnection = pintRes
End Function
Public Function ConnectToInternet() As Long
    ConnectToInternet = StartConnection
End Function


Public Function SendEmail() As Integer
'(FromName As String
', Subject As String
', Message As String
', arrToAddresses() As String
', arrBCCAddresses() As String
', arrCCAddresses() As String
', WSock As Winsock
', RichText As RichTextBox
', Optional ProgsBar As ProgressBar) As Integer
    Dim pintC As Integer
    Dim pintRec As Integer
    Dim pblnConnected As Boolean

    pblnConnected = CheckConnectionPing
    If pblnConnected = False Then
        SendEmail = 1
        Exit Function
    End If

    Set SocketSSL = sckMail
    'Set RTBX = RichText

    gstrFromName = Me.FromDisplayName
    gstrSubject = Me.Subject
    gstrMessage = Me.Message
    gstrToAddress(0) = Me.Recipient
    gstrCC(0) = Me.CcRecipient
    gstrBCC(0) = Me.BCcRecipient
    gblnDone = False
    pbConnected = False
    gintTransfer = 0
    gintSent = 0

    Call CheckAddress

    For pintC = 0 To UBound(gstrCC)
        If gstrCC(pintC) = "" Or gstrCC(pintC) = "<>" Then Exit For
        ReDim Preserve gstrAllAddresses(pintRec)
        gstrAllAddresses(pintRec) = gstrCC(pintC)
        pintRec = pintRec + 1
    Next pintC

    For pintC = 0 To UBound(gstrToAddress)
        If gstrToAddress(pintC) = "" Or gstrToAddress(pintC) = "<>" Then Exit For
        ReDim Preserve gstrAllAddresses(pintRec)
        gstrAllAddresses(pintRec) = gstrToAddress(pintC)
        pintRec = pintRec + 1
    Next pintC

    On Error GoTo AllAddressError
    If gstrAllAddresses(0) = "" Then
        ReDim gstrAllAddresses(0)
    End If

    On Error Resume Next

    Call CreateDocToSend

    If gblnSSL = True Then
        Call ConnectSSLSocket
    Else
'        Here you can add a not SSL eMail procedure
'        Dim Soc As New RegSocket
'        Call Soc.ConnectRegSocket
    End If

    Kill gstrMessageFileName
    Call Reset
    
    SendEmail = 0
    Exit Function

AllAddressError:
    If Err.Number = 9 Then
        ReDim gstrAllAddresses(0)
        Resume Next
    End If

End Function


Public Sub SetUp(ServerName As String, Port As Long, FromAddress As String, SSL As Boolean, UserID As String, UserPass As String)

    gstrServerName = ServerName
    glngPort = Port
    gblnSSL = SSL
    gstrFromAddress = CheckAddressFrom(FromAddress)
    gstrDomainName = GetDomainName(gstrFromAddress)
    gstrUserID = UserID
    gstrUserPass = UserPass

End Sub
Private Function SelectExt(ByVal vsFullPathname As String) As String

    Dim extension As String
    extension = StrReverse(LCase$(Left(StrReverse(vsFullPathname), InStr(1, StrReverse(vsFullPathname), "."))))
    If Len(extension) = 0 Then extension = vsFullPathname

    Select Case extension

        Case ".323"
            SelectExt = "text/h323"

        Case ".aab"
            SelectExt = "application/x-authorware-bin"

        Case ".aam"
            SelectExt = "application/x-authorware-map"

        Case ".ace"
            SelectExt = "application/x-compressed"

        Case ".acp"
            SelectExt = "audio/x-mei-aac"

        Case ".ai"
            SelectExt = "application/postscript"

        Case ".aif"
            SelectExt = "audio/aiff"

        Case ".aifc"
            SelectExt = "audio/aiff"

        Case ".aiff"
            SelectExt = "audio/aiff"

        Case ".aip"
            SelectExt = "text/x-audiosoft-intra"

        Case ".ARJ"
            SelectExt = "application/x-compressed"

        Case ".art"
            SelectExt = "image/x-jg"

        Case ".asa"
            SelectExt = "text/asa"

        Case ".asf"
            SelectExt = "video/x-ms-asf"

        Case ".asp"
            SelectExt = "text/asp"

        Case ".asx"
            SelectExt = "video/x-ms-asf"

        Case ".asx"
            SelectExt = "video/x-ms-asx"

        Case ".au"
            SelectExt = "audio/basic"

        Case ".aut"
            SelectExt = "application/pbautomation"

        Case ".avi"
            SelectExt = "video/avi"

        Case ".avi"
            SelectExt = "video/x-msvideo"

        Case ".bmo"
            SelectExt = "audio/blue-matter-offer"

        Case ".bmp"
            SelectExt = "image/bmp"

        Case ".bmp"
            SelectExt = "image/x-bmp"

        Case ".bmr"
            SelectExt = "text/blue-matter-content-ref"

        Case ".bmt"
            SelectExt = "audio/blue-matter-song"

        Case ".bub"
            SelectExt = "application/photobubble"

        Case ".cat"
            SelectExt = "application/vnd.ms-pki.seccat"

        Case ".cdf"
            SelectExt = "application/x-cdf"

        Case ".cel"
            SelectExt = "video/flc"

        Case ".cer"
            SelectExt = "application/pkix-cert"

        Case ".cer"
            SelectExt = "application/x-x509-ca-cert"

        Case ".class"
            SelectExt = "java/*"

        Case ".crl"
            SelectExt = "application/pkix-crl"

        Case ".crt"
            SelectExt = "application/pkix-cert"

        Case ".crt"
            SelectExt = "application/x-x509-ca-cert"

        Case ".css"
            SelectExt = "text/css"

        Case ".dcr"
            SelectExt = "application/x-director"

        Case ".der"
            SelectExt = "application/pkix-cert"

        Case ".der"
            SelectExt = "application/x-x509-ca-cert"

        Case ".dib"
            SelectExt = "image/bmp"

        Case ".dib"
            SelectExt = "image/x-bmp"

        Case ".dif"
            SelectExt = "video/x-dv"

        Case ".dir"
            SelectExt = "application/x-director"

        Case ".dll"
            SelectExt = "application/x-msdownload"

        Case ".doc"
            SelectExt = "application/msword"

        Case ".dot"
            SelectExt = "application/msword"

        Case ".dpg"
            SelectExt = "application/vnd.dpgraph"

        Case ".dpgraph"
            SelectExt = "application/vnd.dpgraph"

        Case ".dv"
            SelectExt = "video/x-dv"

        Case ".dxr"
            SelectExt = "application/x-director"

        Case ".eml"
            SelectExt = "message/rfc822"

        Case ".emm"
            SelectExt = "application/x-emms-content"

        Case ".eps"
            SelectExt = "application/postscript"

        Case ".exe"
            SelectExt = "application/x-msdownload"

        Case ".fdf"
            SelectExt = "application/vnd.fdf"

        Case ".fif"
            SelectExt = "application/fractals"

        Case ".flc"
            SelectExt = "video/flc"

        Case ".fli"
            SelectExt = "video/flc"

        Case ".fml"
            SelectExt = "application/file-mirror-list"

        Case ".fpx"
            SelectExt = "image/x-xbitmap"

        Case ".gif"
            SelectExt = "image/gif"

        Case ".grv"
            SelectExt = "application/vnd.groove-injector"

        Case ".gz"
            SelectExt = "application/x-compressed"

        Case ".gz"
            SelectExt = "application/x-gzip"

        Case ".hpf"
            SelectExt = "application/x-icq-hpf"

        Case ".hqx"
            SelectExt = "application/mac-binhex40"

        Case ".hta"
            SelectExt = "application/hta"

        Case ".htc"
            SelectExt = "text/x-component"

        Case ".htm"
            SelectExt = "text/html"

        Case ".html"
            SelectExt = "text/html"

        Case ".htt"
            SelectExt = "text/webviewhtml"

        Case ".htx"
            SelectExt = "text/html"

        Case ".ico"
            SelectExt = "image/x-icon"

        Case ".iii"
            SelectExt = "application/x-iphone"

        Case ".ins"
            SelectExt = "application/x-internet-signup"

        Case ".ips"
            SelectExt = "application/x-ipscript"

        Case ".ipx"
            SelectExt = "application/x-ipix"

        Case ".isp"
            SelectExt = "application/x-internet-signup"

        Case ".IVF"
            SelectExt = "video/x-ivf"

        Case ".ivr"
            SelectExt = "i-world/i-vrml"

        Case ".java"
            SelectExt = "java/*"

        Case ".java"
            SelectExt = "text/java"

        Case ".jfif"
            SelectExt = "image/pjpeg"

        Case ".jpe"
            SelectExt = "image/jpeg"

        Case ".jpeg"
            SelectExt = "image/jpeg"

        Case ".jpg"
            SelectExt = "image/jpeg"

        Case ".JS"
            SelectExt = "application/x-javascript"

        Case ".la1"
            SelectExt = "audio/x-liquid-file"

        Case ".lar"
            SelectExt = "application/x-laplayer-reg"

        Case ".latex"
            SelectExt = "application/x-latex"

        Case ".lav"
            SelectExt = "audio/x-liquid"

        Case ".lavs"
            SelectExt = "audio/x-liquid-secure"

        Case ".lha"
            SelectExt = "application/x-compressed"

        Case ".lks"
            SelectExt = "application/x-lk-rlestream"

        Case ".lmsff"
            SelectExt = "audio/x-la-lms"

        Case ".lqt"
            SelectExt = "audio/x-liquid-file"

        Case ".ls"
            SelectExt = "application/x-javascript"

        Case ".lsf"
            SelectExt = "video/x-la-asf"

        Case ".lsx"
            SelectExt = "video/x-la-asf"

        Case ".LZH"
            SelectExt = "application/x-compressed"

        Case ".m1v"
            SelectExt = "video/mpeg"

        Case ".m3u"
            SelectExt = "audio/mpegurl"

        Case ".m3u"
            SelectExt = "audio/x-mpegurl"

        Case ".mac"
            SelectExt = "image/x-macpaint"

        Case ".man"
            SelectExt = "application/x-troff-man"

        Case ".mbc"
            SelectExt = "application/x-pn-virtualink"

        Case ".mbo"
            SelectExt = "application/x-previewsystems-vbox-music"

        Case ".mbox"
            SelectExt = "application/x-previewsystems-vbox-music"

        Case ".mdb"
            SelectExt = "application/msaccess"

        Case ".med"
            SelectExt = "application/x-att-a2bmusic-purchase"

        Case ".mes"
            SelectExt = "application/x-att-a2bmusic"

        Case ".mht"
            SelectExt = "message/rfc822"

        Case ".mhtml"
            SelectExt = "message/rfc822"

        Case ".mid"
            SelectExt = "audio/mid"

        Case ".midi"
            SelectExt = "audio/mid"

        Case ".mix"
            SelectExt = "image/x-xbitmap"

        Case ".mjf"
            SelectExt = "audio/x-vnd.AudioExplosion.MjuiceMediaFile"

        Case ".mjv"
            SelectExt = "audio/audio/mjuice_voucher"

        Case ".mmjb_mime"
            SelectExt = "application/x-musicmatch-mmjb5.20detect"

        Case ".mmz"
            SelectExt = "application/x-mmjb-mmz"

        Case ".mocha"
            SelectExt = "application/x-javascript"

        Case ".mov"
            SelectExt = "video/quicktime"

        Case ".movie"
            SelectExt = "video/x-sgi-movie"

        Case ".mp1"
            SelectExt = "audio/mpeg"

        Case ".mp2"
            SelectExt = "video/mpeg"

        Case ".mp2v"
            SelectExt = "video/mpeg"

        Case ".mp3"
            SelectExt = "audio/mpeg"

        Case ".mpa"
            SelectExt = "video/mpeg"

        Case ".mpe"
            SelectExt = "video/mpeg"

        Case ".mpeg"
            SelectExt = "video/mpeg"

        Case ".mpg"
            SelectExt = "video/mpeg"

        Case ".mpga"
            SelectExt = "audio/mpeg"

        Case ".mpv"
            SelectExt = "video/mpg"

        Case ".mpv2"
            SelectExt = "video/mpeg"

        Case ".mwc"
            SelectExt = "application/vnd.dpgraph"

        Case ".mxp"
            SelectExt = "application/x-mmxp"

        Case ".npi"
            SelectExt = "application/x-pn-npistream"

        Case ".nws"
            SelectExt = "message/rfc822"

        Case ".p10"
            SelectExt = "application/pkcs10"

        Case ".p12"
            SelectExt = "application/x-pkcs12"

        Case ".p7b"
            SelectExt = "application/x-pkcs7-certificates"

        Case ".p7c"
            SelectExt = "application/pkcs7-mime"

        Case ".p7m"
            SelectExt = "application/pkcs7-mime"

        Case ".p7r"
            SelectExt = "application/x-pkcs7-certreqresp"

        Case ".p7s"
            SelectExt = "application/pkcs7-signature"

        Case ".pct"
            SelectExt = "image/pict"

        Case ".pdf"
            SelectExt = "application/pdf"

        Case ".pfx"
            SelectExt = "application/x-pkcs12"

        Case ".pic"
            SelectExt = "image/pict"

        Case ".pict"
            SelectExt = "image/pict"

        Case ".pko"
            SelectExt = "application/vnd.ms-pki.pko"

        Case ".pl"
            SelectExt = "application/x-perl"

        Case ".plg"
            SelectExt = "text/html"

        Case ".pls"
            SelectExt = "audio/scpls"

        Case ".pls"
            SelectExt = "audio/x-scpls"

        Case ".png"
            SelectExt = "image/png"

        Case ".pnq"
            SelectExt = "application/x-icq-pnq"

        Case ".pntg"
            SelectExt = "image/x-macpaint"

        Case ".POT"
            SelectExt = "application/vnd.ms-powerpoint"

        Case ".ppa"
            SelectExt = "application/vnd.ms-powerpoint"

        Case ".pps"
            SelectExt = "application/vnd.ms-powerpoint"

        Case ".ppt"
            SelectExt = "application/x-mspowerpoint"

        Case ".prf"
            SelectExt = "application/pics-rules"

        Case ".ps"
            SelectExt = "application/postscript"

        Case ".pwz"
            SelectExt = "application/vnd.ms-powerpoint"

        Case ".py"
            SelectExt = "text/plain"

        Case ".pyw"
            SelectExt = "text/plain"

        Case ".qht"
            SelectExt = "text/x-html-insertion"

        Case ".qhtm"
            SelectExt = "text/x-html-insertion"

        Case ".qt"
            SelectExt = "video/quicktime"

        Case ".qti"
            SelectExt = "image/x-quicktime"

        Case ".qtif"
            SelectExt = "image/x-quicktime"

        Case ".qtl"
            SelectExt = "application/x-quicktimeplayer"

        Case ".r00"
            SelectExt = "application/x-rar-compressed"

        Case ".r01"
            SelectExt = "application/x-rar-compressed"

        Case ".r02"
            SelectExt = "application/x-rar-compressed"

        Case ".r03"
            SelectExt = "application/x-rar-compressed"

        Case ".r04"
            SelectExt = "application/x-rar-compressed"

        Case ".r05"
            SelectExt = "application/x-rar-compressed"

        Case ".r06"
            SelectExt = "application/x-rar-compressed"

        Case ".r07"
            SelectExt = "application/x-rar-compressed"

        Case ".r08"
            SelectExt = "application/x-rar-compressed"

        Case ".r09"
            SelectExt = "application/x-rar-compressed"

        Case ".r10"
            SelectExt = "application/x-rar-compressed"

        Case ".r11"
            SelectExt = "application/x-rar-compressed"

        Case ".r12"
            SelectExt = "application/x-rar-compressed"

        Case ".r13"
            SelectExt = "application/x-rar-compressed"

        Case ".r14"
            SelectExt = "application/x-rar-compressed"

        Case ".r15"
            SelectExt = "application/x-rar-compressed"

        Case ".r16"
            SelectExt = "application/x-rar-compressed"

        Case ".r17"
            SelectExt = "application/x-rar-compressed"

        Case ".r18"
            SelectExt = "application/x-rar-compressed"

        Case ".r19"
            SelectExt = "application/x-rar-compressed"

        Case ".r20"
            SelectExt = "application/x-rar-compressed"

        Case ".r21"
            SelectExt = "application/x-rar-compressed"

        Case ".r22"
            SelectExt = "application/x-rar-compressed"

        Case ".r23"
            SelectExt = "application/x-rar-compressed"

        Case ".r24"
            SelectExt = "application/x-rar-compressed"

        Case ".r25"
            SelectExt = "application/x-rar-compressed"

        Case ".r26"
            SelectExt = "application/x-rar-compressed"

        Case ".r27"
            SelectExt = "application/x-rar-compressed"

        Case ".r28"
            SelectExt = "application/x-rar-compressed"

        Case ".r29"
            SelectExt = "application/x-rar-compressed"

        Case ".r30"
            SelectExt = "application/x-rar-compressed"

        Case ".r31"
            SelectExt = "application/x-rar-compressed"

        Case ".r32"
            SelectExt = "application/x-rar-compressed"

        Case ".r33"
            SelectExt = "application/x-rar-compressed"

        Case ".r34"
            SelectExt = "application/x-rar-compressed"

        Case ".r35"
            SelectExt = "application/x-rar-compressed"

        Case ".r36"
            SelectExt = "application/x-rar-compressed"

        Case ".r37"
            SelectExt = "application/x-rar-compressed"

        Case ".r38"
            SelectExt = "application/x-rar-compressed"

        Case ".r39"
            SelectExt = "application/x-rar-compressed"

        Case ".r3t"
            SelectExt = "text/vnd.rn-realtext3d"

        Case ".r40"
            SelectExt = "application/x-rar-compressed"

        Case ".r41"
            SelectExt = "application/x-rar-compressed"

        Case ".r42"
            SelectExt = "application/x-rar-compressed"

        Case ".r43"
            SelectExt = "application/x-rar-compressed"

        Case ".r44"
            SelectExt = "application/x-rar-compressed"

        Case ".r45"
            SelectExt = "application/x-rar-compressed"

        Case ".r46"
            SelectExt = "application/x-rar-compressed"

        Case ".r47"
            SelectExt = "application/x-rar-compressed"

        Case ".r48"
            SelectExt = "application/x-rar-compressed"

        Case ".r49"
            SelectExt = "application/x-rar-compressed"

        Case ".r50"
            SelectExt = "application/x-rar-compressed"

        Case ".r51"
            SelectExt = "application/x-rar-compressed"

        Case ".r52"
            SelectExt = "application/x-rar-compressed"

        Case ".r53"
            SelectExt = "application/x-rar-compressed"

        Case ".r54"
            SelectExt = "application/x-rar-compressed"

        Case ".r55"
            SelectExt = "application/x-rar-compressed"

        Case ".r56"
            SelectExt = "application/x-rar-compressed"

        Case ".r57"
            SelectExt = "application/x-rar-compressed"

        Case ".r58"
            SelectExt = "application/x-rar-compressed"

        Case ".r59"
            SelectExt = "application/x-rar-compressed"

        Case ".r60"
            SelectExt = "application/x-rar-compressed"

        Case ".r61"
            SelectExt = "application/x-rar-compressed"

        Case ".r62"
            SelectExt = "application/x-rar-compressed"

        Case ".r63"
            SelectExt = "application/x-rar-compressed"

        Case ".r64"
            SelectExt = "application/x-rar-compressed"

        Case ".r65"
            SelectExt = "application/x-rar-compressed"

        Case ".r66"
            SelectExt = "application/x-rar-compressed"

        Case ".r67"
            SelectExt = "application/x-rar-compressed"

        Case ".r68"
            SelectExt = "application/x-rar-compressed"

        Case ".r69"
            SelectExt = "application/x-rar-compressed"

        Case ".r70"
            SelectExt = "application/x-rar-compressed"

        Case ".r71"
            SelectExt = "application/x-rar-compressed"

        Case ".r72"
            SelectExt = "application/x-rar-compressed"

        Case ".r73"
            SelectExt = "application/x-rar-compressed"

        Case ".r74"
            SelectExt = "application/x-rar-compressed"

        Case ".r75"
            SelectExt = "application/x-rar-compressed"

        Case ".r76"
            SelectExt = "application/x-rar-compressed"

        Case ".r77"
            SelectExt = "application/x-rar-compressed"

        Case ".r78"
            SelectExt = "application/x-rar-compressed"

        Case ".r79"
            SelectExt = "application/x-rar-compressed"

        Case ".r80"
            SelectExt = "application/x-rar-compressed"

        Case ".r81"
            SelectExt = "application/x-rar-compressed"

        Case ".r82"
            SelectExt = "application/x-rar-compressed"

        Case ".r83"
            SelectExt = "application/x-rar-compressed"

        Case ".r84"
            SelectExt = "application/x-rar-compressed"

        Case ".r85"
            SelectExt = "application/x-rar-compressed"

        Case ".r86"
            SelectExt = "application/x-rar-compressed"

        Case ".r87"
            SelectExt = "application/x-rar-compressed"

        Case ".r88"
            SelectExt = "application/x-rar-compressed"

        Case ".r89"
            SelectExt = "application/x-rar-compressed"

        Case ".r90"
            SelectExt = "application/x-rar-compressed"

        Case ".r91"
            SelectExt = "application/x-rar-compressed"

        Case ".r92"
            SelectExt = "application/x-rar-compressed"

        Case ".r93"
            SelectExt = "application/x-rar-compressed"

        Case ".r94"
            SelectExt = "application/x-rar-compressed"

        Case ".r95"
            SelectExt = "application/x-rar-compressed"

        Case ".r96"
            SelectExt = "application/x-rar-compressed"

        Case ".r97"
            SelectExt = "application/x-rar-compressed"

        Case ".r98"
            SelectExt = "application/x-rar-compressed"

        Case ".r99"
            SelectExt = "application/x-rar-compressed"

        Case ".ra"
            SelectExt = "audio/vnd.rn-realaudio"

        Case ".ram"
            SelectExt = "audio/x-pn-realaudio"

        Case ".rar"
            SelectExt = "application/x-rar-compressed"

        Case ".rat"
            SelectExt = "application/rat-file"

        Case ".rf"
            SelectExt = "image/vnd.rn-realflash"

        Case ".rjs"
            SelectExt = "application/vnd.rn-realsystem-rjs"

        Case ".rjt"
            SelectExt = "application/vnd.rn-realsystem-rjt"

        Case ".rm"
            SelectExt = "application/vnd.rn-realmedia"

        Case ".rmi"
            SelectExt = "audio/mid"

        Case ".rmj"
            SelectExt = "application/vnd.rn-realsystem-rmj"

        Case ".rmm"
            SelectExt = "audio/x-pn-realaudio"

        Case ".rmp"
            SelectExt = "application/vnd.rn-rn_music_package"

        Case ".rmx"
            SelectExt = "application/vnd.rn-realsystem-rmx"

        Case ".rnx"
            SelectExt = "application/vnd.rn-realplayer"

        Case ".rob"
            SelectExt = "application/vnd.rn-objects"

        Case ".rp"
            SelectExt = "image/vnd.rn-realpix"

        Case ".rpm"
            SelectExt = "audio/x-pn-realaudio-plugin"

        Case ".rsml"
            SelectExt = "application/vnd.rn-rsml"

        Case ".rt"
            SelectExt = "text/vnd.rn-realtext"

        Case ".rtf"
            SelectExt = "application/msword"

        Case ".rtsp"
            SelectExt = "application/x-rtsp"

        Case ".rv"
            SelectExt = "video/vnd.rn-realvideo"

        Case ".sc"
            SelectExt = "application/vnd.optx-screenwatch"

        Case ".scm"
            SelectExt = "application/x-icq-scm"

        Case ".sct"
            SelectExt = "text/scriptlet"

        Case ".sd2"
            SelectExt = "audio/x-sd2"

        Case ".sdf"
            SelectExt = "application/x-server-launch"

        Case ".sdp"
            SelectExt = "application/sdp"

        Case ".sgi"
            SelectExt = "image/x-sgi"

        Case ".sit"
            SelectExt = "application/x-stuffit"

        Case ".sma"
            SelectExt = "application/x-smb-directive"

        Case ".smi"
            SelectExt = "application/smil"

        Case ".smil"
            SelectExt = "application/smil"

        Case ".snd"
            SelectExt = "audio/basic"

        Case ".spc"
            SelectExt = "application/x-pkcs7-certificates"

        Case ".spl"
            SelectExt = "application/futuresplash"

        Case ".spn"
            SelectExt = "application/vnd.spinnerplus"

        Case ".ssm"
            SelectExt = "application/streamingmedia"

        Case ".sst"
            SelectExt = "application/vnd.ms-pki.certstore"

        Case ".stl"
            SelectExt = "application/vnd.ms-pki.stl"

        Case ".stm"
            SelectExt = "text/html"

        Case ".svg"
            SelectExt = "image/svg+xml"

        Case ".svg"
            SelectExt = "image/svg-xml"

        Case ".svgz"
            SelectExt = "image/svg+xml"

        Case ".svgz"
            SelectExt = "image/svg-xml"

        Case ".swf"
            SelectExt = "application/x-shockwave-flash"

        Case ".tar"
            SelectExt = "application/x-compressed"

        Case ".tar"
            SelectExt = "application/x-tar"

        Case ".tga"
            SelectExt = "image/x-targa"

        Case ".tgz"
            SelectExt = "application/x-compressed"

        Case ".tif"
            SelectExt = "image/tiff"

        Case ".tiff"
            SelectExt = "image/tiff"

        Case ".txt"
            SelectExt = "text/plain"

        Case ".uin"
            SelectExt = "application/x-icq"

        Case ".uls"
            SelectExt = "text/iuls"

        Case ".ultact"
            SelectExt = "application/x-UltimateAction"

        Case ".ulw"
            SelectExt = "audio/basic"

        Case ".urls"
            SelectExt = "application/x-url-list"

        Case ".UU"
            SelectExt = "application/x-compressed"

        Case ".UUE"
            SelectExt = "application/x-compressed"

        Case ".vcf"
            SelectExt = "text/x-vcard"

        Case ".vcg"
            SelectExt = "application/vnd.groove-vcard"

        Case ".vcl"
            SelectExt = "text/html"

        Case ".vfw"
            SelectExt = "video/x-msvideo"

        Case ".vpg"
            SelectExt = "application/x-vpeg"

        Case ".vsl"
            SelectExt = "application/x-cnet-vsl"

        Case ".wav"
            SelectExt = "audio/wav"

        Case ".wax"
            SelectExt = "audio/x-ms-wax"

        Case ".wiz"
            SelectExt = "application/msword"

        Case ".wm"
            SelectExt = "video/x-ms-wm"

        Case ".wma"
            SelectExt = "audio/x-ms-wma"

        Case ".wmd"
            SelectExt = "application/x-ms-wmd"

        Case ".wme"
            SelectExt = "text/xml"

        Case ".wmp"
            SelectExt = "video/x-ms-wmp"

        Case ".wms"
            SelectExt = "application/x-ms-wms"

        Case ".wmv"
            SelectExt = "video/x-ms-wmv"

        Case ".wmx"
            SelectExt = "video/x-ms-wmx"

        Case ".wmz"
            SelectExt = "application/x-ms-wmz"

        Case ".wsc"
            SelectExt = "text/scriptlet"

        Case ".wvx"
            SelectExt = "video/x-ms-wvx"

        Case ".xbm"
            SelectExt = "image/x-xbitmap"

        Case ".xls"
            SelectExt = "application/vnd.ms-excel"

        Case ".xls"
            SelectExt = "application/x-msexcel"

        Case ".xml"
            SelectExt = "text/xml"

        Case ".xpl"
            SelectExt = "audio/mpegurl"

        Case ".xsl"
            SelectExt = "text/xml"

        Case ".XXE"
            SelectExt = "application/x-compressed"

        Case ".ymg"
            SelectExt = "application/ymsgr"

        Case ".yps"
            SelectExt = "application/ymsgr"

        Case ".z"
            SelectExt = "application/x-compress"

        Case ".zip"
            SelectExt = "application/x-zip-compressed"

        Case "ratfile"
            SelectExt = "application/rat-file"

        Case "smafile"
            SelectExt = "application/x-smb-directive"

        Case Else
            SelectExt = "application/octet-stream"
    End Select

End Function

Private Function MIMEFileHeader(MIMEBoundary As String, FilePath As String)
    Dim sResult As String
    Dim sFileName As String
    Dim pintC

    pintC = InStrRev(FilePath, "\")
    sFileName = Mid(FilePath, pintC + 1)

    sResult = vbCrLf & "--" & MIMEBoundary & vbCrLf
    sResult = sResult & "Content-Type: " & SelectExt(FilePath) & "; " & "name=" & Chr(34) & sFileName & Chr(34) & vbNewLine
    sResult = sResult & "Content-Transfer-Encoding: base64" & vbCrLf
    sResult = sResult & "Content-Disposition: attachment; filename=" & Chr(34) & sFileName & Chr(34) & vbCrLf

    MIMEFileHeader = sResult
End Function

Public Function Base64EncodeFile(sInputFile As String, sOutputFile As String, Optional PrgBar As Object) As Byte()
    Dim bTemp() As Byte
    Dim fh As Long

    fh = FreeFile(0)
    Open sInputFile For Binary Access Read As fh
        ReDim bTemp(0 To LOF(fh) - 1)
        Get fh, , bTemp
    Close fh

    bTemp = EncodeArr(bTemp, PrgBar)
    Open sOutputFile For Binary Access Write As fh
        Put fh, , bTemp
    Close fh

    Base64EncodeFile = bTemp
End Function

'Converts an array of bytes to a string.
Private Sub ByteArrayToString(ByteArray() As Byte, StringOut As String)
  Dim lBytes As Long

  If LBound(ByteArray) > 0 Then Exit Sub 'lBound MUST be 0
  lBytes = UBound(ByteArray) + 1
  StringOut = String$(lBytes, 0)

  RtlMoveMemory ByVal StringOut, ByteArray(0), lBytes
End Sub


'Converts a string to an array of bytes. Just like Asc() on every character.
Private Sub StringToByteArray(ByVal StringIn As String, ByteArray() As Byte)
    Dim lBytes As Long

    If Len(StringIn) = 0 Then Exit Sub
    lBytes = Len(StringIn)
    ReDim ByteArray(lBytes - 1)

    RtlMoveMemory ByteArray(0), ByVal StringIn, lBytes
End Sub


'Takes sets of 3 characters (binary) and returns sets of 4 characters (ASCII)
Private Function EncodeString(ByVal InString As String) As String
    Dim OutString As String
    Dim i As Integer
    Dim UnCodedArray() As Byte
    Dim CodedArray() As Byte

    'Pad will null characters if necessary
    If Len(InString) Mod 3 <> 0 Then
        InString = InString & String(3 - Len(InString) Mod 3, Chr$(0))
    End If

    'Convert string to a byte array. This is MUCH faster than the Asc/Chr combo.
    StringToByteArray InString, UnCodedArray()

    'Make sure our output array is the correct size
    ReDim CodedArray(((Len(InString) / 3) * 4) - 1)

    For i = 0 To (Len(InString) / 3) - 1
        'Encode 4 bytes at a time
        CodedArray(i * 4 + 0) = UnCodedArray(i * 3 + 0) \ 4 + 32
        CodedArray(i * 4 + 1) = ((UnCodedArray(i * 3 + 0) Mod 4) * 16) + (UnCodedArray(i * 3 + 1) \ 16 + 32)
        CodedArray(i * 4 + 2) = ((UnCodedArray(i * 3 + 1) Mod 16) * 4) + (UnCodedArray(i * 3 + 2) \ 64 + 32)
        CodedArray(i * 4 + 3) = (UnCodedArray(i * 3 + 2) Mod 64) + 32

        'Check for spaces and eliminate them
        If CodedArray(i * 4 + 0) = 32 Then CodedArray(i * 4 + 0) = 96
        If CodedArray(i * 4 + 1) = 32 Then CodedArray(i * 4 + 1) = 96
        If CodedArray(i * 4 + 2) = 32 Then CodedArray(i * 4 + 2) = 96
        If CodedArray(i * 4 + 3) = 32 Then CodedArray(i * 4 + 3) = 96
    Next i
    ByteArrayToString CodedArray(), OutString
    EncodeString = OutString
End Function


'Gives a character representing the number of bytes to be decoded from a line. "M" is a full line (45 bytes)
Private Function ENC(ByVal i As Integer) As String
    If i = 0 Then
        ENC = "`"
    Else
        i = i + 32
        ENC = Chr(i)
    End If
End Function


'Returns the filename part from a full filename.
Private Function StripFilename(ByVal Filename As String) As String
    While InStr(Filename, "\") <> 0
        Filename = Right(Filename, Len(Filename) - InStr(Filename, "\"))
    Wend
    StripFilename = Filename
End Function


Private Function UUEncodeFile(ByVal InputFilename As String, ByVal OutputFileName As String) As Boolean
    Dim TotalParts As Long
    Dim Remain As Long
    Dim InString As String
    Dim CurrentPos As Long
    Dim i As Long
    Dim FNumIn As Long
    Dim FNumOut As Long

    On Error GoTo ErrorHandler

    'Open the original file as binary read
    FNumIn = FreeFile
    Open InputFilename For Binary Access Read Shared As #FNumIn

    'Open the target file as binary write
    FNumOut = FreeFile
    Open OutputFileName For Binary Access Write As #FNumOut

    'Write the standard uuencode file header
    Put #FNumOut, , "begin 644 " + StripFilename(InputFilename) + vbCrLf

    'TotalParts equals the number of full sized chunks of "PortionSize" bytes.
    TotalParts = LOF(1) \ PortionSize

    'Remainder equals how many remaining bytes are at the end of the file.
    Remain = LOF(1) Mod PortionSize

    'CurrentPos is the current file position
    CurrentPos = 1

    'Set up InString as a buffer the size of PortionSize
    InString = String(PortionSize, 0)

    'For loop to read the portions one by one
    For i = 1 To TotalParts
        Get #FNumIn, CurrentPos, InString
        'use the ENC() for standard uuencode compatibility, pad "M"
        Put #FNumOut, , ENC(PortionSize) + EncodeString(InString) + vbCrLf
        CurrentPos = CurrentPos + PortionSize
        RaiseEvent EncodeProgress(CurrentPos / LOF(1), CurrentPos)
    Next

    'Set up InString again as a buffer the size of the remaining bytes
    InString = String(Remain, 0)

    'get the remaining bytes toward end of the file
    Get #FNumIn, CurrentPos, InString

    'get the remaining bytes size and calculate ENC() for the last line
    Put #FNumOut, , ENC(LOF(1) - CurrentPos + 1) + EncodeString(InString) + vbCrLf

    'put "end" for standard uuencode compatibility
    Put #FNumOut, , ENC(0) + vbCrLf + "end" + vbCrLf

    'All Done!
    Close #FNumIn
    Close #FNumOut
    UUEncodeFile = True
    Exit Function

ErrorHandler:
    Close #FNumIn
    Close #FNumOut
    UUEncodeFile = False
End Function
Private Function EncodeFile(sFileInput As String, dType As DecodeType, sFileOutput As String, Optional PrgBar As Object) As Byte()
    Dim pint As Integer
    Select Case dType
        Case uu
            Call UUEncodeFile(sFileInput, sFileOutput)
        Case base64
            EncodeFile = Base64EncodeFile(sFileInput, sFileOutput, PrgBar)
    End Select
End Function
Private Function GetTmpFile() As String
    Dim pstrTmpPath As String
    Dim pstrTmpFileName As String

    pstrTmpPath = String(100, Chr$(0))
    Call GetTempPath(100, pstrTmpPath)
    pstrTmpPath = Left$(pstrTmpPath, InStr(pstrTmpPath, Chr$(0)) - 1)
    pstrTmpFileName = String(260, 0)
    Call GetTempFileName(pstrTmpPath, "NPD", 0, pstrTmpFileName)

    GetTmpFile = Left$(pstrTmpFileName, InStr(pstrTmpFileName, Chr$(0)) - 1)

End Function


Public Sub AddAttachment(Filename As String, Optional PrgBar As Object)
    Static pintC As Integer
    Dim pintFF As Integer
    Dim pstrTMPFile As String
    Dim pstrBuffer() As Byte
    Dim pstrTotal As String
    Dim pstrHeader As String

    pintC = pintC + 1
    ReDim Preserve gstrFileNames(pintC)
    gstrFileNames(pintC) = Filename
    gblnAttachments = True
    pintFF = FreeFile
    ReDim Preserve FileCode(pintC)
    pstrTMPFile = GetTmpFile

    pstrBuffer = EncodeFile(gstrFileNames(pintC), base64, pstrTMPFile, PrgBar)


    pstrTotal = StrConv(pstrBuffer(), vbUnicode)

    pstrHeader = MIMEFileHeader(conBoundary, gstrFileNames(pintC))
    FileCode(pintC).Code = pstrHeader _
            & vbCrLf & pstrTotal & vbCrLf
    FileCode(pintC).Filename = gstrFileNames(pintC)
    FileCode(pintC).Send = True
    Kill pstrTMPFile

End Sub


Private Function EncodeArr(bInput() As Byte, Optional PrgBar As Object) As Byte()
Dim bOutput() As Byte
Dim k As Long
Dim l As Long
Dim i As Long
Dim evenBound As Long
Dim CurrentOut As Long
Dim b As Byte
Dim c As Byte
Dim d As Byte
Dim LineLength As Long

    k = LBound(bInput)
    l = UBound(bInput)
    'Calculate the input size
    i = l - k + 1


    'Calculate the input size
    i = l - k + 1

    'Calculate the output size
    Select Case i Mod 3
       Case 0:
            i = (i \ 3) * 4
            evenBound = l
       Case 1:
            i = ((i \ 3) * 4) + 4
            evenBound = l - 1
       Case 2:
            i = ((i \ 3) * 4) + 4
            evenBound = l - 2
       Case 3:
            i = ((i \ 3) * 4) + 4
            evenBound = l - 3
    End Select

    'Add in the line feeds.
    If i Mod MAX_LINELENGTH = 0 Then
        i = i + (i \ MAX_LINELENGTH) * 2 - 2
    Else
        i = i + (i \ MAX_LINELENGTH) * 2
    End If

    'Size the output array
    ReDim bOutput(0 To i - 1)

    CurrentOut = 0
    LineLength = 0

    On Error Resume Next
    Dim progStep As Single, progVal As Single
    PrgBar.Min = 0.0001!
    PrgBar.Max = 100.0001!
    PrgBar.Value = PrgBar.Min
    progStep = CSng(PrgBar.Max / evenBound)
    On Error GoTo 0

    For i = k To evenBound Step 3
        b = bInput(i)
        c = bInput(i + 1)
        d = bInput(i + 2)
        bOutput(CurrentOut) = m_Index1(b And &HFC)
        bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
        bOutput(CurrentOut + 2) = m_Index3((c And &HF) Or (d And &HC0))
        bOutput(CurrentOut + 3) = m_Index4(d And &H3F)
        CurrentOut = CurrentOut + 4
        LineLength = LineLength + 4

        If LineLength >= MAX_LINELENGTH Then
            If i <> l - 2 Then  ' If this is the last line, don't add crlf
                bOutput(CurrentOut) = CHAR_CR
                bOutput(CurrentOut + 1) = CHAR_LF
            End If
            CurrentOut = CurrentOut + 2
            LineLength = 0
        End If
        On Error Resume Next
        'Solve some Flickering in Win 7
        progVal = CSng(i * progStep) + 0.0001!
        If progVal <> PrgBar.Value And CSng(i Mod 1000.0001) = 0 Then
            PrgBar.Value = progVal
        End If
        On Error GoTo 0
    Next i

    Select Case l - i
       Case 1:
            b = bInput(i)
            c = bInput(i + 1)
            d = 0
            bOutput(CurrentOut) = m_Index1(b And &HFC)
            bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
            bOutput(CurrentOut + 2) = m_Index3((c And &HF) Or (d And &HC0))
            bOutput(CurrentOut + 3) = CHAR_EQUAL
            CurrentOut = CurrentOut + 4
            LineLength = LineLength + 4
       Case 0:
            b = bInput(i)
            c = 0
            bOutput(CurrentOut) = m_Index1(b And &HFC)
            bOutput(CurrentOut + 1) = m_Index2((b And &H3) Or (c And &HF0))
            bOutput(CurrentOut + 2) = CHAR_EQUAL
            bOutput(CurrentOut + 3) = CHAR_EQUAL
            CurrentOut = CurrentOut + 4
            LineLength = LineLength + 4
    End Select

    EncodeArr = bOutput

End Function



' Encode a string to a string.
Private Function Encode(sInput As String) As String
   Dim bTemp() As Byte

   'Convert to a byte array then convert.
   'This is faster the repetitive calls to asc() or chr$()
   bTemp = StrConv(sInput, vbFromUnicode)
   Encode = StrConv(EncodeArr(bTemp), vbUnicode)
End Function

Public Function EncodeStringBase64(pstrString) As String

    EncodeStringBase64 = Encode(Trim(pstrString))

End Function


Private Function Base64String(ByVal pstrCode As String) As String

    Base64String = Decode(Trim(pstrCode))

End Function

Private Function DecodeArr(bInput() As Byte) As Byte()
   Dim bOutput() As Byte
   Dim OutLength As Long
   Dim CurrentOut As Long

   Dim k As Long
   Dim l As Long
   Dim i As Long
   Dim j As Long

   Dim b As Byte
   Dim c As Byte
   Dim d As Byte
   Dim e As Byte

   On Error GoTo ErrorHandler

   k = LBound(bInput)
   l = UBound(bInput)

   'Calculate the length of the input
   i = l - k + 1

   'Calculate the expected length of the output
   'It should be no more (but may possible be less)
   j = i Mod (MAX_LINELENGTH + 2)
   If j = 0 Then
      OutLength = (i \ (MAX_LINELENGTH + 2)) * (MAX_LINELENGTH \ 4) * 3
   Else
      j = (j / 4) * 3
      If bInput(l) = CHAR_EQUAL Then j = j - 1
      If bInput(l - 1) = CHAR_EQUAL Then j = j - 1
      OutLength = (i \ (MAX_LINELENGTH + 2)) * (MAX_LINELENGTH \ 4) * 3 + j
   End If

   'Allocate the output
   ReDim bOutput(0 To OutLength - 1)

   CurrentOut = 0

   For i = k To l
      Select Case bInput(i)
         Case CHAR_CR
            'Do nothing
         Case CHAR_LF
            'Do nothing
         Case Else
            If l - i >= 3 Then
               b = bInput(i)
               c = bInput(i + 1)
               d = bInput(i + 2)
               e = bInput(i + 3)

               If e <> CHAR_EQUAL Then

                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                  bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
                  bOutput(CurrentOut + 2) = m_ReverseIndex3(d, 1) Or m_ReverseIndex4(e)
                  CurrentOut = CurrentOut + 3
                  i = i + 3
               ElseIf d <> CHAR_EQUAL Then
                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                  bOutput(CurrentOut + 1) = m_ReverseIndex2(c, 1) Or m_ReverseIndex3(d, 0)
                  CurrentOut = CurrentOut + 2
                  i = i + 3
               Else
                  bOutput(CurrentOut) = m_ReverseIndex1(b) Or m_ReverseIndex2(c, 0)
                  CurrentOut = CurrentOut + 1
                  i = i + 3
               End If

            Else
               'Possible input code error, but may also be
               'an extra CrLf, so we will ignore it.
            End If
      End Select
   Next i

   'On properly formed input we should have to do this.
   If OutLength <> CurrentOut + 1 Then
      ReDim Preserve bOutput(0 To CurrentOut - 1)
   End If

   DecodeArr = bOutput

   Exit Function
ErrorHandler:
    If Err.Number = 9 Then
        Exit Function
    End If
End Function



'Decode a string to a string.
Private Function Decode(sInput As String) As String
   Dim bTemp() As Byte

   'Convert to a byte array then convert.
   'This is faster the repetitive calls to asc() or chr$()
   bTemp = StrConv(sInput, vbFromUnicode)
   Decode = StrConv(DecodeArr(bTemp), vbUnicode)
End Function


Private Sub CertificateToPublicKey()

    'Create CryptoAPI Blob from Certificate
    Dim lPbkLen As Long
    Dim lOffset As Long
    Dim lStart As Long
    Dim sBlkLen As String
    Dim sRevKey As String
    Dim ASNStart As Long
    Dim ASNKEY As String
    '------------------Added by GioRock----------------------------- -> Updated 2012
    Dim RSA_KEY_V1 As String
    Dim RSA_KEY_V3 As String
    

    RSA_KEY_V1 = Chr(&H30) + Chr(&H81) + Chr(&H89) + Chr(&H2) + Chr(&H81) + Chr(&H81) + Chr(0)
    RSA_KEY_V3 = Chr(&H30) + Chr(&H82) + Chr(&H1) + Chr(&HA) + Chr(&H2) + Chr(&H82) + Chr(&H1) + Chr(&H1) + Chr(&H0)

    If InStr(ENCODED_CERT, RSA_KEY_V1) <> 0 Then
        lPbkLen = 1024
        RSA_KEY_LEN = 128
    Else
        lPbkLen = 2048
        RSA_KEY_LEN = 256
    End If
    '---------------------------------------------------------------

    lOffset = CLng(lPbkLen \ 8)
    lStart = 5 + (lOffset \ 128) * 2
    ASNStart = InStr(1, ENCODED_CERT, IIf(RSA_KEY_LEN = 128, RSA_KEY_V1, RSA_KEY_V3)) + lStart
    ASNKEY = Mid(ENCODED_CERT, ASNStart, RSA_KEY_LEN)

    sRevKey = ReverseString(ASNKEY)

    sBlkLen = CStr(Hex(lPbkLen \ 256))
    If Len(sBlkLen) = 1 Then sBlkLen = "0" & sBlkLen

    PUBLIC_KEY = (HexToBin( _
            "06020000" & _
            "00A40000" & _
            "52534131" & _
            "00" & sBlkLen & "0000" & _
            "01000100") & sRevKey)

'    blobheader.bType    = PUBLICKEYBLOB      // 0x06
'    blobheader.bVersion = CUR_BLOB_VERSION   // 0x02
'    blobheader.Reserved = 0                  // 0x0000
'    blobheader.aiKeyAlg = CALG_RSA_KEYX      // 0x0000a400
'    rsapubkey.magic     = 0x31415352         // "RSA1"
'    rsapubkey.bitlen    = 512                // 0x00000200
'    rsapubkey.pubexp    = 65537              // 0x00010001

End Sub



Private Function VerifyMAC(ByVal DecryptedRecord As String) As Boolean

    'Verify the Message Authentication Code
    Dim PrependedMAC As String
    Dim RecordData As String
    Dim CalculatedMAC As String

    PrependedMAC = Mid(DecryptedRecord, 1, 16)
    RecordData = Mid(DecryptedRecord, 17)

    CalculatedMAC = MD5_Hash(CLIENT_READ_KEY & RecordData & RecvSequence)

    Call IncrementRecv

    If CalculatedMAC = PrependedMAC Then
        VerifyMAC = True
    Else
        VerifyMAC = False
    End If

End Function


Private Function RecvSequence() As String
Dim i As Integer

    'Convert Receive Counter to a String
    Dim TempString As String
    Dim TempSequence As Double
    Dim TempByte As Double

    TempSequence = RECV_SEQUENCE_NUMBER

    For i = 1 To 4
        TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
        TempSequence = Int(TempSequence / 256)
        TempString = Chr(TempByte) & TempString
    Next

    RecvSequence = TempString

End Function

Private Sub SendClientHello(ByRef Socket As Winsock)

    'Send Client Hello
    Layer = 0

    Call GenerateRandomBytes(16, CHALLENGE_DATA)

    SEND_SEQUENCE_NUMBER = 0
    RECV_SEQUENCE_NUMBER = 0

    CLIENT_HELLO = Chr(1) & _
                    Chr(0) & Chr(2) & _
                    Chr(0) & Chr(3) & _
                    Chr(0) & Chr(0) & _
                    Chr(0) & Chr(Len(CHALLENGE_DATA)) & _
                    Chr(1) & Chr(0) & Chr(128) & _
                    CHALLENGE_DATA

    If Socket.State = 7 Then Socket.SendData AddRecordHeader(CLIENT_HELLO)

End Sub

Private Sub SendMasterKey(ByRef Socket As Winsock)

    'Send Master Key
    Layer = 1

    Call GenerateRandomBytes(32, MASTER_KEY)

    Call CertificateToPublicKey

    ' Modified for both SSL2.0 & SSL3.0 version of Handshake Message -> Updated 2012
    Socket.SendData AddRecordHeader(Chr(2) & _
                                    Chr(1) & Chr(0) & Chr(128) & _
                                    Chr(0) & Chr(0) & _
                                    Chr(RSA_KEY_LEN \ 256) & Chr(256 - RSA_KEY_LEN) & _
                                    Chr(0) & Chr(0) & _
                                    ExportKeyBlob(MASTER_KEY, CLIENT_READ_KEY, CLIENT_WRITE_KEY, CHALLENGE_DATA, CONNECTION_ID, PUBLIC_KEY))

End Sub

Private Sub SendClientFinish(ByRef Socket As Winsock)

    'Send ClientFinished Message
    Layer = 2
    Call SSLSend(Socket, Chr(3) & CONNECTION_ID)

End Sub


Private Sub IncrementRecv()

    'Increment Counter for Each Record Received
    RECV_SEQUENCE_NUMBER = RECV_SEQUENCE_NUMBER + 1
    If RECV_SEQUENCE_NUMBER = 4294967296# Then RECV_SEQUENCE_NUMBER = 0

End Sub

Private Function BytesToLen(ByVal TwoBytes As String) As Long

    'Convert Byte Pair to Packet Length
    Dim FirstByteVal As Long
    FirstByteVal = Asc(Left(TwoBytes, 1))
    If FirstByteVal >= 128 Then FirstByteVal = FirstByteVal - 128

    BytesToLen = 256 * FirstByteVal + Asc(Right(TwoBytes, 1))

End Function

Private Function HexToBin(ByVal HexString As String) As String

    'Convert a Hexadecimal String to characters
    Dim BinString As String
    Dim i As Integer

    For i = 1 To Len(HexString) Step 2
        BinString = BinString & Chr(Val("&H" & Mid(HexString, i, 2)))
    Next i

    HexToBin = BinString

End Function

Private Function ReverseString(ByVal TheString As String) As String

    'Reverse String
    Dim Reversed As String
    Dim i As Integer

    For i = Len(TheString) To 1 Step -1
        Reversed = Reversed & Mid(TheString, i, 1)
    Next i
    ReverseString = Reversed

End Function

Private Sub ProcessData(ByVal TheData As String, Socket As Winsock)
    Dim pblnStatus As Boolean

    Debug.Print TheData

    Select Case State
        Case Connect
            pblnStatus = CheckResponse(TheData, "220", Socket)
            Call SocketSend(Socket, "HELO" & gstrDomainName & vbCrLf)
            If pblnStatus = True Then
                State = State + 1
            Else
                Socket.Close
                State = 0
                gblnDone = True
            End If
        Case helo
            pblnStatus = CheckResponse(TheData, "250", Socket)
            If pblnStatus = True Then
                Call SocketSend(Socket, "MAIL FROM: " & gstrFromAddress & vbCrLf)
                State = State + 1
            Else
                Socket.Close
                State = 0
                gblnDone = True
                Exit Sub
            End If
        Case MailFrom
            pblnStatus = CheckResponse(TheData, "250", Socket)
            If pblnStatus = True Then
                If gintSent > UBound(gstrAllAddresses) And gblnBCCMode = False And gstrAllAddresses(0) <> "" Then
                    gintsentTo = 0
                    gintsentCC = 0
                    State = State + 1
                    Call ProcessData("250 ", Socket)
                Else
                    If gstrAllAddresses(0) = "" Then
                        gblnDone = True
                        Exit Sub
                    End If
                    If gblnBCCMode = True Then
                        pintSend = 0
                        gintSent = 0
                    End If
                    Call SocketSend(Socket, "RCPT TO:" & gstrAllAddresses(gintSent) & vbCrLf)
                End If
                If gblnBCCMode Then
                    State = State + 1
                    gintSent = 0
                Else
                    gintSent = gintSent + 1
                End If
            Else
                If Left(TheData, 3) = "530" Then
                    Call SocketSend(Socket, "AUTH LOGIN" & vbCrLf)
                ElseIf Left(TheData, 3) = "334" Then
                    Dim pstrIn As String
                    Dim pstrAnswer As String
                    pstrIn = Mid(TheData, 4)
                    pstrIn = Left(pstrIn, Len(pstrIn) - 2)
                    pstrIn = Base64String(Trim(pstrIn))
                    If InStr(LCase(pstrIn), "user") Then
                        pstrAnswer = gstrUserID
                        pstrAnswer = EncodeStringBase64(gstrUserID)
                    ElseIf InStr(LCase(pstrIn), "pass") Then
                        pstrAnswer = gstrUserPass
                        pstrAnswer = EncodeStringBase64(pstrAnswer)
                    End If
                    Call SocketSend(Socket, pstrAnswer & vbCrLf)
                    ElseIf Left(TheData, 3) = "501" Then
                        gblnDone = True
                    ElseIf Left(TheData, 3) = "235" Then
                        Call SocketSend(Socket, "MAIL FROM: " & gstrFromAddress & vbCrLf)
                        State = MailFrom
                    End If
                End If
        Case sendto
            Call SocketSend(Socket, "DATA" & vbCrLf)
            State = State + 1
        Case Data
            Call RTBX.LoadFile(gstrMessageFileName, rtfText)
            If gblnBCCMode = True Then
                Dim pintLine As Long
                Dim pintPos As Long
                Dim pintStringStart As Long
                Dim pintStringEnd As Long
                With RTBX
                    pintPos = .Find("CC: ")
                    pintLine = .GetLineFromChar(pintPos)
                    pintStringEnd = InStr(pintPos, .Text, vbCrLf)
                    .SelStart = pintPos
                    .SelLength = pintStringEnd - pintPos
                    .SelText = "CC: "
                    pintPos = .Find("TO: ")
                    pintLine = .GetLineFromChar(pintPos)
                    pintStringEnd = InStr(pintPos, .Text, vbCrLf)
                    .SelStart = pintPos
                    .SelLength = pintStringEnd - pintPos
                    .SelText = "TO: " & gstrAllAddresses(0)
                End With
            End If
            gstrData = Split(RTBX.Text, vbCrLf)
            If gstrBCC(0) <> "" Then
                ProgBar.Max = (UBound(gstrData) * (UBound(gstrBCC) + 1))
            Else
                ProgBar.Max = UBound(gstrData)
            End If
            For pintSend = 0 To UBound(gstrData)
                Call SocketSend(Socket, gstrData(pintSend) & vbCrLf)
                On Error Resume Next
                ProgBar.Value = gintTransfer
                gintTransfer = gintTransfer + 1
                On Error GoTo 0
                DoEvents
            Next pintSend
            State = State + 1
            Call SocketSend(Socket, vbCrLf & "." & vbCrLf)
        Case MessageData
            gblnDone = True
            State = State + 1

    End Select

End Sub

Private Function CheckResponse(pstrGot, PstrExpect, Socket As Winsock) As Boolean
    Dim pstrResp As String

    pstrResp = Mid(pstrGot, 1, 3)

    If StrComp(pstrResp, PstrExpect) = 0 Then
        CheckResponse = True
    Else
        CheckResponse = False
    End If

End Function



Private Function RC4_Decrypt(ByVal Ciphertext As String) As String

    'Decrypt with Client Read Key
    Dim lngLength As Long
    Dim lngReturnValue As Long

    lngLength = Len(Ciphertext)
    lngReturnValue = CryptDecrypt(hClientReadKey, 0, False, 0, Ciphertext, lngLength)

    RC4_Decrypt = Ciphertext

End Function



Private Function GenerateRandomBytes(ByVal Length As Long, ByRef TheString As String) As Boolean

    'Generate Random Bytes
    Dim i As Integer

    Randomize
    TheString = ""
    For i = 1 To Length
        TheString = TheString & Chr(Int(Rnd * 256))
    Next

    GenerateRandomBytes = CryptGenRandom(hCryptProv, Length, TheString)

End Function


Private Function ExportKeyBlob(ByRef StrMasterKey As String, ByRef StrReadKey As String, ByRef StrWriteKey As String, ByVal StrChallenge As String, ByVal StrConnectionID As String, ByVal StrPublicKey As String) As String

    'Create Keys and Return PKCS Block
    Dim lngReturnValue As Long
    Dim lngLength As Long
    Dim rgbBlob As String
    Dim hPublicKey As Long

    Call CreateKey(hMasterKey, StrMasterKey)
    StrMasterKey = MD5_Hash(StrMasterKey)

    Call CreateKey(hClientReadKey, StrMasterKey & "0" & StrChallenge & StrConnectionID)
    Call CreateKey(hClientWriteKey, StrMasterKey & "1" & StrChallenge & StrConnectionID)

    StrReadKey = MD5_Hash(StrMasterKey & "0" & StrChallenge & StrConnectionID)
    StrWriteKey = MD5_Hash(StrMasterKey & "1" & StrChallenge & StrConnectionID)

    lngReturnValue = CryptImportKey(hCryptProv, StrPublicKey, Len(StrPublicKey), 0, 0, hPublicKey)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not import a Public Key (CryptImportKey API)"

    lngReturnValue = CryptExportKey(hMasterKey, hPublicKey, SIMPLEBLOB, 0, vbNullString, lngLength)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not allocate a Pre-Master Key (CryptExportKey API)"
    rgbBlob = String(lngLength, 0)
    lngReturnValue = CryptExportKey(hMasterKey, hPublicKey, SIMPLEBLOB, 0, rgbBlob, lngLength)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not export a Pre-Master Key (CryptExportKey API)"

    If hPublicKey <> 0 Then CryptDestroyKey hPublicKey
    If hMasterKey <> 0 Then CryptDestroyKey hMasterKey


    ExportKeyBlob = StrReverse(Right(rgbBlob, RSA_KEY_LEN)) '  -> Updated 2012

End Function


Private Sub CreateKey(ByRef KeyName As Long, ByVal HashData As String)

    'Create a Session Key from a Hash
    Dim lngParams As Long
    Dim lngReturnValue As Long
    Dim lngHashLen As Long
    Dim hHash As Long

    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"

    lngReturnValue = CryptHashData(hHash, HashData, Len(HashData), 0)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"

    lngParams = GEN_KEY_BITS Or CRYPT_EXPORTABLE
    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, KeyName)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a session key (CryptDeriveKey API)"

    If hHash <> 0 Then CryptDestroyHash hHash

End Sub

Private Sub Class_Initialize2()

    'Initiate Secure Session
    Dim lngReturnValue As Long
    Dim TheAnswer As Long

    lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) 'try to make a new key container

    If lngReturnValue = 0 Then
        lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) 'try to get a handle to a key container that already exists, and if it fails...
        If lngReturnValue = 0 Then TheAnswer = MsgBox("GCN has detected that you do not have the required High Encryption Pack installed." & vbCrLf & "Would like to download this pack from Microsoft's website?", 16 + vbYesNo)
    End If

'    If TheAnswer = vbYes Then
'        Call Shell("START http://www.microsoft.com/windows/ie/downloads/recommended/128bit/default.asp", vbHide)
'        SocketSSL_Close
'    End If
'
'    If TheAnswer = vbNo Then
'        SocketSSL_Close
'    End If

   Dim i As Long

   'Setup the encodeing and decoding lookup arrays.
   'Essentially we speed up the routine by pre-shifting
   'the data so it only needs combined with And and Or.
   m_Index4(0) = 65 'Asc("A")
   m_Index4(1) = 66 'Asc("B")
   m_Index4(2) = 67 'Asc("C")
   m_Index4(3) = 68 'Asc("D")
   m_Index4(4) = 69 'Asc("E")
   m_Index4(5) = 70 'Asc("F")
   m_Index4(6) = 71 'Asc("G")
   m_Index4(7) = 72 'Asc("H")
   m_Index4(8) = 73 'Asc("I")
   m_Index4(9) = 74 'Asc("J")
   m_Index4(10) = 75 'Asc("K")
   m_Index4(11) = 76 'Asc("L")
   m_Index4(12) = 77 'Asc("M")
   m_Index4(13) = 78 'Asc("N")
   m_Index4(14) = 79 'Asc("O")
   m_Index4(15) = 80 'Asc("P")
   m_Index4(16) = 81 'Asc("Q")
   m_Index4(17) = 82 'Asc("R")
   m_Index4(18) = 83 'Asc("S")
   m_Index4(19) = 84 'Asc("T")
   m_Index4(20) = 85 'Asc("U")
   m_Index4(21) = 86 'Asc("V")
   m_Index4(22) = 87 'Asc("W")
   m_Index4(23) = 88 'Asc("X")
   m_Index4(24) = 89 'Asc("Y")
   m_Index4(25) = 90 'Asc("Z")
   m_Index4(26) = 97 'Asc("a")
   m_Index4(27) = 98 'Asc("b")
   m_Index4(28) = 99 'Asc("c")
   m_Index4(29) = 100 'Asc("d")
   m_Index4(30) = 101 'Asc("e")
   m_Index4(31) = 102 'Asc("f")
   m_Index4(32) = 103 'Asc("g")
   m_Index4(33) = 104 'Asc("h")
   m_Index4(34) = 105 'Asc("i")
   m_Index4(35) = 106 'Asc("j")
   m_Index4(36) = 107 'Asc("k")
   m_Index4(37) = 108 'Asc("l")
   m_Index4(38) = 109 'Asc("m")
   m_Index4(39) = 110 'Asc("n")
   m_Index4(40) = 111 'Asc("o")
   m_Index4(41) = 112 'Asc("p")
   m_Index4(42) = 113 'Asc("q")
   m_Index4(43) = 114 'Asc("r")
   m_Index4(44) = 115 'Asc("s")
   m_Index4(45) = 116 'Asc("t")
   m_Index4(46) = 117 'Asc("u")
   m_Index4(47) = 118 'Asc("v")
   m_Index4(48) = 119 'Asc("w")
   m_Index4(49) = 120 'Asc("x")
   m_Index4(50) = 121 'Asc("y")
   m_Index4(51) = 122 'Asc("z")
   m_Index4(52) = 48 'Asc("0")
   m_Index4(53) = 49 'Asc("1")
   m_Index4(54) = 50 'Asc("2")
   m_Index4(55) = 51 'Asc("3")
   m_Index4(56) = 52 'Asc("4")
   m_Index4(57) = 53 'Asc("5")
   m_Index4(58) = 54 'Asc("6")
   m_Index4(59) = 55 'Asc("7")
   m_Index4(60) = 56 'Asc("8")
   m_Index4(61) = 57 'Asc("9")
   m_Index4(62) = 43 'Asc("+")
   m_Index4(63) = 47 'Asc("/")

   'Calculate the other Arrays
   For i = 0 To 63
      m_Index1((i * 4) And &HFC) = m_Index4(i)
      m_Index2(((i And &HF) * 16) Or ((i And &H30) \ 16)) = m_Index4(i)
      m_Index3((i \ 4 And &HF) Or ((i And &H3) * 64)) = m_Index4(i)
   Next i


   m_ReverseIndex4(65) = 0 'Asc("A")
   m_ReverseIndex4(66) = 1 'Asc("B")
   m_ReverseIndex4(67) = 2 'Asc("C")
   m_ReverseIndex4(68) = 3 'Asc("D")
   m_ReverseIndex4(69) = 4 'Asc("E")
   m_ReverseIndex4(70) = 5 'Asc("F")
   m_ReverseIndex4(71) = 6 'Asc("G")
   m_ReverseIndex4(72) = 7 'Asc("H")
   m_ReverseIndex4(73) = 8 'Asc("I")
   m_ReverseIndex4(74) = 9 'Asc("J")
   m_ReverseIndex4(75) = 10 'Asc("K")
   m_ReverseIndex4(76) = 11 'Asc("L")
   m_ReverseIndex4(77) = 12 'Asc("M")
   m_ReverseIndex4(78) = 13 'Asc("N")
   m_ReverseIndex4(79) = 14 'Asc("O")
   m_ReverseIndex4(80) = 15 'Asc("P")
   m_ReverseIndex4(81) = 16 'Asc("Q")
   m_ReverseIndex4(82) = 17 'Asc("R")
   m_ReverseIndex4(83) = 18 'Asc("S")
   m_ReverseIndex4(84) = 19 'Asc("T")
   m_ReverseIndex4(85) = 20 'Asc("U")
   m_ReverseIndex4(86) = 21 'Asc("V")
   m_ReverseIndex4(87) = 22 'Asc("W")
   m_ReverseIndex4(88) = 23 'Asc("X")
   m_ReverseIndex4(89) = 24 'Asc("Y")
   m_ReverseIndex4(90) = 25 'Asc("Z")
   m_ReverseIndex4(97) = 26 'Asc("a")
   m_ReverseIndex4(98) = 27 'Asc("b")
   m_ReverseIndex4(99) = 28 'Asc("c")
   m_ReverseIndex4(100) = 29 'Asc("d")
   m_ReverseIndex4(101) = 30 'Asc("e")
   m_ReverseIndex4(102) = 31 'Asc("f")
   m_ReverseIndex4(103) = 32 'Asc("g")
   m_ReverseIndex4(104) = 33 'Asc("h")
   m_ReverseIndex4(105) = 34 'Asc("i")
   m_ReverseIndex4(106) = 35 'Asc("j")
   m_ReverseIndex4(107) = 36 'Asc("k")
   m_ReverseIndex4(108) = 37 'Asc("l")
   m_ReverseIndex4(109) = 38 'Asc("m")
   m_ReverseIndex4(110) = 39 'Asc("n")
   m_ReverseIndex4(111) = 40 'Asc("o")
   m_ReverseIndex4(112) = 41 'Asc("p")
   m_ReverseIndex4(113) = 42 'Asc("q")
   m_ReverseIndex4(114) = 43 'Asc("r")
   m_ReverseIndex4(115) = 44 'Asc("s")
   m_ReverseIndex4(116) = 45 'Asc("t")
   m_ReverseIndex4(117) = 46 'Asc("u")
   m_ReverseIndex4(118) = 47 'Asc("v")
   m_ReverseIndex4(119) = 48 'Asc("w")
   m_ReverseIndex4(120) = 49 'Asc("x")
   m_ReverseIndex4(121) = 50 'Asc("y")
   m_ReverseIndex4(122) = 51 'Asc("z")
   m_ReverseIndex4(48) = 52 'Asc("0")
   m_ReverseIndex4(49) = 53 'Asc("1")
   m_ReverseIndex4(50) = 54 'Asc("2")
   m_ReverseIndex4(51) = 55 'Asc("3")
   m_ReverseIndex4(52) = 56 'Asc("4")
   m_ReverseIndex4(53) = 57 'Asc("5")
   m_ReverseIndex4(54) = 58 'Asc("6")
   m_ReverseIndex4(55) = 59 'Asc("7")
   m_ReverseIndex4(56) = 60 'Asc("8")
   m_ReverseIndex4(57) = 61 'Asc("9")
   m_ReverseIndex4(43) = 62 'Asc("+")
   m_ReverseIndex4(47) = 63 'Asc("/")

   'Calculate the other arrays.
   For i = 0 To 255
      If m_ReverseIndex4(i) <> 0 Then
         m_ReverseIndex1(i) = m_ReverseIndex4(i) * 4

         m_ReverseIndex2(i, 0) = m_ReverseIndex4(i) \ 16
         m_ReverseIndex2(i, 1) = (m_ReverseIndex4(i) And &HF) * 16

         m_ReverseIndex3(i, 0) = m_ReverseIndex4(i) \ 4
         m_ReverseIndex3(i, 1) = (m_ReverseIndex4(i) And &H3) * 64
      End If
   Next i

End Sub

Private Sub Class_Terminate2()
    'Free up Memory
    If hClientWriteKey <> 0 Then CryptDestroyKey hClientWriteKey
    If hClientReadKey <> 0 Then CryptDestroyKey hClientReadKey
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0

    Set SocketSSL = Nothing
    Set RTBX = Nothing

End Sub
Private Sub SocketSSL_Connect()
Dim strHello As String

    State = Connect

' Added by GioRock
'--------------------------------------------------------------'
    If InStr(gstrUserID, "@") <> 0 Then
        strHello = Left$(gstrUserID, InStr(gstrUserID, "@") - 1)
    Else
        strHello = gstrUserID
    End If

    SocketSSL.SendData "EHLO " + strHello + vbCrLf
'--------------------------------------------------------------'

    pbConnected = True

End Sub

Public Function ConnectSSLSocket() As Integer
    Dim pintC As Integer

    With SocketSSL
        .Close
        .protocol = sckTCPProtocol
        .Connect gstrServerName, glngPort
    End With

    Do
        DoEvents
        If pbConnected = True Then Exit Do
    Loop

    Do
        DoEvents
        If gblnDone = True Then Exit Do
    Loop

    ReDim gstrCC(0)
    ReDim gstrToAddress(0)
    ReDim gstrAllAddresses(0)

    pintC = 0

    On Error Resume Next

    On Error GoTo 0

    Do Until pintC > UBound(gstrBCC)
        gblnBCCMode = True
        If gstrBCC(pintC) = "" Or gstrBCC(pintC) = "<>" Then
            Exit Do
        End If
        gstrAllAddresses(0) = gstrBCC(pintC)
        gblnDone = False
        State = helo
        Call ProcessData("250", SocketSSL)
        Do
            DoEvents
        Loop Until gblnDone = True
        pintC = pintC + 1
    Loop

    SocketSSL.Close

    gblnBCCMode = False

End Function

Private Sub SocketSSL_DataArrival(ByVal bytesTotal As Long)

    'Parse each SSL Record
    Dim TheData As String
    Static lCMDSequence As Long

' Added by Giorock
'---------------------------------------------------------'
    ' Starting a Secure Socket Layer connection
    If lCMDSequence < 3 Then
        Select Case lCMDSequence
            Case 0
                SocketSSL.GetData TheData, vbString
'                wsSSL.SendData "EHLO giorock" + vbCrLf
'                Send on Connect Event
            Case 1
                SocketSSL.GetData TheData, vbString
                SocketSSL.SendData "STARTTLS" + vbCrLf
            Case 2
                SocketSSL.GetData TheData, vbString
                SendClientHello SocketSSL
        End Select
        lCMDSequence = lCMDSequence + 1
        Debug.Print TheData
        Exit Sub
    End If
'---------------------------------------------------------'

    Do

        If SeekLen = 0 Then
            If bytesTotal >= 2 Then
                SocketSSL.GetData TheData, vbString, 2
                SeekLen = BytesToLen(TheData)
                bytesTotal = bytesTotal - 2
            Else
                Exit Sub
            End If
        End If

        If bytesTotal >= SeekLen Then
            SocketSSL.GetData TheData, vbString, SeekLen
            bytesTotal = bytesTotal - SeekLen
        Else
            Exit Sub
        End If

        Debug.Print "Layer -> " & Layer

        Select Case Layer
            Case 0:
                'Debug.Print BytesToLen(Mid(TheData, 6, 2))
                ENCODED_CERT = Mid(TheData, 12, BytesToLen(Mid(TheData, 6, 2)))
                CONNECTION_ID = Right(TheData, BytesToLen(Mid(TheData, 10, 2)))
                Call IncrementRecv
                Call SendMasterKey(SocketSSL)
            Case 1:
                TheData = RC4_Decrypt(TheData)
                If Right(TheData, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
                    If VerifyMAC(TheData) Then Call SendClientFinish(SocketSSL)
                Else
                    SocketSSL.Close
                End If
             Case 2:
                TheData = RC4_Decrypt(TheData)
                If VerifyMAC(TheData) = False Then
                    SocketSSL.Close
                Else
' Added by Giorock
'-----------------------------------------------------------'
                    State = MailFrom
                    SSLSend SocketSSL, "AUTH LOGIN" + vbCrLf
'-----------------------------------------------------------'
                End If
                Layer = 3
             Case 3
                TheData = RC4_Decrypt(TheData)
                If VerifyMAC(TheData) Then
                    Call ProcessData(Mid(TheData, 17), SocketSSL)
                End If
        End Select

        SeekLen = 0

    Loop Until bytesTotal = 0

End Sub

