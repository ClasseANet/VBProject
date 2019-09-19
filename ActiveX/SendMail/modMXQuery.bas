Attribute VB_Name = "modMXQuery"
'**********************************************************
'   The DNS & MXQuery code in this module was adapted from
'   MX.OCX code.
'**********************************************************
Option Explicit

' winsock
Private Const DNS_RECURSION As Byte = 1
Private Const AF_INET = 2
Private Const SOCKET_ERROR = -1
Private Const ERROR_BUFFER_OVERFLOW = 111
Private Const SOCK_DGRAM = 2
Private Const INADDR_NONE = &HFFFFFFFF
Private Const INADDR_ANY = &H0
' registry access
Private Const REG_SZ = 1&
Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY

' winsock
Private Type WSADATA
    wVersion                As Integer
    wHighVersion            As Integer
    szDescription(256)      As Byte
    szSystemStatus(128)     As Byte
    iMaxSockets             As Integer
    iMaxUdpDg               As Integer
    lpVendorInfo            As Long
End Type

Private Type DNS_HEADER
    qryID                   As Integer
    options                 As Byte
    response                As Byte
    qdcount                 As Integer
    ancount                 As Integer
    nscount                 As Integer
    arcount                 As Integer
End Type

Private Type IP_ADDRESS_STRING
    IpAddressStr(4 * 4 - 1) As Byte
End Type
 
Private Type IP_MASK_STRING
    IpMaskString(4 * 4 - 1) As Byte
End Type
 
Private Type IP_ADDR_STRING
    Next                    As Long
    IpAddress               As IP_ADDRESS_STRING
    IpMask                  As IP_MASK_STRING
    Context                 As Long
End Type

Private Type FIXED_INFO
    HostName(128 + 4 - 1)   As Byte
    DomainName(128 + 4 - 1) As Byte
    CurrentDnsServer        As Long
    DnsServerList           As IP_ADDR_STRING
    NodeType                As Long
    ScopeId(256 + 4 - 1)    As Byte
    EnableRouting           As Long
    EnableProxy             As Long
    EnableDns               As Long
End Type

Private Type SOCKADDR
    sin_family              As Integer
    sin_port                As Integer
    sin_addr                As Long
    sin_zero                As String * 8
End Type

Private Type HostEnt
    h_name                  As Long
    h_aliases               As Long
    h_addrtype              As Integer
    h_length                As Integer
    h_addr_list             As Long
End Type

' registry
Private Type FILETIME
    dwLowDateTime           As Long
    dwHighDateTime          As Long
End Type

' public type for passing DNS info
Public Type DNS_INFO
    Servers()               As String
    Count                   As Long
    LocalDomain             As String
    RootDomain              As String
End Type

' used below
Public Type MX_RECORD
    Server                  As String
    Pref                    As Integer
End Type

' public type for passing MX info
Public Type MX_INFO
    Best                    As String
    Domain                  As String
    List()                  As MX_RECORD
    Count                   As Long
End Type

Public DNS                  As DNS_INFO
Public MX                   As MX_INFO


' API prototypes

' winsock, 'wsock32.dll' used instead of 'ws2_32.dll' for wider compatibility
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
Private Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function recvfrom Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, from As SOCKADDR, fromlen As Long) As Long
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
Private Declare Function sendto Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, to_addr As SOCKADDR, ByVal tolen As Long) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

' Registry access
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

' misc
Private Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Public Sub GetDNSInfo()

    ' get the DNS servers and the local IP Domain name
    
    Dim sBuffer                 As String
    Dim sDNSBuff                As String
    Dim sDomainBuff             As String
    Dim sKey                    As String
    Dim lngFixedInfoNeeded      As Long
    Dim bytFixedInfoBuffer()    As Byte
    Dim udtFixedInfo            As FIXED_INFO
    Dim lngIpAddrStringPtr      As Long
    Dim udtIpAddrString         As IP_ADDR_STRING
    Dim strDnsIpAddress         As String
    Dim nRet                    As Long
    Dim sTmp()                  As String
    Dim i                       As Long
       
    ' get dns servers with the new GetNetworkParams call (only works on 98/ME/2000)
    ' if GetNetworkParams is not supported then try reading from the registry
    If Exported("iphlpapi.dll", "GetNetworkParams") Then
        nRet = GetNetworkParams(ByVal vbNullString, lngFixedInfoNeeded)
        If nRet = ERROR_BUFFER_OVERFLOW Then
            ReDim bytFixedInfoBuffer(lngFixedInfoNeeded)
            nRet = GetNetworkParams(bytFixedInfoBuffer(0), lngFixedInfoNeeded)
            CopyMemory udtFixedInfo, bytFixedInfoBuffer(0), Len(udtFixedInfo)
            With udtFixedInfo
                ' get the DNS servers
                lngIpAddrStringPtr = VarPtr(.DnsServerList)
                Do While lngIpAddrStringPtr
                    CopyMemory udtIpAddrString, ByVal lngIpAddrStringPtr, Len(udtIpAddrString)
                    With udtIpAddrString
                        strDnsIpAddress = StripTerminator(StrConv(.IpAddress.IpAddressStr, vbUnicode))
                        sDNSBuff = sDNSBuff & strDnsIpAddress & ","
                        lngIpAddrStringPtr = .Next
                    End With
                Loop
                ' get the ip domain name
                sDomainBuff = StripTerminator(StrConv(.DomainName, vbUnicode))
            End With
        End If
    End If
    
    ' if GetNetworkParams didn't get the data we need,
    ' try known locations in the registry for DNS & domain info
    If Len(sDNSBuff) = 0 Or Len(sDomainBuff) = 0 Then

        ' DNS servers configured through Network control panel applet (95/98/ME)
        sKey = "System\CurrentControlSet\Services\VxD\MSTCP"
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "NameServer", "")
        If Len(sBuffer) Then sDNSBuff = sBuffer & ","
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "Domain", "")
        If Len(sBuffer) Then sDomainBuff = sBuffer

        ' DNS servers configured through Network control panel applet (NT/2000)
        sKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "NameServer", "")
        If Len(sBuffer) Then sDNSBuff = sBuffer & ","
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "Domain", "")
        If Len(sBuffer) Then sDomainBuff = sBuffer

        ' DNS servers configured DHCP (NT/2000/XP)
        sKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "DhcpNameServer", "")
        If Len(sBuffer) Then sDNSBuff = sBuffer & ","
        sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey, "DHCPDomain", "")
        If Len(sBuffer) Then sDomainBuff = sBuffer

        ' DNS servers configured through Network control panel applet (XP)
        sKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces"
        sTmp = EnumRegKey(HKEY_LOCAL_MACHINE, sKey)
        For i = 0 To UBound(sTmp)
            sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey & "\" & sTmp(i), "NameServer", "")
            If Len(sBuffer) Then sDNSBuff = sBuffer & ","
            sBuffer = GetRegStr(HKEY_LOCAL_MACHINE, sKey & "\" & sTmp(i), "Domain", "")
            If Len(sBuffer) Then sDomainBuff = sBuffer
        Next
        
        ' DNS servers configured DHCP (95/98/ME)
        ' *** haven't found one ***
    
    End If

    ' get rid of any space delimiters (2000)
    sDNSBuff = Replace(sDNSBuff, " ", ",")

    ' trim any trailing commas
    If Right(sDNSBuff, 1) = "," Then sDNSBuff = Left(sDNSBuff, Len(sDNSBuff) - 1)

    ' load our type struc
    DNS.Servers = Split(sDNSBuff, ",")
    DNS.Count = UBound(DNS.Servers) + 1
    DNS.LocalDomain = sDomainBuff

    ' cheap trick
    If sDomainBuff = "" And DNS.Count > 0 Then
        sDomainBuff = GetRemoteHostName(DNS.Servers(0))
        nRet = InStr(sDomainBuff, ".")
        If nRet Then
            DNS.LocalDomain = Mid$(sDomainBuff, nRet + 1)
        End If
    End If

    sTmp = Split(sDomainBuff, ".")
    nRet = UBound(sTmp)
    If nRet > 0 Then
        DNS.RootDomain = sTmp(nRet - 1) & "." & sTmp(nRet)
    Else
        DNS.RootDomain = sDomainBuff
    End If

End Sub

Public Function MX_Query(ByVal ms_Domain As String) As String
    
    ' Performs the actual IP work to contact the DNS server,
    ' calls the other functions to parse and return the
    ' best server to send email through
    
    Dim StartupData     As WSADATA
    Dim SocketBuffer    As SOCKADDR
    Dim IpAddr          As Long
    Dim iRC             As Integer
    Dim dnsHead         As DNS_HEADER
    Dim iSock           As Integer
    Dim dnsQuery()      As Byte
    Dim sQName          As String
    Dim dnsQueryNdx     As Integer
    Dim iTemp           As Integer
    Dim iNdx            As Integer
    Dim dnsReply(2048)  As Byte
    Dim iAnCount        As Integer
    Dim dwFlags         As Long


    MX.Count = 0
    MX.Best = vbNullString
    ReDim MX.List(0)

    ' if DNSInfo hasn't been called, call it now
    If DNS.Count = 0 Then GetDNSInfo
    
    ' check to see that we found a dns server
    If DNS.Count = 0 Then
        ' problem
        Err.Raise 20000, "MXQuery", "No DNS entries found, MX Query cannot contine."
        Exit Function
    End If
   
    ' if null was passed in then use the local domain name
    If Len(ms_Domain) = 0 Then ms_Domain = DNS.LocalDomain
    
    ' validate domain name
    If Len(ms_Domain) < 5 Then
        Err.Raise 20000, "MXQuery", "No Valid Domain Specified"
        Exit Function
    End If
   
    MX.Domain = ms_Domain
   
    ' Initialize the Winsock, request v1.1
    If WSAStartup(&H101, StartupData) <> ERROR_SUCCESS Then
        iRC = WSACleanup
        Exit Function
    End If
    
    ' Create a socket
    iSock = socket(AF_INET, SOCK_DGRAM, 0)
    If iSock = SOCKET_ERROR Then Exit Function

    ' convert the IP address string to a network ordered long
    IpAddr = GetHostByNameAlias(DNS.Servers(0))
    If IpAddr = -1 Then Exit Function
    
    ' Setup the connnection parameters
    SocketBuffer.sin_family = AF_INET
    SocketBuffer.sin_port = htons(53)
    SocketBuffer.sin_addr = IpAddr
    SocketBuffer.sin_zero = String$(8, 0)
    
    ' Set the DNS parameters
    dnsHead.qryID = htons(&H11DF)
    dnsHead.options = DNS_RECURSION
    dnsHead.qdcount = htons(1)
    dnsHead.ancount = 0
    dnsHead.nscount = 0
    dnsHead.arcount = 0
    
    dnsQueryNdx = 0
    
    ReDim dnsQuery(4000)
    
    ' Setup the dns structure to send the query in
    ' First goes the DNS header information
    CopyMemory dnsQuery(dnsQueryNdx), dnsHead, 12
    dnsQueryNdx = dnsQueryNdx + 12
    
    ' Then the domain name (as a QNAME)
    sQName = MakeQName(MX.Domain)
    iNdx = 0
    While (iNdx < Len(sQName))
        dnsQuery(dnsQueryNdx + iNdx) = Asc(Mid(sQName, iNdx + 1, 1))
        iNdx = iNdx + 1
    Wend

    dnsQueryNdx = dnsQueryNdx + Len(sQName)
    
    ' Null terminate the string
    dnsQuery(dnsQueryNdx) = &H0
    dnsQueryNdx = dnsQueryNdx + 1
    
    ' The type of query (15 means MX query)
    iTemp = htons(15)
    CopyMemory dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ' The class of query (1 means INET)
    iTemp = htons(1)
    CopyMemory dnsQuery(dnsQueryNdx), iTemp, Len(iTemp)
    dnsQueryNdx = dnsQueryNdx + Len(iTemp)
    
    ReDim Preserve dnsQuery(dnsQueryNdx - 1)
    ' Send the query to the DNS server
    iRC = sendto(iSock, dnsQuery(0), dnsQueryNdx + 1, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Or (iRC = 0) Then
        Err.Raise 20000, "MXQuery", "Problem sending MX query"
        iRC = WSACleanup
        Exit Function
    End If

    ' Wait for answer from the DNS server
    iRC = recvfrom(iSock, dnsReply(0), 2048, 0, SocketBuffer, Len(SocketBuffer))
    If (iRC = SOCKET_ERROR) Or (iRC = 0) Then
        Err.Raise 20000, "MXQuery", "Problem receiving MX query"
        iRC = WSACleanup
        Exit Function
    End If

    ' Get the number of answers
    CopyMemory iAnCount, dnsReply(6), 2
    iAnCount = ntohs(iAnCount)
    
    iRC = WSACleanup
    
    If iAnCount Then
        ' Parse the answer buffer
        MX_Query = GetMXName(dnsReply(), 12, iAnCount)
        
    Else
        ' if we didn't find anything and we are part of
        ' a sub domain, go up one level and try again
        ' the last pass is at the root domain level
        If InStr(MX.Domain, DNS.RootDomain) > 1 Then
            MX.Domain = Mid$(MX.Domain, InStr(MX.Domain, ".") + 1)
            MX_Query = MX_Query(MX.Domain)
        End If
    End If
    
End Function

Private Sub ParseName(dnsReply() As Byte, iNdx As Integer, sName As String)
    
' Parse the server name out of the MX record, returns it in variable sName.
' iNdx is also modified to point to the end of the parsed structure.
    
    Dim iCompress       As Integer      ' Compression index (index to original buffer)
    Dim iChCount        As Integer      ' Character count (number of chars to read from buffer)
        
    ' While we dont encounter a null char (end-of-string specifier)
    While (dnsReply(iNdx) <> 0)
        ' Read the next character in the stream (length specifier)
        iChCount = dnsReply(iNdx)
        ' If our length specifier is 192 (0xc0) we have a compressed string
        If (iChCount = 192) Then
            ' Read the location of the rest of the string (offset into buffer)
            iCompress = dnsReply(iNdx + 1)
            ' Call ourself again, this time with the offset of the compressed string
            ParseName dnsReply(), iCompress, sName
            ' Step over the compression indicator and compression index
            iNdx = iNdx + 2
            ' After a compressed string, we are done
            Exit Sub
        End If
        
        ' Move to next char
        iNdx = iNdx + 1
        ' While we should still be reading chars
        While (iChCount)
            ' add the char to our string
            sName = sName + Chr(dnsReply(iNdx))
            iChCount = iChCount - 1
            iNdx = iNdx + 1
        Wend
        ' If the next char isn't null then the string continues, so add the dot
        If (dnsReply(iNdx) <> 0) Then sName = sName + "."
    Wend
    
End Sub

Private Function GetMXName(dnsReply() As Byte, iNdx As Integer, iAnCount As Integer) As String
    
' Parses the buffer returned by the DNS server, returns the best
' MX server (lowest preference number), iNdx is modified to point
' to the current buffer position (should be the end of the buffer
' by the end, unless a record other than MX is found)
    
    Dim iChCount        As Integer     ' Character counter
    Dim sTemp           As String      ' Holds the original query string
    Dim iBestPref       As Integer     ' Holds the "best" preference number (lowest)
    Dim iMXCount        As Integer
    
    
    MX.Count = 0
    MX.Best = vbNullString
    ReDim MX.List(0)

    iMXCount = 0
    iBestPref = -1
    
    ParseName dnsReply(), iNdx, sTemp
    
    ' Step over null
    iNdx = iNdx + 2
    
    ' Step over 6 bytes, not sure what the 6 bytes are, but
    ' all other documentation shows steping over these 6 bytes
    iNdx = iNdx + 6
    
    While (iAnCount)
        ' Check to make sure we received an MX record
        If (dnsReply(iNdx) = 15) Then
            Dim sName As String
            Dim iPref As Integer
            
            sName = ""
            
            ' Step over the last half of the integer that specifies the record type (1 byte)
            ' Step over the RR Type, RR Class, TTL (3 integers - 6 bytes)
            iNdx = iNdx + 1 + 6
            
            ' Step over the MX data length specifier (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            CopyMemory iPref, dnsReply(iNdx), 2
            iPref = ntohs(iPref)
            ' Step over the MX preference value (1 integer - 2 bytes)
            iNdx = iNdx + 2
            
            ' Have to step through the byte-stream, looking for 0xc0 or 192 (compression char)
            ParseName dnsReply(), iNdx, sName
            
            If Trim(sName) <> "" Then
                iMXCount = iMXCount + 1
                ReDim Preserve MX.List(iMXCount - 1)
                MX.List(iMXCount - 1).Server = sName
                MX.List(iMXCount - 1).Pref = iPref
                MX.Count = iMXCount
                If (iBestPref = -1 Or iPref < iBestPref) Then
                    iBestPref = iPref
                    MX.Best = sName
                End If
            End If
            ' Step over 3 useless bytes
            iNdx = iNdx + 3
        Else
            GetMXName = MX.Best
            SortMX MX.List
            Exit Function
        End If
        iAnCount = iAnCount - 1
    Wend
    
    SortMX MX.List
        
    GetMXName = MX.Best

End Function

Private Function MakeQName(sDomain As String) As String
    
' Takes sDomain and converts it to the QNAME-type string.
' QNAME is how a DNS server expects the string.
'
' Example:  Pass -        mail.com
'           Returns -     &H4mail&H3com
'                          ^      ^
'                          |______|____ These two are character counters, they count
'                                       the number of characters appearing after them
    
    Dim iQCount         As Integer      ' Character count (between dots)
    Dim iNdx            As Integer      ' Index into sDomain string
    Dim iCount          As Integer      ' Total chars in sDomain string
    Dim sQName          As String       ' QNAME string
    Dim sDotName        As String       ' Temp string for chars between dots
    Dim sChar           As String       ' Single char from sDomain string
    
    iNdx = 1
    iQCount = 0
    iCount = Len(sDomain)
    
    ' While we haven't hit end-of-string
    While (iNdx <= iCount)
        ' Read a single char from our domain
        sChar = Mid(sDomain, iNdx, 1)
        ' If the char is a dot, then put our character count and the part of the string
        If (sChar = ".") Then
            sQName = sQName & Chr(iQCount) & sDotName
            iQCount = 0
            sDotName = ""
        Else
            sDotName = sDotName + sChar
            iQCount = iQCount + 1
        End If
        iNdx = iNdx + 1
    Wend
    
    sQName = sQName & Chr(iQCount) & sDotName
    
    MakeQName = sQName
    
End Function

Private Function GetHostByNameAlias(ByVal sHostName As String) As Long
    
    'Return IP address as a long, in network byte order
    
    Dim phe             As Long
    Dim heDestHost      As HostEnt
    Dim addrList        As Long
    Dim retIP           As Long
    
    retIP = inet_addr(sHostName)
    
    If retIP = INADDR_NONE Then
        phe = gethostbyname(sHostName)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, LenB(heDestHost)
            CopyMemory addrList, ByVal heDestHost.h_addr_list, 4
            CopyMemory retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    
    GetHostByNameAlias = retIP
    
End Function

Private Function StripTerminator(ByVal strString As String) As String
    
    ' strip off trailing NULL's from API calls
    
    Dim intZeroPos      As Integer

    intZeroPos = InStr(strString, vbNullChar)
    
    If intZeroPos > 1 Then
        StripTerminator = Trim$(Left$(strString, intZeroPos - 1))
    ElseIf intZeroPos = 1 Then
        StripTerminator = vbNullString
    Else
        StripTerminator = strString
    End If
    
End Function

Private Function GetRegStr(hKeyRoot As Long, ByVal sKeyName As String, ByVal sValueName As String, Optional ByVal Default As String = "") As String
   
   Dim lRet             As Long
   Dim hKey             As Long
   Dim lType            As Long
   Dim lBytes           As Long
   Dim sBuff            As String
   
   ' in case there's a permissions violation
   On Local Error GoTo Err_Reg

   ' Assume failure and set return to Default
   GetRegStr = Default

   ' Open the key
   lRet = RegOpenKeyEx(hKeyRoot, sKeyName, 0&, KEY_READ, hKey)
   If lRet = ERROR_SUCCESS Then
      
      ' Determine the buffer size
      lRet = RegQueryValueEx(hKey, sValueName, 0&, lType, ByVal sBuff, lBytes)
      If lRet = ERROR_SUCCESS Then
         ' size the buffer & call again
         If lBytes > 0 Then
            sBuff = Space(lBytes)
            lRet = RegQueryValueEx(hKey, sValueName, 0&, lType, ByVal sBuff, Len(sBuff))
            If lRet = ERROR_SUCCESS Then
               ' Trim NULL and return
               GetRegStr = Left(sBuff, lBytes - 1)
            End If
         End If
      End If
      Call RegCloseKey(hKey)
   End If
   
   Exit Function
   
Err_Reg:

  If hKey Then Call RegCloseKey(hKey)
   
End Function

Private Function EnumRegKey(hKeyRoot As Long, sKeyName As String) As String()
    

    Dim lRet            As Long
    Dim ft              As FILETIME
    Dim hKey            As Long
    Dim CurIdx          As Long
    Dim KeyName         As String
    Dim ClassName       As String
    Dim KeyLen          As Long
    Dim ClassLen        As Long
    Dim RESERVED        As Long
    Dim sEnum()         As String
    
    On Local Error GoTo Err_Enum
    
    ' initialize array
    EnumRegKey = Split("", "")
    
    ' Open the key
    lRet = RegOpenKeyEx(hKeyRoot, sKeyName, 0&, KEY_READ, hKey)
    If lRet <> ERROR_SUCCESS Then Exit Function
    
    ' the key opened so get all the sub keys
    Do
        ' get each sub key until lRet = error
        KeyLen = 2000
        ClassLen = 2000
        KeyName = String$(KeyLen, 0)
        ClassName = String$(ClassLen, 0)
        lRet = RegEnumKeyEx(hKey, CurIdx, KeyName, KeyLen, RESERVED, ClassName, ClassLen, ft)

        If lRet = ERROR_SUCCESS Then
            ReDim Preserve sEnum(CurIdx)
            sEnum(CurIdx) = Left$(KeyName, KeyLen)
        End If
    
        CurIdx = CurIdx + 1
        
    Loop While lRet = ERROR_SUCCESS
      
Err_Enum:

    EnumRegKey = sEnum
    If hKey Then Call RegCloseKey(hKey)

End Function

Private Function Exported(ByVal ModuleName As String, ByVal ProcName As String) As Boolean
   
    ' see if the api supports a call
    
    Dim hModule         As Long
    Dim lpProc          As Long
    Dim FreeLib         As Boolean
   
    ' check to see if the module is already
    ' mapped into this process.
    hModule = GetModuleHandle(ModuleName)
    If hModule = 0 Then
        ' not mapped, load the module into this process.
        hModule = LoadLibrary(ModuleName)
        FreeLib = True
    End If
   
    ' check the procedure address to verify it's exported.
    If hModule Then
        lpProc = GetProcAddress(hModule, ProcName)
        Exported = (lpProc <> 0)
    End If
   
    ' unload library if we loaded it here.
    If FreeLib Then Call FreeLibrary(hModule)
    
End Function

Private Sub SortMX(arr() As MX_RECORD, Optional ByVal bSortDesc As Boolean = False)

    ' simple bubble sort

    Dim ValMX           As MX_RECORD
    Dim index           As Long
    Dim firstItem       As Long
    Dim indexLimit      As Long
    Dim lastSwap        As Long

    firstItem = LBound(arr)
    lastSwap = UBound(arr)
    
    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For index = firstItem To indexLimit
            ValMX.Pref = arr(index).Pref
            ValMX.Server = arr(index).Server
            If (ValMX.Pref > arr(index + 1).Pref) Xor bSortDesc Then
                ' if the items are not in order, swap them
                arr(index).Pref = arr(index + 1).Pref
                arr(index).Server = arr(index + 1).Server
                arr(index + 1).Pref = ValMX.Pref
                arr(index + 1).Server = ValMX.Server
                lastSwap = index
            End If
        Next
    Loop While lastSwap

End Sub

Public Function GetRemoteHostName(ByVal strIpAddress As String) As String

    Dim udtHostEnt      As HostEnt  ' HOSTENT structure
    Dim lngPtrHostEnt   As Long     ' pointer to HOSTENT
    Dim lngInetAddr     As Long     ' address as a Long value
    Dim strHostName     As String   ' string buffer for host name

    ' initialize the buffer
    strHostName = String(256, 0)

    ' Convert IP address to Long
    lngInetAddr = inet_addr(strIpAddress)
    If lngInetAddr = INADDR_NONE Then Exit Function
        
    ' Get the HostEnt structure pointer
    lngPtrHostEnt = gethostbyaddr(lngInetAddr, 4, AF_INET)
    If lngPtrHostEnt = 0 Then Exit Function
            
    ' Copy data into the HostEnt structure
    CopyMemory udtHostEnt, ByVal lngPtrHostEnt, LenB(udtHostEnt)
    CopyMemory ByVal strHostName, ByVal udtHostEnt.h_name, Len(strHostName)

    GetRemoteHostName = StripTerminator(strHostName)

End Function
