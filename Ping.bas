Attribute VB_Name = "Ping"
Option Explicit

'   14Oct2006 Module created by copying Microsoft code sample, stripping out un-necessary code,
'   and making minor alterations such as adding a timeout parameter to Ping functions
'   (Ping function was supplied but has been stripped back, IsPingSuccessful function was written)
'   Stripped out a number of status functions because only really need to know whether ping
'   succeeded or failed. Not concerned with why

'Icmp constants converted from
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Private Const ICMP_SUCCESS As Long = 0
Private Const WS_VERSION_REQD As Long = &H101

'Clean up sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

'Open the socket connection.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, ByRef lpWSADATA As WSADATA) As Long

'Create a handle on which Internet Control Message Protocol (ICMP) requests can be issued.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpcreatefile.asp
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

'Convert a string that contains an (Ipv4) Internet Protocol dotted address into a correct address.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winsock/wsapiref_4esy.asp
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long

'Close an Internet Control Message Protocol (ICMP) handle that IcmpCreateFile opens.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpclosehandle.asp
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

'Information about the Windows Sockets implementation
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets As Long
   iMaxUDPDG As Long
   lpVendorInfo As Long
End Type

'Send an Internet Control Message Protocol (ICMP) echo request, and then return one or more replies.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIcmpSendEcho.asp
'   Requirements
'   Client Requires Windows Vista or Windows XP.
'   Server Requires Windows Server "Longhorn" or Windows Server 2003.
'   Header Declared in Icmpapi.h.
'   Library Use Iphlpapi.lib.
'   DLL Requires Iphlpapi.dll on Windows Server "Longhorn", Windows Vista, Windows Server 2003, and Windows XP.
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ByRef ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
 
'This structure describes the options that will be included in the header of an IP packet.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIP_OPTION_INFORMATION.asp
Private Type IP_OPTION_INFORMATION
   Ttl             As Byte
   Tos             As Byte
   Flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

'This structure describes the data that is returned in response to an echo request.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmp_echo_reply.asp
Private Type ICMP_ECHO_REPLY
   address         As Long
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Long
   Reserved        As Integer
   ptrData         As Long
   Options         As IP_OPTION_INFORMATION
   Data            As String * 250
End Type

'-- Ping a string representation of an IP address.
' -- Return a reply.
' -- Return long code.
Public Function Ping(ByVal sIpAddress As String, ByRef Reply As ICMP_ECHO_REPLY, Optional ByVal pTimeout As Long = 1000) As Long
'ublic Function ping(sIpAddress As String, Reply As ICMP_ECHO_REPLY) As Long

Dim hIcmp As Long
Dim lAddress As Long
Dim StringToSend As String
'''Dim lTimeOut As Long
''''ICMP (ping) timeout
'''lTimeOut = 1000 'ms

'Short string of data to send
StringToSend = "hello"

'Convert string address to a long representation.
lAddress = inet_addr(sIpAddress)

If (lAddress <> -1) And (lAddress <> 0) Then
        
    'Create the handle for ICMP requests.
    hIcmp = IcmpCreateFile()
    
    If hIcmp Then
        'Ping the destination IP address.
'---------------------------------------------------------------------------------------------------
'   IcmpSendEcho Requirements
'   Client Requires Windows Vista or Windows XP.
'   Server Requires Windows Server "Longhorn" or Windows Server 2003.
'   Header Declared in Icmpapi.h.
'   Library Use Iphlpapi.lib.
'   DLL Requires Iphlpapi.dll on Windows Server "Longhorn", Windows Vista, Windows Server 2003, and Windows XP.
        IcmpSendEcho hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), pTimeout
'---------------------------------------------------------------------------------------------------

        'Reply status
        Ping = Reply.Status
        
        'Close the Icmp handle.
        IcmpCloseHandle hIcmp
    Else
    '   Failure opening icmp handle.
        Ping = -1
    End If
Else
    Ping = -1
End If

End Function

'Clean up the sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Sub SocketsCleanup()
   
   WSACleanup
    
End Sub

'Get the sockets ready.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA

   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = ICMP_SUCCESS

End Function

Public Function IsPingSuccessful(ByVal pIpAddress As String, Optional ByVal pTimeout As Long = 1000) As Boolean
Dim Reply As ICMP_ECHO_REPLY
Dim bSuccess As Boolean
   
'   Get the sockets ready.
    If SocketsInitialize() Then
    '   Ping the IP that is passing the address and get a reply.
        bSuccess = Ping(pIpAddress, Reply, pTimeout) = ICMP_SUCCESS
    '   Clean up the sockets.
        SocketsCleanup
'   Else
'   '   Winsock error failure, initializing the sockets.
    End If
   
   IsPingSuccessful = bSuccess

End Function
