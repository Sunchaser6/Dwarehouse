Attribute VB_Name = "WNetFunctions"
Option Explicit

' This module collects the WNet network functions together (NOT communication functions eg RAS etc)
' The Windows network functions enable your application to explore and manage network connections
' directly, or to give direct control of the network connections to your users. To call these
' functions, you must link to the multiple provider router library (MPR.LIB).

' At times, it is necessary to map a drive letter to a network share.
' There are several API functions that can be used to accomplish this,
' such as WNetAddConnection, WNetAddConnection2, WNetAddConnection3,
' and WNetUseConnection. The primary difference is that with the
' WNetUseConnection function, you do not need to specify the drive letter
' to be used, while it is required with the other API functions.

'   Function, type and constant declarations are provided as Public declarations but
'   I made them Private because I have wrapped all of them in procedures in this module.
'   Variable names in type declarations changed to use VB data type prefixes (eg lng instead of dw, etc)
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
'   Downloaded 'Declare Function WNetUseConnection' from internet as it wasn't available in my copy of WIN32API.TXT
'Private Declare Function WNetUseConnection Lib "mpr.dll" Alias "WNetUseConnectionA" (ByVal hwndOwner As Long, ByRef lpNetResource As NETRESOURCE, ByVal lpUsername As String, ByVal lpPassword As String, ByVal dwFlags As Long, ByVal lpAccessName As Any, ByRef lpBufferSize As Long, ByRef lpResult As Long) As Long

Private Type NETRESOURCE
    lngScope As Long         ' To specify scope during enumeration
    lngType As Long          ' Defines the type of resource
    lngDisplayType As Long   ' How resources will be displayed
    lngUsage As Long         ' Specifies the resource usage
    strLocalName As String   ' Local device for the connection
    strRemoteName As String  ' Indicates the network resource
    strComment As String     ' For a provider-supplied comment
    strProvider As String    ' Name of provider who owns the resource
End Type

'   Constants for NETRESOURCE
 Private Const RESOURCETYPE_DISK As Long = &H1
'Private Const RESOURCETYPE_PRINT As Long = &H2
'Private Const RESOURCETYPE_ANY As Long = &H0
'Private Const RESOURCE_CONNECTED As Long = &H1
'Private Const RESOURCE_REMEMBERED As Long = &H3
 Private Const RESOURCE_GLOBALNET As Long = &H2
'Private Const RESOURCEDISPLAYTYPE_DOMAIN As Long = &H1
'Private Const RESOURCEDISPLAYTYPE_GENERIC As Long = &H0
'Private Const RESOURCEDISPLAYTYPE_SERVER As Long = &H2
'Private Const RESOURCEDISPLAYTYPE_SHARE As Long = &H3
'Private Const RESOURCEUSAGE_CONNECTABLE As Long = &H1
'Private Const RESOURCEUSAGE_CONTAINER As Long = &H2

'   Miscellaneous Constants
'Private Const CONNECT_UPDATE_PROFILE As Long = &H1
Private Const NO_ERROR As Long = 0

'   Network Error Constants
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_ALREADY_ASSIGNED As Long = 85
Private Const ERROR_ALREADY_CONNECTED As Long = 52          ' VB4 Unleashed from internet - not in Win32API.txt
Private Const ERROR_BAD_DEV_TYPE As Long = 66
Private Const ERROR_BAD_DEVICE As Long = 1200
Private Const ERROR_BAD_NET_NAME As Long = 67
Private Const ERROR_BAD_NETPATH As Long = 53
Private Const ERROR_BAD_PROFILE As Long = 1206
Private Const ERROR_BAD_PROVIDER As Long = 1204
Private Const ERROR_BAD_USERNAME As Long = 2202
Private Const ERROR_BUSY As Long = 170
Private Const ERROR_CANCEL_VIOLATION As Long = 173
Private Const ERROR_CANCELLED As Long = 1223                ' VB4 Unleashed from internet - not in Win32API.txt
Private Const ERROR_CANNOT_OPEN_PROFILE As Long = 1205
Private Const ERROR_CONNECTION_UNAVAIL As Long = 1201
Private Const ERROR_DEVICE_ALREADY_REMEMBERED As Long = 1202
Private Const ERROR_DEVICE_IN_USE As Long = 2404
Private Const ERROR_EXTENDED_ERROR As Long = 1208
Private Const ERROR_INTERNAL_ERROR As Long = 1359
Private Const ERROR_INVALID_FUNCTION As Long = 1
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_INVALID_PASSWORD As Long = 86
Private Const ERROR_INVALID_PRINTER_NAME As Long = 1801
Private Const ERROR_LOCAL_DRIVE As Long = 2250              ' VB4 Unleashed from internet - not in Win32API.txt
Private Const ERROR_MORE_DATA As Long = 234
Private Const ERROR_NETWORK_UNREACHABLE As Long = 1231
Private Const ERROR_NO_NET_OR_BAD_PATH As Long = 1203
Private Const ERROR_NO_NETWORK As Long = 1222
Private Const ERROR_NO_RESOURCE_NAME As Long = 487          ' VB4 Unleashed from internet - not in Win32API.txt
Private Const ERROR_NOT_CONNECTED As Long = 2250
Private Const ERROR_OPEN_FILES As Long = 2401
Private Const ERROR_OUTOFMEMORY  As Long = 14

Public Function NetConnectShareDisk(ByVal pRemoteName As String, _
                                    ByRef pErrMsg As String, _
                           Optional ByVal pUsr As String = vbNullString, _
                           Optional ByVal pPwd As String = vbNullString) As Boolean
                       
'   Caters for caller passing un-initialised/empty pErrMsg parameter

' The WNetAddConnection2 function makes a connection to a network resource.
' The function can redirect a local device to the network resource.
' If you can pass a handle to a window that the provider of network resources
' can use as an owner window for dialog boxes, call the WNetAddConnection3 function instead. AUrban WILL NEED TO GET SOME DOCO ON THIS AND DO IT LATER
Dim bSuccess As Boolean
Dim lngErrCode As Long
Dim strErrMsg As String
Dim udtNetResource As NETRESOURCE
    
    With udtNetResource
        .lngScope = RESOURCE_GLOBALNET  ' To specify scope during enumeration
        .lngType = RESOURCETYPE_DISK    ' Defines the type of resource
    '   .lngDisplayType =               ' How resources will be displayed
    '   .lngUsage =                     ' Specifies the resource usage
    '   .strLocalName                   ' Local device for the connection.
                                        ' Specifies the name of a local device to be redirected (eg "X:", "COM1")
' AUrban NEEDS work and THOUGHT        .strLocalName = "X:"    ' A string that specifies the name of a local device to be redirected, such as "X:"
                        ' or "COM1". The case of the string is unimportant. If the string is empty or NULL,
                        ' the function makes a connection but does not redirect the resource to a local device.
        .strRemoteName = pRemoteName    ' Indicates the network resource (eg "\\ServerName\ShareName")
    '   .strComment =   ' For a provider-supplied comment
    '   .strProvider =  ' Name of provider who owns the resource
                        ' If not NULL then O/S attempts to make a connection only to the named network provider.
                        ' You should set this member only if you know for sure which network provider you want
                        ' to use. Otherwise, let the O/S determine which provider the network name maps to
    End With
    
'   Call to WNetAddConnection2()
'    udtNetResource: Populated above
'    If User and Password arguments are NULL, the user context for the process provides the default user name
'    The final parameter specifies connection options. When this parameter is set to CONNECT_UPDATE_PROFILE,
'    the network resource connection is remembered, and Windows automatically attempts to restore the connection
'    when the user logs on. If this parameter is 0, the user's profile is not updated and Windows will not
'    automatically restore the connection at logon.

'   The program should make no changes to the user profile therefore
'   we pass 0 as last parameter instead of CONNECT_UPDATE_PROFILE
'   We don't want it attempting a reconnection when the user logs on again.

'   If the Username and Password arguments are null, the user context
'   for the process provides the default user name
' lpPassword

' WNetAddConnection2
' ------------------
'   lpPassword [in]
'       Pointer to a constant null-terminated string
'       that specifies a password to be used in making the network connnection.
'
'       If lpPassword is NULL, the function uses the current default
'       password associated with the user specified by the lpUserName
'       If lpPassword points to an empty string, the function does not use a password
'
'   lpUserName [in]
'       Pointer to a constan null-terminated string that specifies a user name for making the connection.
'
'       If lpUserName is NULL, the function uses the default user name. (The user context for the process
'       provides the default user name.)
'       The lpUserName parameter is specified when users want to connect to a network resource for
'       password have been assigned a user name or account other than the default user name or account.
'
'       The user-name string represents a security context.
'       *********************************************************************
'       * WINDOWS ME/98/95: This parameter must be NULL or an empty string. *
'       *********************************************************************
'
'  dwFlags [in]
'       Connection options. The following values are currently defined
'   0                        System does not update information about the connection.
'   CONNECT_UPDATE_PROFILE   If the connection was marked as persistent in the registry, the system continues
'                            to restore the connection at the next logon. If the connection was NOT marked as
'                            persistent the function ignores CONNECT_UPDATE_PROFILE flag.

''''   UNFORTUNATELY IsDestinationReachable() BOMBS OUT OVER VPN WHEN A DESTINATION IS UNREACHABLE
''''   SOME COMMENTS I HAVE READ OVER THE NET SUGGEST THE API CALL IT MAKES (IsDestinationReachableA) MAY
''''   CONFLICT WITH SOME VPN SOFTWARE. I KNOW IT CONFLICTS WITH Cisco VPN software (version 3.5.4).
''''   Test remote folder can be reached with IsDestinationReachable before attempting to connect to it.
''''   If remote folder can't be reached the test will fail quickly (~ 1 sec) where a failed connection
''''   attempt with WNetAddConnection2 may take a long time waiting to timeout.
''''   (Note. In some cases where no network of any type is present WNetAddConnection2 may fail quickly)
    
'''    If Not IsDestinationReachable(pRemoteName) Then
'''        strErrMsg = DQ(pRemoteName) & " cannot be reached."
'''    Else
        lngErrCode = WNetAddConnection2(lpNetResource:=udtNetResource, lpPassword:=pPwd, lpUsername:=pUsr, dwFlags:=0)
        If lngErrCode = 0 Then
            bSuccess = True
        Else
            strErrMsg = NetError(lngErrCode)
        End If
'''    End If
   
    pErrMsg = strErrMsg
    NetConnectShareDisk = bSuccess
    
End Function

Public Function NetDisconnectShare(ByVal pRemoteOrLocalName As String, ByRef pErrMsg As String) As Boolean
'   Caters for caller passing un-initialised/empty pErrMsg parameter
Dim bSuccess As Boolean
Dim lngErrCode As Long
Dim strErrMsg As String

'' AUrban Might be the case that fForce should be changed on successive attempts. If I use
''        the flag as originally used then I can report errors where files are not closed
''        and progressively deal with the errors.
''        If I use successive attempts, one with the parameter as false and one with it true then
''        I could report the error without jeopardizing the smooth flow of current code.

' WNetCancelConnection2
' ---------------------
'  lpName  May specify either lpRemoteName/strRemoteName or lpLocalName/strLocalName ("\\ServerName\ShareName" or "X:")
'
'  dwFlags
'   0                        System does not update information about the connection.
'   CONNECT_UPDATE_PROFILE   If the connection was marked as persistent in the registry, the system continues
'                            to restore the connection at the next logon. If the connection was NOT marked as
'                            persistent the function ignores CONNECT_UPDATE_PROFILE flag.
'  fForce
'   True     Close open files or jobs so share can be cancelled
'   False    Fail if open files or jobs

'   Program should make no changes to user profile -> pass 0 as 2nd parameter instead of CONNECT_UPDATE_PROFILE
    lngErrCode = WNetCancelConnection2(lpName:=pRemoteOrLocalName, dwFlags:=0, fForce:=True)
    
    If lngErrCode = 0 Then
        bSuccess = True
    Else
        bSuccess = False
        strErrMsg = NetError(lngErrCode)
    End If

    pErrMsg = strErrMsg
    NetDisconnectShare = bSuccess
    
End Function

Public Function NetError(ByVal pErrCode As Long) As String
'   The WNetGetLastError function retrieves the most recent
'   extended error code set by a WNet function. The network
'   provider reported this error code; it will not generally
'   be one of the errors included in the SDK header file WinError.h.

Dim strResult As String

    Select Case pErrCode
        Case ERROR_ACCESS_DENIED: strResult = _
            "The caller does not have access to the network resource." ' MS Online Help
        '   "ERROR_ACCESS_DENIED The caller does not have access to the network resource." ' MS Online Help
        '   "ERROR_ACCESS_DENIED Access is denied" ' VB4 Unleashed
        
        Case ERROR_ALREADY_ASSIGNED: strResult = _
            "The local device specified by the lpLocalName member is already connected to a network resource." ' MS Online Help
        '   "ERROR_ALREADY_ASSIGNED The local device specified by the lpLocalName member is already connected to a network resource." ' MS Online Help
        '   "ERROR_ALREADY_ASSIGNED Local Drive Identifier Is In Use. / The device specified in the strLocalName parameter already is connected" ' VB4 Unleashed
        
        Case ERROR_ALREADY_CONNECTED: strResult = _
            "Local Device Already Connected." ' VB4 Unleashed
        '   "ERROR_ALREADY_CONNECTED Local Device Already Connected." ' VB4 Unleashed
        
        Case ERROR_BAD_DEV_TYPE: strResult = _
            "The type of local device and the type of network resource do not match." ' MS Online Help
        '   "ERROR_BAD_DEV_TYPE The type of local device and the type of network resource do not match." ' MS Online Help
        '   "ERROR_BAD_DEV_TYPE The device type and the resource type do not match" ' VB4 Unleashed
        
        Case ERROR_BAD_DEVICE: strResult = _
            "The value specified by lpLocalName is invalid." ' MS Online Help
        '   "ERROR_BAD_DEVICE The value specified by lpLocalName is invalid." ' MS Online Help
        '   "ERROR_BAD_DEVICE Value specified in strLocalName is invalid"     ' VB4 Unleashed
        
        Case ERROR_BAD_NET_NAME: strResult = _
            "The value specified by the lpRemoteName member is not acceptable to any network resource provider, either because the resource name is invalid, or because the named resource cannot be located." ' MS Online Help
        '   "ERROR_BAD_NET_NAME The value specified by the lpRemoteName member is not acceptable to any network resource provider, either because the resource name is invalid, or because the named resource cannot be located." ' MS Online Help
        '   "ERROR_BAD_NET_NAME Value specified in the strRemoteName parameter is not valid or cannot be located." ' VB4 Unleashed
        
        Case ERROR_BAD_NETPATH: strResult = _
            "The network path was not found." ' MS Online Help
        '   "ERROR_BAD_NETPATH The network path was not found." ' MS Online Help
        
        Case ERROR_BAD_PROFILE: strResult = _
            "The user profile is in an incorrect format." ' MS Online Help
        '   "ERROR_BAD_PROFILE The user profile is in an incorrect format." ' MS Online Help
        '   "ERROR_BAD_PROFILE System is unable to open the user profile to process persistent connections" ' VB4 Unleashed
        
        Case ERROR_BAD_PROVIDER: strResult = _
            "The value specified by the lpProvider member does not match any provider." ' MS Online Help
        '   "ERROR_BAD_PROVIDER The value specified by the lpProvider member does not match any provider." ' MS Online Help
        
        Case ERROR_BAD_USERNAME: strResult = _
            "No Such User." ' VB4 Unleashed
        '   "ERROR_BAD_USERNAME No Such User." ' VB4 Unleashed
        
        Case ERROR_BUSY: strResult = _
            "The router or provider is busy, possibly initializing. The caller should retry." ' MS Online Help
        '   "ERROR_BUSY The router or provider is busy, possibly initializing. The caller should retry." ' MS Online Help
        '   "ERROR_BUSY The router or provider is busy, possibly congested or initializing. The function should be retried." ' VB4 Unleashed
        
        Case ERROR_CANCEL_VIOLATION: strResult = _
            "Cancel Violation"
        '   "ERROR_CANCEL_VIOLATION Cancel Violation"
        
        Case ERROR_CANCELLED: strResult = _
            "The attempt to make the connection was cancelled by the user through a dialog box from one of the network resource providers, or by a called resource." ' MS Online Help
        '   "ERROR_CANCELLED The attempt to make the connection was cancelled by the user through a dialog box from one of the network resource providers, or by a called resource." ' MS Online Help
        '   "ERROR_CANCELLED The connection attempt was cancelled by the user through a dialog box from one of the network resource providers, or by a called resource or other process." ' VB4 Unleashed
        
        Case ERROR_CANNOT_OPEN_PROFILE: strResult = _
            "The system is unable to open the user profile to process persistent connections."  ' MS Online Help
        '   "ERROR_CANNOT_OPEN_PROFILE The system is unable to open the user profile to process persistent connections."  ' MS Online Help
        
        Case ERROR_CONNECTION_UNAVAIL: strResult = _
            "The Resource Is Not Shared." ' VB4 Unleashed
        '   "ERROR_CONNECTION_UNAVAIL The Resource Is Not Shared." ' VB4 Unleashed
        
        Case ERROR_DEVICE_ALREADY_REMEMBERED: strResult = _
            "An entry for the device specified by lpLocalName is already in the user profile."  ' MS Online Help
        '   "ERROR_DEVICE_ALREADY_REMEMBERED An entry for the device specified by lpLocalName is already in the user profile."  ' MS Online Help
        '   "ERROR_DEVICE_ALREADY_REMEMBERED An entry for the device specified in strLocalName is already in the user profile"  ' VB4 Unleashed
        
        Case ERROR_DEVICE_IN_USE: strResult = _
            "The device is in use by an active process and cannot be disconnected."           ' MS Online Help
        '   "ERROR_DEVICE_IN_USE The device is in use by an active process and cannot be disconnected."           ' MS Online Help
        '   "ERROR_DEVICE_IN_USE The specified device is in use by an active process and cannot be disconnected"  ' VB4 Unleashed
        
        Case ERROR_INTERNAL_ERROR: strResult = _
            "Internal Error!"  ' VB4 Unleashed
        '   "ERROR_INTERNAL_ERROR Internal Error!"  ' VB4 Unleashed
        
        Case ERROR_INVALID_FUNCTION: strResult = _
            "Function is not supported." ' VB4 Unleashed
        '   "ERROR_INVALID_FUNCTION Function is not supported." ' VB4 Unleashed
            
        Case ERROR_INVALID_PARAMETER: strResult = _
            "The parameter is incorrect. ' MS Online Help"
        '   "ERROR_INVALID_PARAMETER The parameter is incorrect. ' MS Online Help"

        Case ERROR_INVALID_PASSWORD: strResult = _
            "The specified password is invalid and the CONNECT_INTERACTIVE flag is not set."  ' MS Online Help
        '   "ERROR_INVALID_PASSWORD The specified password is invalid and the CONNECT_INTERACTIVE flag is not set."  ' MS Online Help
        '   "ERROR_INVALID_PASSWORD Specified password is invalid"  ' VB4 Unleashed
        
        Case ERROR_INVALID_PRINTER_NAME: strResult = _
            "No Such Printer."     ' VB4 Unleashed
        '   "ERROR_INVALID_PRINTER_NAME No Such Printer."     ' VB4 Unleashed
        
        Case ERROR_LOCAL_DRIVE: strResult = _
            "Can't Disconnect Local Drive." ' VB4 Unleashed
        '   "ERROR_LOCAL_DRIVE Can't Disconnect Local Drive." ' VB4 Unleashed
        
        Case ERROR_MORE_DATA: strResult = _
            "Buffer to small to hold network name, make lpnLength bigger"
        '   "ERROR_MORE_DATA Buffer to small to hold network name, make lpnLength bigger"
        
        Case ERROR_NETWORK_UNREACHABLE: strResult = _
            "The network location cannot be reached. For information about network troubleshooting, see Windows Help." ' MS Online Help
        '   "ERROR_NETWORK_UNREACHABLE The network location cannot be reached. For information about network troubleshooting, see Windows Help." ' MS Online Help
        
        Case ERROR_NO_NET_OR_BAD_PATH: strResult = _
            "The operation cannot be performed because a network component is not started or because a specified name cannot be used." ' MS Online Help
        '   "ERROR_NO_NET_OR_BAD_PATH The operation cannot be performed because a network component is not started or because a specified name cannot be used." ' MS Online Help
        '   "ERROR_NO_NET_OR_BAD_PATH Operation cannot be performed because either a network component is not started or the specified name cannot be used" ' VB4 Unleashed
        
        Case ERROR_NO_NETWORK: strResult = _
            "The network is unavailable."  ' MS Online Help
        '   "ERROR_NO_NETWORK No network is present"        ' VB4 Unleashed
        
        Case ERROR_NO_RESOURCE_NAME: strResult = _
            "Must Enter A Valid Network Resource Name." ' VB4 Unleashed
        '   "ERROR_NO_RESOURCE_NAME Must Enter A Valid Network Resource Name." ' VB4 Unleashed
        
        Case ERROR_NOT_CONNECTED: strResult = _
            "The name specified by the lpName parameter is not a redirected device, or the system is not currently connected to the device specified by the parameter." ' MS Online Help
        '   "ERROR_NOT_CONNECTED The name specified by the lpName parameter is not a redirected device, or the system is not currently connected to the device specified by the parameter." ' MS Online Help
        '   "ERROR_NOT_CONNECTED The name specified by the strName parameter is not a redirected device, or the system is not currently connected to the device specified by the parameter" ' VB4 Unleashed
        
        Case ERROR_OPEN_FILES: strResult = _
            "Files open and the force parameter is false"
        '   "ERROR_OPEN_FILES Files open and the force parameter is false"
        '   "ERROR_OPEN_FILES There are open files or uncompleted processes, and the bForce parameter is FALSE"   ' VB4 Unleashed There are open files or uncompleted processes, and the bForce parameter is FALSE
        
        Case ERROR_OUTOFMEMORY: strResult = _
            "Out of Memory." ' VB4 Unleashed
        '   "ERROR_OUTOFMEMORY Out of Memory." ' VB4 Unleashed
            
        '--------------------------------------------------------------------------------------------------------------
        Case ERROR_EXTENDED_ERROR:  strResult = _
            "ERROR_EXTENDED_ERROR " & NetGetLastExtendedError()
        '   VB4 Unleashed A network-specific error occurred
        '   WNetGetLastError retrieves the most recent extended error code set by a Windows network function.
        '   An extended error code is network-specific; that is, it is supplied by a particular network
        '   provider—Novell or Banyan, for example—and therefore contains error information that pertains
        '   only to that provider's network protocol.
        '--------------------------------------------------------------------------------------------------------------
        
        Case Else:
            strResult = "Unrecognized Network Error - Code " & pErrCode
            
    End Select
    
'   strResult = Prefix & pErrCode & " - " & strResult
    
    NetError = strResult
    
End Function

Private Function NetGetLastExtendedError() As String
Dim lngErrCode As Long        ' Stores the error code
Dim lngErrDescLength As Long  ' Size of error description variable
Dim lngProviderLength As Long ' Size of provider name variable
Dim strErrDesc As String      ' Stores error description
Dim strProvider As String     ' Stores network provider name
Dim strResult As String

'   Prepare return variables for API call to WNetGetLastError
    strErrDesc = Space$(1024)           ' VB terminates the string with a NULL character [ie Chr$(0)]
    lngErrDescLength = Len(strErrDesc)
    strProvider = Space$(255)           ' VB terminates the string with a NULL character [ie Chr$(0)]
    lngProviderLength = Len(strProvider)
    
    If WNetGetLastError(lngErrCode, strErrDesc, lngErrDescLength, strProvider, lngProviderLength) = NO_ERROR Then
    '   Clean up Null terminated return strings & build network specific error message
        strErrDesc = Left$(strErrDesc, InStr(1, strErrDesc, vbNullChar, vbBinaryCompare) - 1)
        strProvider = Left$(strProvider, InStr(1, strProvider, vbNullChar, vbBinaryCompare) - 1)
        strResult = strErrDesc & " [Network Provider: " & strProvider & "][ExtdErr " & lngErrCode & "]"
    Else
        strResult = "Unable To Process Extended Error!"   ' Return generic msg.
    End If
    
    NetGetLastExtendedError = strResult

End Function

'Public Function UseConnection(ByVal pRemoteName As String)

''*  At times, it is necessary to map a drive letter to a network share.
''*  There are several API functions that can be used to accomplish this,
''*  such as WNetAddConnection, WNetAddConnection2, WNetAddConnection3,
''*  and WNetUseConnection. The primary difference is that with the
''*  WNetUseConnection function, you do not need to specify the drive letter
''*  to be used, while it is required with the other API functions.
'Dim NetR As NETRESOURCE    ' NetResouce structure
'Dim lngErrCode As Long        ' Return value from API
'Dim buffer As String       ' Drive letter assigned to resource
'Dim bufferlen As Long      ' Size of the buffer
'Dim success As Long        ' Additional info about API call
'
'   ' Initialize the NetResouce structure
'   NetR.lngScope = RESOURCE_GLOBALNET
'   NetR.lngType = RESOURCETYPE_DISK
'   NetR.lngDisplayType = RESOURCEDISPLAYTYPE_SHARE
'   NetR.lngUsage = RESOURCEUSAGE_CONNECTABLE
'   NetR.strLocalName = vbNullString
'   NetR.strRemoteName = txtUNC.Text
'
'   ' Initialize the return buffer and buffer size
'   buffer = Space(32)
'   bufferlen = Len(buffer)
'
'   ' Call API to map the drive
'   lngErrCode = WNetUseConnection(Me.hWND, NetR, txtPWD.Text, txtUser.Text, _
'      CONNECT_REDIRECT, buffer, bufferlen, success)
'
'   ' Check if call to API failed. According to the MSDN help, there
'   ' are some versions of the operating system that expect the userid
'   ' as the 3rd parameter and the password as the 4th, while other
'   ' versions of the operating system have them in reverse order, so
'   ' if first call to API fails, try reversing these two parameters.
'   If lngErrCode <> NO_ERROR Then
'      ' Call API with userid and password switched
'      lngErrCode = WNetUseConnection(Me.hWND, NetR, txtUser.Text, _
'         txtPWD.Text, CONNECT_REDIRECT, buffer, bufferlen, success)
'   End If
'
'   ' Check for success
'   If (lngErrCode = NO_ERROR) And (success = CONNECT_LOCALDRIVE) Then
'      ' Store the mapped drive letter for later usage
'      MappedDrive = Left$(buffer, InStr(1, buffer, ":"))
'
'      ' Display the mapped drive letter
'      MsgBox "Connect Succeeded to " & MappedDrive
'   Else
'      MsgBox "ERROR: " & str(lngErrCode) & " - Connect Failed!"
'   End If
'
''Sub ShowDriveList()
''    Dim fs, d, dc, s, n
''    Set fs = CreateObject("Scripting.FileSystemObject")
''    Set dc = fs.Drives
''    For Each d In dc
''        s = s & d.DriveLetter & " - "
''        If d.DriveType = 3 Then
''            n = d.ShareName
''        Else
''            n = d.VolumeName
''        End If
''        s = s & n & vbCrLf
''    Next
''    MsgBox s
''End Sub
'End Function
