VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------------------
' NOTES FROM CuteFTP HELP '
' -------------------------

' Transfer Engine
' -----------------------------------------------
' You can set how long the Transfer Engine will leave the data connection open after a completed transfer.
' In the CuteFTP GUI, go to Tools > Global Options > Transfer Settings and set
' the time in Close the file transfer n seconds after the transfer is completed.

' Any time you call a Transfer function, the connect function is called indirectly.
' It is recommended you explicitly invoke the Connect function though it is not necessary.
' You can always use the IsConnected property to determine whether you are connected at any given time.

' Upload Method: Upload(BSTR strLocalName ,BSTR strRemoteName , long nMultiPartNumber)
' ------------------------------------------------------------------------------------
' strRemoteName
'   This is optional, use it only if you want to change the destination name or path for the uploaded files or folders.
'   You can use absolute or relative paths with or without wildcards.
' strLocalName
'   This is the path to the local item you are uploading. You can use absolute or relative paths with or without wildcards.
' nMultiPartNumber
'   Use this to split the upload into multiple parts.  The default value = 1.  The value specifies the number of parts used for the download.

' Overwriting files on upload
' ----------------------------
' If implementing overwriting we need a way to supress messages.- see CuteFTP CLIENT AND BATA FTP SERVER
' When attempting to upload a file over an existing file and using the default CuteFTP option settings,
' the TE object gives an interactive prompt 'Confirm Replace File?'.
' Code to handle pOverwrite True will be written if and when it is needed (13May2008)
' The following TE object members may be of use: RemoteExists, DeleteDestination, GetList, GetResult
'------------------------------------------------------------------------------------------------------
'NB Upload2 method described in help does not exist (confirmed in CuteFTP user forum as at 01Jul2008) '
'------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------
' CuteFTP CLIENT AND BATA FTP SERVER (Optimisation efforts prompted by slow test BATA SFTP Server)
'-----------------------------------------------------------------------------------------------------------------
'1. RemoteExists method was PROHIBITIVELY time consuming until night of 02Jul2008. BATA cutover to
'   live on night of 01Jul2008 and uploading failed that night becuase of problems with FTP server.
'   From Test to Live configuration uploading using RemoteExists test went from 30 seconds to 2 seconds.
'   BEFORE THIS IMPROVEMENT OCCURED OPTIMISATIONS WERE MADE TO NOT USE RemoteExists.
'   THOUGH THIS WOULD NOT NOW BE SELECTED FOR OPTIMISATION, THE OPTIMISATION WAS TRIALLED.
'   IMPROVEMENTS WERE NOT SIGNIFICANT ENOUGH WITH NEW FTP SERVER TO WARRANT INCLUSION GIVEN
'   THAT THEY WERE RELIANT ON CuteFTP CONFIGURATION AND ADDED COMPLEXITY TO THE CODE
'2. Upload2 method described in help does not exist (confirmed in CuteFTP user forum as at 01Jul2008)
'3. To avoid using RemoteExists method, the obvious alternative is to attempt uploading and handle any errors.
'   This approach has the following limitations. You can supress showing a VB runtime error message box by
'   using the ThrowError option of CuteFTPPro.TEConnection object [eg .Option("ThrowError") = False] BUT
'   The only way of running the Transfer Engine with interactive prompts disabled is to launch the TE manually
'   and supply the -noprompts parameter which requires a number of assumptions (path, etc)
'   eg "C:\Program Files\Globalscape\CuteFTP Professional\TE\ftpte.exe" -noprompts.
'   Suppressing prompts may cause lockups if no default action is available for the event in question,
'   or if the prompt was produced by an error.
'   The prompt we are trying to avoid is 'Confirm Replace File?' which is the default configuration on installation.
'   CuteFTP client options supported for Smart Overwrite of matching files in combo with BATA SFTP 02June2008:-
'   Prompt, Skip, Numerate, ... (Overwrite not currently supported)
'   NB. When TE configured for Overwrite and overwrite fails, returned error msg relates to insufficient privileges
'   -----------------------------------------------------------------------------------------------------------------
'   VB runtime error message box supressed by mFtpCn.Option("ThrowError") = False in Class_Initialize
'   If Upload fails code execution will continue and mFtpCn.ErrorDescription will contain the error
'   (Depending on configuration of TE an interatctive prompt may display and time out)
'   -----------------------------------------------------------------------------------------------------------------

Public Enum eTransferType
    eTfr_ASCII  ' All files should be transferred in ASCII mode
    eTfr_BINARY ' All files should be transferred in BINARY mode
    eTfr_AUTO   ' (CuteFTP default)
                ' The TE will reference an internal list editable from CuteFTP's Global Options
                ' to determine the proper transfer type for that particular file. For example,
                ' if the ASCII list contains a filter mask of "txt" then all files with extension
                ' ".txt" will be transferred in ASCII. If a file doesn't correspond to any mask
                '  in this list then it will be transferred in binary mode.
End Enum

Public Enum eProtocol
    eFTP = 1            ' "FTP"             File Transfer Protocol
    eFTPS = 2           ' "FTPS"            FTP using SSL in explicit mode (standard port 21)
    eFTPS_IMPLICIT = 3  ' "FTPS_IMPLICIT"   Extension of FTP using SSL -- FTP using SSL in implicit mode (specific port)
    eSFTP = 4           ' "SFTP"            Secure protocol based on SSH2
    eFTP_SKEY_MD4 = 5   ' "FTP_SKEY_MD4"    Secure one time password login using MD4
    eFTP_SKEY_MD5 = 6   ' "FTP_SKEY_MD5"    Secure one time password login using MD5
    eHTTP = 7           ' "HTTP"            Hypertext Transfer Protocol
    eHTTPS = 8          ' "HTTPS"           HTTP with SSL
End Enum
    

Dim mFtpCn As CuteFTPPro.TEConnection

Public Function RemoteExists(ByVal pRemoteName As String, Optional ByRef pErrMsg As String) As Boolean
Dim bResult As Boolean
Dim bConnected As Boolean
Dim strErrMsg As String

    With mFtpCn
    '   Take care when programming to connect to remote site.
    '   Within IDE it seems IsConnected property can sometimes incorrectly return False
    '   and continue to return False after successful calls to the Connect method
        bConnected = .IsConnected
        If Not bConnected Then
            bConnected = .Connect()
        End If
        
        If Not bConnected Then
            strErrMsg = .ErrorDescription
        Else
            bResult = .RemoteExists(pRemoteName)
        End If
    End With
    
    pErrMsg = strErrMsg
    RemoteExists = bResult
    
End Function

Public Function Upload(ByVal pLocalName As String, _
                       ByVal pRemoteName As String, _
              Optional ByRef pErrMsg As String) As Boolean
'~            Optional ByVal pOverwrite As Boolean = False, _
'~ pOverwrite default to False b/c is behaviour we want when transferring to Bata (at least initially)
Dim bResult As Boolean
Dim bConnected As Boolean
Dim strErr As String

    With mFtpCn
    '---------------------------------------------------------------------------------------------------------
    '   Any time you call a Transfer function, the connect function is called indirectly.                    '
    '   It is recommended you explicitly invoke the Connect function though it is not necessary.             '
    '   You can always use the IsConnected property to determine whether you are connected at any given time.'
    '---------------------------------------------------------------------------------------------------------
    '   Within IDE it seems IsConnected property sometimes incorrectly returns False '
    '   and continues to return False after successful calls to the Connect method   '
    '---------------------------------------------------------------------------------
        bConnected = .IsConnected
        If Not bConnected Then
            bConnected = .Connect()
        End If
        
        If Not bConnected Then
            strErr = .ErrorDescription
'~      ElseIf pOverwrite Then
'~      '   If implementing overwriting see module header notes for hints
'~          MsgBox "NOT CALLED THIS WAY AT THE MOMENT. IF THIS BOX IS SHOWN THEN CODE NEEDS TO BE CHANGED"
'~          strErr = "Overwrite not yet supported."
        ElseIf .RemoteExists(pRemoteName) Then
            strErr = "Remote file exists: " & pRemoteName
        ElseIf Not .Upload(bstrLocalName:=pLocalName, bstrRemoteName:=pRemoteName) Then
            strErr = .ErrorDescription
        Else
            bResult = True
        End If
    
    End With
    
    pErrMsg = strErr
    Upload = bResult

End Function

Private Sub Class_Initialize()
'-------------------------------------------------------------------------------------------
' Improvement? - How to suppress dialog box when incorrect pwd given [.Password = "ZUrban"]'
'-------------------------------------------------------------------------------------------

'   SSL enabled server
'   ------------------
'   If your script is connecting to an SSL enabled server, double check the steps listed under Handling SSL Certificates.
'   The primary cause for a locked TE during an SSL connection is the fact that the TE could not locate the
'   client or server certificate, or the client certificate does not reside in the server’s trusted list.

'.  Create FtpCn object and initialise with default values
'
    Set mFtpCn = New CuteFTPPro.TEConnection
    With mFtpCn
    ' ThrowError Option
    '   True(default): TE COM will call AtlReportError (showing a VB runtime error message box) on
    '                  connection/transfer/IO/other error which will terminate script execution/program
    '   False: Disable ATL (asynchronus task list?) exceptions
    '          If some transfer method fails the script will continue to the next command.
        .Option("ThrowError") = False
    ' AutoCloseMethod Option
    '   Handles hidden prompts (such as SSL Accept Cert Prompt when running a script while not logged in)
    '   0: Don't accept (default), 1: auto accept, 2: auto reject
    '   eg. .Option("AutoCloseMethod") = 1  'will auto accept cert and continue
    '   Tell TE to auto-accept SSL or SSH certificates using the AutoCloseMethod parameter of the Option method.
        .Option("AutoCloseMethod") = 1
    ' AutoCloseDelay Option (default value is 60 seconds)
    '   Time delay before CuteFTP should perform the action specified by the AutoCloseMethod option.
    ' Set number of retries
    '   Set to 1 to let code control retries rathter than TE going off and doing it's own retries
        .Retries = 1
    ' Set delay between retries
    '   .Delay = 2 ' delay in seconds - may want to set a default delay
    
    '   BATA SFTP specific settings
        .Protocol = "SFTP"
        Me.TransferType = eTransferType.eTfr_ASCII
        
    End With
    
End Sub

Public Property Let Password(ByVal pPassword As String)
    mFtpCn.Password = pPassword
End Property

Public Property Let HostAddress(ByVal pHostAddress As String)
    mFtpCn.Host = pHostAddress
End Property

Public Property Let Login(ByVal pLogin As String)
    mFtpCn.Login = pLogin
End Property

Public Property Let TransferType(ByVal pTransferType As eTransferType)
    Select Case pTransferType
        Case eTfr_ASCII
            mFtpCn.TransferType = "AUTO"    ' All files should be transferred in ASCII mode
        Case eTfr_BINARY
            mFtpCn.TransferType = "BINARY"  ' All files should be transferred in BINARY mode
        Case eTfr_AUTO
            mFtpCn.TransferType = "ASCII"   ' The TE will reference an internal list editable from CuteFTP's Global Options
                                            ' to determine the proper transfer type for that particular file. For example,
                                            ' if the ASCII list contains a filter mask of "txt" then all files with extension
                                            ' ".txt" will be transferred in ASCII. If a file doesn't correspond to any mask
                                            ' in this list then it will be transferred in binary mode.
    End Select
    
End Property

Private Sub Class_Terminate()
    mFtpCn.Disconnect   ' Not testing with IsConnected as I have found it to give false negatives in the IDE

    mFtpCn.Close ("Exit No Pending")   ' Engine
    Set mFtpCn = Nothing
    DoEvents    ''' Not sure if I need to do this yet & perhaps can remove the wimpy commands above - will possibly try later
                ''' Clearing object variable added b/c I got an error when one object/code segment using and then disposiing
                ''' of this object was followed by antoher which started Ok but then got the following error
                ''' ("Class does not support Automation or does not support expected interface")  which suggests the
                ''' object may no longer exist - perhaps a slow closing by previous calls. It might be that I need to
                ''' add the DoEvents subsequent to the clearing object variable to ensure that the TE is removed
End Sub

Public Property Get Protocol() As eProtocol
Dim eResult As eProtocol
    
    Select Case mFtpCn.Protocol
        Case "FTP"
        '   File Transfer Protocol
            eResult = eFTP
        Case "FTPS" = 2
        '   FTP using SSL in explicit mode (standard port 21)
            eResult = eFTPS
        Case "FTPS_IMPLICIT"
        '   Extension of FTP using SSL -- FTP using SSL in implicit mode (specific port)
            eResult = eFTPS_IMPLICIT
        Case "SFTP"
        '   Secure protocol based on SSH2
            eResult = eSFTP
        Case "FTP_SKEY_MD4"
        '   Secure one time password login using MD4
            eResult = eFTP_SKEY_MD4
        Case "FTP_SKEY_MD5"
        '   Secure one time password login using MD5
            eResult = eFTP_SKEY_MD5
        Case "HTTP"
        '   Hypertext Transfer Protocol
            Case Is = eHTTP
        Case "HTTPS"
        '   HTTP with SSL
            eResult = eHTTPS
    End Select
    
    Protocol = eResult
    
End Property

Public Property Let Protocol(ByVal pProtocol As eProtocol)
Dim strProtocol As String

    Select Case pProtocol
        Case eFTP
        '   File Transfer Protocol
            strProtocol = "FTP"
        Case eFTPS
        '   FTP using SSL in explicit mode (standard port 21)
            strProtocol = "FTPS"
        Case eFTPS_IMPLICIT
        '   Extension of FTP using SSL -- FTP using SSL in implicit mode (specific port)
            strProtocol = "FTPS_IMPLICIT"
        Case eSFTP
        '   Secure protocol based on SSH2
            strProtocol = "SFTP"
        Case eFTP_SKEY_MD4
        '   Secure one time password login using MD4
            strProtocol = "FTP_SKEY_MD4"
        Case eFTP_SKEY_MD5
        '   Secure one time password login using MD5
            strProtocol = "FTP_SKEY_MD5"
        Case eHTTP
        '   Hypertext Transfer Protocol
            Case Is = "HTTP"
        Case eHTTPS
        '   HTTP with SSL
            strProtocol = "HTTPS"
    End Select
    
    mFtpCn.Protocol = strProtocol

End Property

'Public Function Connect() As Boolean
'' Any time you call a Transfer function, the connect function is called indirectly.
'' It is recommended you explicitly invoke the Connect function though it is not necessary.
'' You can always use the IsConnected property to determine whether you are connected at any given time.
'
''------------------------------------------------------------------
''   Use the Connect method to log in to the remote server.
''   Before calling it, you set the protocol, host address,
''   user name, and password to establish a successful connection.
''   -> is probably an idea to check here that all parameters have been supplied
'
''   Connect method returns a long value of zero on failure
'    Connect = mFtpCn.Connect()
''------------------------------------------------------------------
'End Function

'Public Property Get ErrorDescription() As String
'    ErrorDescription = mFtpCn.ErrorDescription
'End Property

'Public Property Let LocalFolder(ByVal pLocalFolder As String)
'    mFtpCn.LocalFolder = pLocalFolder
'End Property

'Public Property Let RemoteFolder(ByVal pRemoteFolder As String)
'    mFtpCn.RemoteFolder = pRemoteFolder
'End Property

'Public Property Let Retries(ByVal pRetries As Long)
'    mFtpCn.Retries = pRetries
'End Property

'Public Property Let Delay(ByVal pDelay As Long)
'   From help -> int Object.Delay
'End Property

'Public Property Get Port() As Variant
'    Port = mFtpCn.Protocol
'End Property
'
'Public Property Let Port(ByVal vNewValue As Variant)
''   "FTPS" protocol is documented by CuteFTP as: using SSL in explicit mode (standard port 21)
''   Not sure if setting port for this protocol is ignored or causes cuteFTP component to EXPLODE!
'    mFtpCn.Protocol = Port
'End Property

