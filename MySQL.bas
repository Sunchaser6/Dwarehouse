Attribute VB_Name = "MySQL"
Option Explicit
' --------------------------------------------------------------------------------------------------------------------------------------
' MySQL Connection String (Connector/ODBC)
'  Option Parameters
'   All sample connection strings use Option = 3. MySQL manuals recommend and explain
'   setting second bit (ie 2 component of 3) but not the least significatnt bit
'   Added ability to apply multiple SQL statements separated by a semi-colon
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   Recommended Connector/ODBC Option Values for Different Configurations
'   Configuration                   Parameter Settings  Option Value
'   Microsoft Access, Visual Basic  FOUND_ROWS=1;       2
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   Parameter Name      GUI Option                      Constant Value  Description
'   FOUND_ROWS          Return matched rows instead     2               The client cannot handle when MySQL returns the true value
'                                                                       of affected rows. If this flag is set, MySQL returns “found rows”
'                                                                       instead. You must have MySQL 3.21.14 or newer for this to work.
'   MULTI_STATEMENTS    Allow multiple statements       67108864            Enables support for batched statements. Added in 3.51.18.
'
'   OPTION = 67108867 = 67108864 + 3 (3 as per code examples)
'
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   From a web search (): http://www.continuumconcepts.com/Blog/CommentView,guid,7e26dec8-3a76-4819-83de-d15dd278198e.aspx
'   You might see OPTION=3 in your MySQL connection string. That number 3 in this case is the sum of a couple MySQL option flags.
'   In this case, it's FLAG_FIELD_LENGTH: "Do not Optimize Column Width", and FLAG_FOUND_ROWS: "Return Matching Rows
'   So, that option setting allows you to direct your MySQL server to behave in a specific manner for the duration of each connection.
'   A complete table of these options is available in the MySQL 5.0 Reference Manual.
'   You might also want to use OPTION=67108864, which allows you to execute multiple sql statements in a single MySQL Connector/ODBC batch,
'   separated by semicolons. To keep other things working the way most people expect, just use 67108867, which is all three options combined.
' --------------------------------------------------------------------------------------------------------------------------------------

Public Function GetCnnMySql(Optional ByVal pDriver As String, _
                            Optional ByVal pServer As String, _
                            Optional ByVal pDatabase As String, _
                            Optional ByVal pUID As String, _
                            Optional ByVal pPwd As String, _
                            Optional ByVal pOption As String, _
                            Optional ByRef pErrMsg As String) As ADODB.Connection
'                           Optional ByVal pCnnMode As ConnectModeEnum = adModeRead, _
'                           Optional ByRef pErrMsg As String) As ADODB.Connection
' -----------------------------------------------'
' pCnnMode has no effect {MySQL ODBC 5.1 Driver} '
' - so optional pCnnMode parameter commented out '
' pDatabase in MySQL terms is the Schema         '
' -----------------------------------------------'
Const kMySqlCnnString As String = "DRIVER={MySQL ODBC 5.1 Driver};" _
                                & "SERVER=localhost;" _
                                & "DATABASE=;" _
                                & "UID=;" _
                                & "PWD=;" _
                                & "OPTION=67108867" ' see module header notes for OPTION values

Dim strErrMsg As String
Dim strCnn As String
Dim cnn As ADODB.Connection

    strCnn = kMySqlCnnString
    
    If Len(pDriver) Then SetValStringValue pVString:=strCnn, pVName:="DRIVER", pValue:=pDriver
    If Len(pServer) Then SetValStringValue pVString:=strCnn, pVName:="SERVER", pValue:=pServer
    If Len(pDatabase) Then SetValStringValue pVString:=strCnn, pVName:="DATABASE", pValue:=pDatabase
    If Len(pUID) Then SetValStringValue pVString:=strCnn, pVName:="UID", pValue:=pUID
    If Len(pPwd) Then SetValStringValue pVString:=strCnn, pVName:="PWD", pValue:=pPwd
    If Len(pOption) Then SetValStringValue pVString:=strCnn, pVName:="OPTION", pValue:=pOption
    
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'   Cursor Location HARD CODED ?May later provide provide optional parameter?'
'   NOTE: BeginTrans, CommitTrans, and RollbackTrans methods are             '
'         not available on a client-side Connection object.                  '
'   pCnnMode has no effect {MySQL ODBC 5.1 Driver}                           '
    Set cnn = GetCnn(pDataSource:=strCnn, _
                     pCnnMode:=adModeUnknown, _
                     pCursorLocn:=adUseServer, _
                     pDataSourceType:=eCnnString, _
                     pErrMsg:=strErrMsg)
'                    pDataSourceType:=eCnnString, _
'                    pErrMsg:=strErrMsg)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

    pErrMsg = strErrMsg
    Set GetCnnMySql = cnn
    Set cnn = Nothing

End Function

Public Function GetCnnMySqlFromCnnString(ByVal pCnnString As String, _
                                Optional ByRef pErrMsg As String) As ADODB.Connection
Dim strDriver As String
Dim strServer As String
Dim strDatabase As String
Dim strUID As String
Dim strPwd As String
Dim strOption As String
Dim cnn As ADODB.Connection

    strDriver = GetValStringValue(pVString:=pCnnString, pVName:="DRIVER")
    strServer = GetValStringValue(pVString:=pCnnString, pVName:="SERVER")
    strDatabase = GetValStringValue(pVString:=pCnnString, pVName:="DATABASE")
    strUID = GetValStringValue(pVString:=pCnnString, pVName:="UID")
    strPwd = GetValStringValue(pVString:=pCnnString, pVName:="PWD")
    strOption = GetValStringValue(pVString:=pCnnString, pVName:="OPTION")

    Set cnn = GetCnnMySql(pDriver:=strDriver, _
                          pServer:=strServer, _
                          pDatabase:=strDatabase, _
                          pUID:=strUID, _
                          pPwd:=strPwd, _
                          pOption:=strOption, _
                          pErrMsg:=pErrMsg)
            
    Set GetCnnMySqlFromCnnString = cnn
    Set cnn = Nothing

End Function
