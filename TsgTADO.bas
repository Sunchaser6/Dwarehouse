Attribute VB_Name = "TsgTADO"
Option Explicit
' Requires references for:
'  *Microsoft Scripting Runtime
'  *Microsoft Jet and Replication Objects Library (msjro.dll)
'    Required for CompactMDB function (TsgDW only case where DBUtils used & CompactMDB is not b/c compacting by original DAO code)
'    Required DLL files (Msjro.dll, Msjetoledb40.dll - Jet OLE DB Provider) first available in MDAC 2.1.
'  *Microsoft ActiveX Data Objects Library
'    MDAC 2.5 distributed with Win2K & WinMe
'     [2.5 was last version to include Jet Drivers as part of distribution]
'     [a good version choice to distribute & include in references when using Jet]
'    To use later MDACs first install the MDAC and then the Jet Service Packs
'    MDAC 2.7 10/2001 distributed with Windows XP
'    MDAC 2.8 08/2003 distributed with Windows Server 2003
'    MDAC 2.8 SP1 08/2004 distributed with Windows XP SP2

'----------
' ADO Notes
'----------
' -* RecordCount *-
' RecordCount property returns -1 when ADO cannot determine the number of
' records or if the provider or cursor type does not support RecordCount.
' The cursor type of the Recordset object affects whether the number of records can be determined.
' The RecordCount property will return -1 for a forward-only cursor, the actual count for a static
' or keyset cursor, and either -1 or the actual count, depending on the data source, for a dynamic cursor.
' If the Recordset object supports approximate positioning or bookmarks [ie Supports (adApproxPosition)
' or Supports (adBookmark) return True] this value will be the exact number of records in the
' recordset regardless of whether it has been fully populated. ** If the recordset object does not support
' approximate positioning, this property may be a significant drain on resources because all records will
' have to be retrieved and counted to return an accurate RecordCount value. **

'   ADODB Properties
'   ----------------
' ADODB.Properties & ADODB.Property are NOT createable. Access is provided by the ADODB Provider
' ADO objects have two types of properties: built-in and dynamic.
' Built-in properties are those properties implemented in ADO and immediately
' available to any new object, using the MyObject.Property syntax.
' They do not appear as Property objects in an object's Properties collection,
' so although you can change their values, you cannot modify their characteristics.
' Dynamic properties are defined by the underlying data provider, and appear in the Properties collection for the appropriate ADO object.
' For example, a property specific to the provider may indicate if a Recordset object supports transactions or updating.
' These additional properties will appear as Property objects in that Recordset object's Properties collection.
' Dynamic properties can be referenced only through the collection, using the MyObject.Properties(0) or MyObject.Properties("Name") syntax.
'
' You cannot delete either kind of property.
'
' A dynamic Property object has four built-in properties of its own:
'
' The Name property is a string that identifies the property.
' The Type property is an integer that specifies the property data type.
' The Value property is a variant that contains the property setting. Value is the default property for a Property object.
' The Attributes property is a long value that indicates characteristics of the property specific to the provider.

'   CONNECTION OBJECT PROPERTIES (Notes are from a variety of sources including "ADO Provider Properties and Settings"    )
'   ---------------------------- (which provides a list of ADO Provider Proerties-both standard and jet provider specific.)
'                                 For a complete list of properties use code like the following in the immediate window
'                                "For i = 0 To cn.Properties.Count -1 : debug.Print cn(i).Name: next i"
'
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   STANDARD ADO Connection Object Initialization Properties
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~##############~~~~~~~~~~~
' Cache Authorization
' Encrypt Password
' Mask Password
' Password
' Persist Encrypted
' Persist Security Info
' User ID
' Asynchronous Processing
' Data Source
' Window Handle
' Locale Identifier
' Mode [ An MDB can only be open in one mode at a time. The first user to open ]
'      [ the MDB determines the locking mode to be used while the MDB is open. ]
'   adModeRead            (1)-Indicates read-only permissions.
'   adModeReadWrite       (3)-Indicates read/write permissions.
'   adModeRecursive(&H400000)-Used in conjunction with the other *ShareDeny* values to propagate sharing restrictions to all sub-records
'                             of the current Record. It has no affect if the Record does not have any children. ........................
'   adModeShareDenyNone  (16)-Allows others to open a connection with any permissions. Neither read nor write access can be denied to others.
'   adModeShareDenyRead   (4)-Prevents others from opening a connection with read permissions.
'   adModeShareDenyWrite  (8)-Prevents others from opening a connection with write permissions.
'   adModeShareExclusive (12)-Prevents others from opening a connection.
'   adModeUnknown         (0)-Default. Indicates that the permissions have not yet been set or cannot be determined.
'   adModeWrite           (2)-Indicates write-only permissions.
' Prompt
' Extended Properties
'   A String value (read/write) that specifies a string that contains provider specific connection
'   information that can’t be explicitly described through standard ADO properties.
'   For the Microsoft Jet provider, this property is used to pass the Microsoft Jet
'   connection string for opening or creating databases of other file formats.

'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ (cannot be set via the properties collection or ConnectionString)
'   NON- STANDARD ADO Connection Object Properties (must be set via their own properties on the connection objet   )
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   ConnectionTimeout & CursorLocation are supported by Jet Access but are not required to be supported by all providers
' ConnectionTimeout
'   Sets or returns a Long value that indicates, in seconds, how long to wait while establishing
'   a connection before terminating the attempt and generating an error. Default is 15.
' CursorLocation

'   ************
'   JET PROVIDER
'   ************
'   The "Microsoft.Jet.OLEDB.4.0" provider will open any Mdb file type & is the last Jet OLEDB
'   Provider version and the only one to suport compacting. (previously available through DAO)
'   Compacting is available through the Microsoft Jet and Replication Objects Library (msjro.dll)
'   DLL files required for compacting (Msjro.dll, Msjetoledb40.dll - Jet OLE DB Provider) first available in MDAC 2.1.
'   Refer to "ADO Provider Properties and Settings" for list of ADO Provider Proerties (both standard and jet provider specific).
'   The OLE DB Provider for Microsoft Jet supports several provider-specific connection parameters in addition to those defined by ADO.
'   The properties collection of a connection object contains these provider specific connection properties as well as Standard ADO
'   Connection Object Properties (initialization, information and sesssion) and provider specific properties. Any of these properties
'   (& the provider property) may be set in the ConnectionString.
'   ConnectionString may be provided via the property on the connection or as a parameter of the open method of a connection.
'   Neither ConnectionTimeout or CursorLocation are Standard ADO Connection Object Properties.
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   ALL Jet Provider Specific ADO Connection Object Initialization Properties
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~-------##############~~~~~~~~~~~
'   As with all other connection parameters, they can be set via the Connection
'   object's Properties collection or as part of the connection string.
' Jet OLEDB:Database Password
'   The database password.
' Jet OLEDB:Registry Path
'   The Windows registry key that contains values for the Microsoft Jet database engine.
' Jet OLEDB:System Database
'   The path and file name for the workgroup information file. (eg TsgMsgCenter.mdw)
' Jet OLEDB:Engine Type
'   see 'Public Enum MdbVersionEnum' below
' Jet OLEDB:Database Locking Mode
'   A Long value (read/write) that specifies the mode used when locking the database to read or modify records.
'   Property applies only to ADO connections made with the Jet OLE DB provider
'   The Jet OLEDB:Database Locking Mode property can be set to any of the following values:
'   Page-level Locking 0, Row-level Locking 1
' Jet OLEDB:Global Partial Bulk Ops
' Jet OLEDB:Global Bulk Transactions
' Jet OLEDB:New Database Password
' Jet OLEDB:Create System Database
' Jet OLEDB:Encrypt Database
' Jet OLEDB:Don’t Copy Locale on Compact
' Jet OLEDB:Compact Without Replica Repair
' Jet OLEDB: SFP
' Jet OLEDB:Compact Reclaimed Space Amount

'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   ALL Jet Provider Specific ADO Connection Object Session Properties
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~-------#######~~~~~~~~~~~
'   If you set a session property using the Properties collection,
'   you must set the property AFTER the connection has been opened.
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Jet OLEDB:Recycle Long-Valued Pages
' Jet OLEDB: Page Timeout
' Jet OLEDB:Shared Async Delay
' Jet OLEDB:Exclusive Async Delay
' Jet OLEDB: Lock Retry
' Jet OLEDB:User Commit Sync
' Jet OLEDB:Max Buffer Size
' Jet OLEDB: Lock Delay
' Jet OLEDB:Flush Transaction Timeout
' Jet OLEDB:Implicit Commit Sync
' Jet OLEDB:Max Locks Per File
' Jet OLEDB:ODBC Command Timeout
' Jet OLEDB:Reset ISAM Stats
' Jet OLEDB: Connection Control

'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   Some Jet Provider Specific ADO Command and Recordset Object properties
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~---------------------~~~~~~~~~~~~~~~~~~~~~~~
' Jet OLEDB:Locking Granularity
'   A Long value (read/write) that determines if a table is opened by using row-level (per-record) locking or page-level
'   locking. This property is ignored unless the Jet OLEDB:Database Locking Mode property is set to 1 (row-level locking).
'   Page-level locking 1, Row-level locking 2

' "Microsoft.Jet.OLEDB.4.0" driver
'   - Opens any Mdb file type
'   - Is the last Jet OLEDB Provider version
'   - Is the only Jet OLEDB version to support compacting (previously available through DAO)
Private Const mkDBJetProvider As String = "Microsoft.Jet.OLEDB.4.0"

Public Enum RstTypeEnum
    eReadOnlyFwdOnly
    eReadOnlyStatic
    eReadOnlyDynamic
    eEditableFwdOnly
'   eEditableKeyset
    eEditableDynamic
' * ADO Equivalents to DAO Recordset Types * '
'   eDaoDynaset
'   eDaoSnapshot
'   eDaoSnapshotForwardOnly
'   eDaoTable
End Enum

Public Enum DataSourceTypeEnum
    eMdb
    eDSN
    eCnnString
End Enum

Public Enum DbTypeEnum
    eJetDb
    eMySqlDb
End Enum

Public Enum SqlTypeEnum
    eJetSql
    eMySql
End Enum

Public Enum TxEnum
    eBeginTx
    eCommitTx
    eRollbackTx
End Enum

Public Function CBoolDb(ByVal pValue As Boolean, ByVal pDbType As DbTypeEnum) As Byte
'   pValue parameter is a boolean type (hence no processing of Value when using JetSQL)
'   May later define pValue parameter as a variant and convert all variant values, but for
'   now parameter definition will coerce arguments or fall over if wrong arguments are provided
    Select Case pDbType
        Case DbTypeEnum.eJetDb
            CBoolDb = pValue
        Case DbTypeEnum.eMySqlDb
            CBoolDb = CBoolMySql(pValue)
        Case Else
            Err.Raise Number:=1, Source:="DbBool()", Description:="Invalid pSqlType Parameter"
    End Select
End Function

Public Function CBoolMySql(ByVal pValue As Boolean) As Byte
'   Possibly better renamed as CBoolMySql() - then would match other conversion fns like VBA.CBool
' Also DbBool probably better renamed as CBoolDbVal??????
' Also other conversion functions like JetSqlDate & MySqlDate shoudld probably be renamed MsSqlDateStr & MySqlDateStr

'   Parameter data type will coerce arguments or fall over if wrong arguments are provided
    CBoolMySql = Abs(pValue)
End Function

Private Function CnvTxEnumToStr(pTxEnum As TxEnum) As String
Dim strResult As String

    Select Case pTxEnum
        Case TxEnum.eBeginTx
            strResult = "eBeginTx"
        Case TxEnum.eCommitTx
            strResult = "eCommitTx"
        Case TxEnum.eRollbackTx
            strResult = "eRollbackTx"
    End Select
    
    CnvTxEnumToStr = strResult

End Function

Public Function GetCnn(ByVal pDataSource As String, _
                       ByVal pCnnMode As ADODB.ConnectModeEnum, _
              Optional ByVal pDataSourceType As DataSourceTypeEnum = eMdb, _
              Optional ByVal pCursorLocn As ADODB.CursorLocationEnum = adUseClient, _
              Optional ByVal pCnnTimeout As Long = 15, _
              Optional ByVal pPropertyNamesArray As Variant, _
              Optional ByVal pPropertyValsArray As Variant, _
              Optional ByRef pErrMsg As String) As ADODB.Connection
'   See Module Header for notes on Connection Object Properties

'   Had considered passing properties as a collection rather than arrays of matching name and value pairs,
'   but collection items would require at least a name and value member and therefore some additional coding
'   May later create a custom collection but not currently worth the effort.
'   (ADODB.Properties & ADODB.Property are NOT createable. Access is provided by the ADODB Provider)
'   See Module Header notes for details of ADODB properties
Dim lngPrpLoop As Long
Dim strErrMsg As String
Dim fso As Scripting.FileSystemObject
Dim fil As Scripting.File
Dim cnn As ADODB.Connection

   If IsIDE() Then
   '   Allows others to open a connection with any permissions. Neither read nor write access can be denied to others.
   '   When running code in IDE we may want to inspect/manipulate data with another application (eg MS Access)
       pCnnMode = adModeShareDenyNone
   End If

'   Instantiate connection object & set properties
'   Properties available via properties collection can be set via ConnectionString & include the Provider property,
'   Standard ADO Connection Object Properties & Provider specific properties (see "ADO Provider Properties and Settings")
    Set cnn = New ADODB.Connection
    With cnn
    
        .CursorLocation = pCursorLocn
        .ConnectionTimeout = pCnnTimeout
        .Mode = pCnnMode ' eg. adModeRead, adModeShareExclusive, ...
        
        Select Case pDataSourceType
        
            Case DataSourceTypeEnum.eDSN
            '   Microsoft OLE DB Provider for ODBC is the default provider for ADO
                .ConnectionString = "DSN=" & pDataSource
                
            Case DataSourceTypeEnum.eMdb
            '   An MDB can only be open in one mode at a time.
            '   1st user to open MDB determines locking mode used while MDB is open.
            '   PROVIDER PROPERTY MUST BE SET PRIOR TO SETTING PROVIDER SPECIFIC PROPERTIES TO PROVIDE ACCESS
                .Provider = mkDBJetProvider
                .ConnectionString = "Data Source=" & pDataSource
            '   Clear read-only mdb file attribute when appropriate
                If pCnnMode <> ADODB.ConnectModeEnum.adModeRead Then
                    Set fso = New Scripting.FileSystemObject
                    If fso.FileExists(pDataSource) Then
                        Set fil = fso.GetFile(pDataSource)
                        If (fil.Attributes And ReadOnly) Then
                        '   Nb. File attributes can be changed while a cn is open, but the
                        '   cn must be closed and re-opened for it to be used as read-write
                            fil.Attributes = fil.Attributes - ReadOnly
                        End If
                    End If
                    Set fso = Nothing
                End If
                
            Case DataSourceTypeEnum.eCnnString
            '   Using this for GetCnnMySQL BUT MAY AT SOME STAGE COPY THIS
            '   PROCDURE AND ALTER IT TO BE GetCnnMySQL () THAT CALLS OpenCnn
            '   DIRECTLY. FOR THE TIME BEING WE WILL SEE IF THERE ARE
            '   COMMONALITIES AND COMMON BENEFITS FROM USING THIS PROCEDURE.

                .ConnectionString = pDataSource
'
        '   Case DataSourceTypeEnum.eSqlServer
        '       .DefaultDatabase =  ' Used when provider allows multiple dbs per connection (eg Sql Server)
        
        End Select
                
        If IsArray(pPropertyNamesArray) Then
        '   If pPropertyNamesArray is an array assume matching items in pPropertyValsArray
        '    - incorrect parameter passing should turn up early in testing
            For lngPrpLoop = LBound(pPropertyNamesArray, 1) To UBound(pPropertyNamesArray, 1)
                .Properties(pPropertyNamesArray(lngPrpLoop)) = pPropertyValsArray(lngPrpLoop)
            Next lngPrpLoop
        End If
        
    End With
    
    OpenCnn pCnn:=cnn, pErrMsg:=strErrMsg, pDelBracketedErrPrefix:=True
        
    Set GetCnn = cnn
    Set cnn = Nothing
    pErrMsg = strErrMsg
    
End Function

Public Function GetCollectionFromRst(ByRef pRst As ADODB.Recordset, _
                                     ByVal pFldName As String, _
                            Optional ByVal pForceMoveFirst As Boolean = True) As VBA.Collection
Dim colResult As VBA.Collection

    Set colResult = New VBA.Collection
    If Not (pRst Is Nothing) Then
        If Not (pRst.BOF And pRst.EOF) Then
            If pForceMoveFirst Then
                pRst.MoveFirst
            End If
            Do While Not pRst.EOF
                colResult.Add Item:=pRst(pFldName).Value, Key:=CStr(pRst(pFldName).Value)
                pRst.MoveNext
            Loop
        End If
    End If

    Set GetCollectionFromRst = colResult
    
End Function

Public Function GetRecordCount(ByRef pCnn As ADODB.Connection, ByVal pSource As String) As Long
' DOES NOT CATER FOR table names OR query names with a space in them becuase would then
' need conditional code for SqlType. Could do in same way as MsDate() MySqlDate() if
' want to head down generic path
Dim lngResult As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset

'   [MySQL][ODBC 5.1 Driver][mysqld-5.1.50-community]
'   Two issues
'   1 Every derived table must have its own alias
'           "SELECT COUNT(*) AS RecCount FROM (" & pSource & ")"
'        -> "SELECT COUNT(*) AS RecCount FROM (" & pSource & ") AS SubQuery"
'   2 Can't simply enclose a table/view name in brackets
'           "SELECT COUNT(*) AS RecCount FROM (Franchises)"
'        -> "SELECT COUNT(*) AS RecCount FROM Franchises"
    
    If IsSourceSQL(pSource) Then
        strSQL = "SELECT COUNT(*) AS RecCount FROM (" & pSource & ") AS SubQuery"
    Else
        strSQL = "SELECT COUNT(*) AS RecCount FROM " & pSource
    End If

    Set rst = GetRst(pCnn:=pCnn, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    lngResult = rst!RecCount
    rst.Close
    Set rst = Nothing
    
    GetRecordCount = lngResult

End Function

Public Function GetRst(ByRef pCnn As ADODB.Connection, _
                       ByVal pSource As String, _
                       ByVal pSourceType As CommandTypeEnum, _
              Optional ByVal pCursorLocn As ADODB.CursorLocationEnum = 0, _
              Optional ByVal pRstType As RstTypeEnum = eReadOnlyFwdOnly, _
              Optional ByRef pErrMsg As String) As ADODB.Recordset
' Default rst returned is Forward Only, Read Only
' (eDao prefixed RstTypeEnums added to assist translating DAO projects to ADO)
Dim strErrMsg As String
Dim rst As ADODB.Recordset
    
    Set rst = New ADODB.Recordset
    With rst
        .ActiveConnection = pCnn
        .Source = pSource
        
    '   Test CursorLocation passed/nominated (ie. pCursorLocn <> Default Param value)
    '   [ADODB.CursorLocationEnum can be adUseClient (ie 3) or adUseServer (ie 2)  ]
    '   If pCursorLocn passed set rst.CursorLocation otherwise inherit from connection
        If pCursorLocn <> 0 Then
            .CursorLocation = pCursorLocn
        End If
        
        Select Case pRstType
            Case eReadOnlyFwdOnly
                .LockType = adLockReadOnly
                .CursorType = adOpenForwardOnly
            
            Case eReadOnlyStatic
                .LockType = adLockReadOnly
                .CursorType = adOpenStatic
            
            Case eReadOnlyDynamic
                .LockType = adLockReadOnly
                .CursorType = adOpenDynamic
            
            Case eEditableFwdOnly
                .LockType = adLockOptimistic
                .CursorType = adOpenForwardOnly
            
            Case eEditableDynamic
                .LockType = adLockOptimistic
                .CursorType = adOpenDynamic

'           Case eEditableKeyset
'               .LockType = adLockOptimistic
'               .CursorType = adOpenKeyset
'
'           Case RstTypeEnum.eEditableFwdOnlyPessimistic
'               .LockType = adLockPessimistic
'               .CursorType = adOpenForwardOnly
'
'           Case RstTypeEnum.eEditableDynamicPessimistic
'               .LockType = adLockPessimistic
'               .CursorType = adOpenDynamic

'       '   ----------------------------------------
'       '*  ADO Equivalents to DAO Recordset Types *
'       '   ----------------------------------------
'       '   From http://msdn.microsoft.com/en-us/library/office/aa141422(v=office.10).aspx
'       '   The combined CursorType, LockType, and Options arguments of the ADO Open method
'       '   determine the type of ADO Recordset object that is returned. The table below
'       '   shows how the Type and Options arguments of the DAO OpenRecordset method can be
'       '   mapped to ADO Recordset Open method argument settings when you use ADO and the
'       '   Microsoft Jet 4.0 OLE DB Provider to work with Access databases.
'       '   (A snapshot derived from a Microsoft Jet-connected data source can’t be updated.)
'       '   ---------------------------------------------------------------------------------------
'       '   DAO arguments                 | ADO arguments
'       '   Type            Options       | CursorType          LockType           Options
'       '   ------------------------------|--------------------------------------------------------
'       '   dbOpenDynaset                 | adOpenKeyset        adLockOptimistic
'       '   dbOpenSnapshot                | adOpenStatic        adLockReadOnly
'       '   dbOpenForwardOnly             | adOpenForwardOnly   adLockReadOnly
'       '   dbOpenSnapshot  dbForwardOnly |  "   "   "   "       "   "   "   "
'       '   dbOpenTable                   | adOpenKeyset        adLockOptimistic   adCmdTableDirect
'       '   ---------------------------------------------------------------------------------------
'       '   dbOpenDynamic (DAO.RecordsetTypeEnum.dbOpenDynamic)
'       '                 Opens a dynaset-type Recordset (ODBCDirect workspaces only
'       '   ---------------------------------------------------------------------------------------
'       '   dbOpenDynaset (DAO.RecordsetTypeEnum.dbOpenDynaset)
'           Case RstTypeEnum.eDaoDynaset
'               .LockType = adLockOptimistic
'               .CursorType = adOpenKeyset
'       '   dbOpenSnapshot (DAO.RecordsetTypeEnum.dbOpenSnapshot)
'           Case RstTypeEnum.eDaoSnapshot
'               .LockType = adLockReadOnly
'               .CursorType = adOpenStatic
'       '   dbOpenForwardOnly (DAO.RecordsetTypeEnum.dbOpenForwardOnly)
'           Case RstTypeEnum.eDaoSnapshotForwardOnly
'               .LockType = adLockReadOnly
'               .CursorType = adOpenForwardOnly
'       '   dbOpenTable (DAO.RecordsetTypeEnum.dbOpenTable)
'           Case RstTypeEnum.eDaoTable
'               .LockType = adLockOptimistic
'               .CursorType = adOpenKeyset
        End Select

        #If cc_TsgDW = -1 Then
        '   Code executed only when procedure is used by TsgDw.exe
        '   [TsgDw.exe identified by conditional constant cc_TsgDw = -1 in TsgDw project file]
            If Not g.bMaster Then
            '   Ensure only Master instance of TsgDw has write access to TsgDw database
                .LockType = adLockReadOnly
            End If
        #End If
    
    End With
    
'~  Options can be one or more CommandTypeEnum or ExecuteOptionEnum values,
'~  which can be which can be combined with a bitwise AND operator.
'~  If you pass something other than a Command object in the Source argument, you can
'~  use the Options argument to optimize evaluation of the Source argument.
'~  If the Options argument is not defined, you may experience diminished performance because
'~  ADO must make calls to the provider to determine if the argument is an SQL statement,
'~  a stored procedure, a URL, or a table name. If you know what Source type you're using,
'~  setting the Options argument instructs ADO to jump directly to the relevant code.

'   If pRstType = RstTypeEnum.eDaoTable Then
'      '   Ensure correct Options for RstTypeEnum regardless of what pSourceType was passed
'          On Error Resume Next
'              rst.Open Options:=adCmdTableDirect
'              strErrMsg = VBA.Err.Description ' See GetCn_ADO() & OpenCn() FOR WAYS OF RETURNING MORE ERROR INFO (EG FROM ADO.Errors COLLECTION)
'                                              ' Could write a common procedure that returns AODDB.cnn errors
'          On Error GoTo 0
'   Else
        On Error Resume Next
            rst.Open Options:=pSourceType
            strErrMsg = VBA.Err.Description ' See GetCn_ADO() & OpenCn() FOR WAYS OF RETURNING MORE ERROR INFO (EG FROM ADO.Errors COLLECTION)
                                            ' Could write a common procedure that returns AODDB.cnn errors
        On Error GoTo 0
'   End If
    
    If Len(strErrMsg) Then
        Set rst = Nothing
    End If

    pErrMsg = strErrMsg
    Set GetRst = rst
    Set rst = Nothing

End Function

Public Function GetRstAddOnly(ByRef pCnn As ADODB.Connection, _
                              ByVal pSource As String, _
                     Optional ByRef pErrMsg As String) As ADODB.Recordset
                     
'   Adjust source to return no records before passing it on to GetRst()
'   Sql statement intentionally returns no records becuase when connected to MySQL database
'   and based on the table name with no filter it takes 11 mins to open the recordset

'   Sql statement intentionally returns no records becuase when connected to MySQL database
'   and based on the table name with no filter it takes 11 mins to open the recordset
''  strSQL = "SELECT * FROM EventLog WHERE True = False"

'   CODE CURRENTLY IGNORES ALL CASES OF SQL, BUT THE MOST SIMPLE
'   WITH NO CLAUSES SUBSEQUENT TO A WHERE CLAUSE
'   CAN LATER BE IMPROVED TO SPLIT INTO COMPOSITE CLAUSES IF REQUIRED
Dim strSource As String
    
    If Not IsSourceSQL(pSource) Then
    '   Assume pSource is a table, view, proc or query name
        strSource = "SELECT * FROM " & pSource & " WHERE False"
    Else
    '   pSource is an SQL string
        If InStr(pSource, "WHERE", vbTextCompare) Then
        '   Assume SQL has Where Clause as the last clause!!!!
            strSource = pSource & " AND False"
        Else
        '   Assume SQL has only a SELECT and FROM clause!!!!
            strSource = pSource & ") AS SubQuery"
        End If
    End If
                     
    Set GetRstAddOnly = GetRst(pCnn:=pCnn, _
                               pSource:=strSource, _
                               pSourceType:=adCmdText, _
                               pCursorLocn:=adUseClient, _
                               pRstType:=eEditableFwdOnly, _
                               pErrMsg:=pErrMsg)

End Function

Public Function GetRstVal(ByRef pCnn As ADODB.Connection, _
                          ByVal pSource As String, _
                 Optional ByVal pDefaultVal As Variant = Empty) As Variant
'''              Optional ByVal pDefaultVal As Variant = vbEmpty) As Variant
' ******** RETURN TO GIVE THIS FUNCTION BETTER HEADER COMMENTS ******** ''' Review
'
' ------------------------------------------------------
' For most SQL (ie SQL not using Min() or Max functions)
' ------------------------------------------------------
'   Returns: Field value of first returned row
'            If empty rst returned then returns pDefaultVal
'            (If pDefaultVal not supplied defaults to vbEmpty)
'
' *--------------------------------------*
' * For SQL USING Min() or Max functions *
' *--------------------------------------*
'   Returns: Field value of first returned row if rst field value <> Null
'            otherwise returns pDefaultVal
'            (If pDefaultVal not supplied defaults to vbEmpty)
'           Returns pDefaultVal when rst field value = Null because
'           Min() and Max() functions return Null when applied to empty rst
'
'   Queries using aggregate/column functions always return at least one row
'   When Min() and Max() are used but selection doesn't apply to any rows they return NULL

'RETURN TO GIVE THIS FUNCTION BETTER HEADER COMMENTS ''' Review

'   Reasoning behind using only one field (=> must test for one field)
'   Is GetVal not GetValS. So if you're after one field your SQL should only
'   return one field and we can assume it is the only field of the rst
'   and do a test to check there is only one field in the rst and return
'   an RAISE an error otherwise.
' Catering for Min() and Max() b/c it makes for more readable calling
' code without the caller having to remember the case where these
' fns return Null when it is an empty rst. Also, although you could return
' one or no rows with the Min or Max according to the SQL you passed
' instead of using Min and Max functions, the SQL would have to be
' specific to the DB Engine. e.g. MySQL would use Limit 1 at end of
' SQL and MS Sql would use TOP 1 as the at the start of the SQL.
' By using this code we needn't limit the calling code to a specific
' type of database or to complicated calling code.
'
'
' ******** RETURN TO GIVE THIS FUNCTION BETTER HEADER COMMENTS ******** ''' Review

Dim strTemp As String
Dim strErrMsg As String
Dim vntResult As Variant
Dim vntRstValue As Variant
Dim rst As ADODB.Recordset

    Set rst = GetRst(pCnn:=pCnn, _
                     pSource:=pSource, _
                     pSourceType:=adCmdUnknown, _
                     pRstType:=eReadOnlyFwdOnly, _
                     pErrMsg:=strErrMsg)
                     
    If rst.Fields.Count <> 1 Then
        rst.Close
        Set rst = Nothing
        Err.Raise Number:=1, _
                  Source:="GetRstVal()", _
                  Description:="Field count of recordset returned by pSource <> 1"
    Else
        If (rst.BOF And rst.EOF) Then
            vntResult = pDefaultVal
        Else
            vntRstValue = rst(0).Value
            strTemp = UCase$(Replace$(Expression:=pSource, Find:=" ", Replace:=""))
        '   Search for MIN( or MAX( to exclude Min or Max being part of a field name etc
' Works ''' If InStr(strTemp, "MIN(") + InStr(strTemp, "MAX(") = 0 Then
' FAILS ''' If InStr(strTemp, "MIN(", vbBinaryCompare) + InStr(strTemp, "MAX(", vbBinaryCompare) = 0 Then   ''' Review Doesn't work. Would like to find out why
            If VBA.InStrB(strTemp, "MIN(", Compare:=vbBinaryCompare) + VBA.InStrB(strTemp, "MAX(", Compare:=vbBinaryCompare) = 0 Then   ''' WORKS ONCE InStr[B] is qualified with VBA, and Compare argument is named
            '   SQL is NOT using Min() or Max() functions
                vntResult = vntRstValue
            Else
            '   SQL using Min() or Max() which return Null when applied to empty recordset
            '   [Empty rst: source tables may be empty or Where Clause may exlude all rows]
                If IsNull(vntRstValue) Then
                    vntResult = pDefaultVal
                Else
                    vntResult = vntRstValue
                End If
            End If
        End If
        
        rst.Close
        Set rst = Nothing
    
        GetRstVal = vntResult
    End If

End Function

Public Function IsFldExists(ByRef pRst As ADODB.Recordset, ByVal pFldName As String) As Boolean
Dim fld As ADODB.Field

    On Error Resume Next
        Set fld = pRst.Fields(pFldName)
        IsFldExists = (Err.Number = 0)
    On Error GoTo 0
    
    Set fld = Nothing
    
'Dim bResult As Boolean
'Dim fld As ADODB.Field
'
'    For Each fld In pRst.Fields
'        If UCase$(fld.Name) = UCase$(pFldName) Then
'            bResult = True
'            Exit For
'        End If
'    Next fld
'    Set fld = Nothing
'
'    IsFldExists = bResult
End Function

Public Function IsSourceSQL(ByVal pSource As String) As Boolean
Dim astrSourceWords() As String

'   bResult: Assume saved queries/views/procs DON'T have embedded spaces and therefore if
'            pSource has a space it is an SQL statement (eg "SELECT * FROM Franchises")

'   -----------------------------------------------------------------------------------------
'   SIMPLE APPROACH THAT WILL BE WRONG IF EITHER TABLE NAMES OR VEIW/PROC/QUERY NAMES HAVE  -
'   EMBEDDED SPACES. THIS LOGIC IS ISOLATED IN THIS PROCEDURE SO IT CAN LATER BE IMPROVED   -
'   TO CATER FOR ALL CASES                                                                  -
'   -----------------------------------------------------------------------------------------
    
    astrSourceWords = Split(pSource, " ")
    IsSourceSQL = (UBound(astrSourceWords, 1) > 0)

End Function

Public Function JetSqlDate(ByVal pDate As Date) As String
   JetSqlDate = Format$(pDate, "\#dd mmm yyyy#\")
End Function

Public Function JetSqlQIdentifier(ByVal pFldName As String) As String
    JetSqlQIdentifier = "[" & pFldName & "]"
End Function

Public Sub LoadCombo_Rst(ByRef pCombo As VB.ComboBox, _
                         ByRef pCnn As ADODB.Connection, _
                         ByVal pSource As String, _
                         ByVal pDisplayFld As String, _
                Optional ByVal pDataFld As String, _
                Optional ByVal pPreserveSel As Boolean = True)
' LoadCbo_Rst   - passed a rst
' LoadCombo_Rst - pass a cnn and source (table or query name, sql, ...)

'   Uses ItemData property to maintains current combo selection if pPreserveSel is True

'''   ComboBox click event and LoadCombo procedure work in concert
'''   ToDo Candidate for a generic function that is passed a combobox control, rst,
'''   ... noting that ComboBox click event and LoadCombo must work in concert together

Dim lngSavedListindex As Long
Dim lngSavedItemData As Long
Dim lngLoop As Long
Dim rst As ADODB.Recordset

    If pPreserveSel Then
        lngSavedListindex = pCombo.ListIndex
        If lngSavedListindex > -1 Then
            lngSavedItemData = pCombo.ItemData(lngSavedListindex)
        End If
    End If
    
    pCombo.Clear    ' Clear list items and selection
    Set rst = pCnn.Execute(CommandText:=pSource)
    With rst
        While Not .EOF
            pCombo.AddItem .Fields(pDisplayFld).Value
            If Len(pDataFld) Then
                pCombo.ItemData(pCombo.NewIndex) = .Fields(pDataFld).Value
            End If
            .MoveNext
        Wend
    End With
    
    If pPreserveSel Then
        If lngSavedListindex > -1 Then
            With pCombo
                For lngLoop = 0 To .ListCount - 1
                    If .ItemData(lngLoop) = lngSavedItemData Then
                        .ListIndex = lngLoop
                        Exit For
                    End If
                Next lngLoop
            End With
        End If
    End If
    
End Sub

Public Sub LoadListBox_Rst(ByRef pListBox As VB.ListBox, _
                           ByRef pCnn As ADODB.Connection, _
                           ByVal pSource As String, _
                           ByVal pDisplayFld As String, _
                  Optional ByVal pDataFld As String = vbNullString)
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    pListBox.Clear
    Set rst = GetRst(pCnn:=pCnn, pSource:=pSource, pSourceType:=adCmdUnknown, pErrMsg:=strErrMsg)
    If Not rst Is Nothing Then
        With rst
            Do While Not .EOF
                pListBox.AddItem .Fields(pDisplayFld).Value
                If Len(pDataFld) Then
                    pListBox.ItemData(pListBox.NewIndex) = .Fields(pDataFld).Value
                End If
                .MoveNext
            Loop
            .Close
        End With
        Set rst = Nothing
    End If
                
End Sub

Public Function MySqlDate(ByVal pDate As Date) As String
'   COULD MAKE MySqlDateTime test for time portion and only used copmlete
'   format ("yyyy-mm-dd Hh:Nn:Ss") if there is a time portion
'   MAYBE THE FULL PORTION WITH HH:NN:SS BEING 0 WOULD WORK FOR DATES?
'
'   Args NOT passed ByRef for speed, BECAUSE IT APPEARS YOU CAN'T COERCE A VARIANT OF VARIANT
'   DATA TYPE HOLDING A vbDate VALUE INTO A Date DATA TYPE
    MySqlDate = Format$(pDate, "\'yyyy-mm-dd'\")
End Function

Public Function MySqlDateTime(ByVal pDate As Date) As String
'   Args NOT passed ByRef for speed, BECAUSE IT APPEARS YOU CAN'T COERCE A VARIANT OF VARIANT
'   DATA TYPE HOLDING A vbDate VALUE INTO A Date DATA TYPE
    MySqlDateTime = Format$(pDate, "\'yyyy-mm-dd hh:nn:ss'\")
End Function

Public Function MySqlQ(ByVal pString As String) As String
Dim strResult As String
'   Works for ANSI SQL when there is a single quote in the literal string.
'   Tested for with Jet and ADO connections to BM and Access databases.
'   NB. Does NOT work for two adjacent single quotes in a literal string.
'       I don't know if/how the standard caters for this.

'   Note Access allows the use of double quotes in literal strings to get around
'   the problem, but ANSI SQL uses double quotes like Access uses square brackets
'   (ie. for database objects/identifiers eg. "WHERE [Field Name] = 'Fred'")

'   Replace$ is slow and is therefore only called when needed
'   InStrB is quicker than InStr and can be used when only checking for existence
'   of characters inside the range of 1-255 (presume that is Non Wide characters Kanji etc)

'   From Help: The InStrB function is used with byte data contained in a string.
'   Instead of returning the character position of the first occurrence of one
'   string within another, InStrB returns the byte position.
    
    
'   Might be a good idea to create a MySqlConfig object that I can pass around that
'   contains all the MySql settings such as NO_BACKSLASH_ESCAPES SQL
    
'
'  LITERAL STRINGS (From MySQL help)
'  ---------------------------------
'  Within a string, certain sequences have special meaning unless the NO_BACKSLASH_ESCAPES SQL mode
'  is enabled.
'
'   Each of these sequences begins with a backslash (“\”), known as the escape character. MySQL recognizes
'  the escape sequences shown in Table 9.1, “Special Character Escape Sequences”. For all other escape
'  sequences, backslash is ignored. That is, the escaped character is interpreted as if it was not escaped.
'  For example, “\x” is just “x”. These sequences are case sensitive. For example, “\b” is interpreted as
'  a backspace, but “\B” is interpreted as “B”. Escape processing is done according to the character set
'  indicated by the character_set_connection system variable. This is true even for strings that are
'  preceded by an introducer that indicates a different character set, as discussed in Section 10.1.3.5,
'  “Character String Literal Character Set and Collation”.
'
'  Table 9.1 Special Character Escape Sequences
'
'  Escape Sequence Character Represented by Sequence
'  \0  An ASCII NUL (X'00') character
'  \'  A single quote (“'”) character
'  \"  A double quote (“"”) character
'  \b  A backspace character
'  \n  A newline (linefeed) character
'  \r  A carriage return character
'  \t  A tab character
'  \Z  ASCII 26 (Control+Z); see note following the table
'  \\  A backslash (“\”) character
'  \%  A “%” character; see note following the table
'  \_  A “_” character; see note following the table
'
'  The ASCII 26 character can be encoded as “\Z” to enable you to work around the problem that ASCII 26
'  stands for END-OF-FILE on Windows. ASCII 26 within a file causes problems if you try to use mysql
'  db_name < file_name.
'
'  The “\%” and “\_” sequences are used to search for literal instances of “%” and “_” in pattern-matching
'  contexts where they would otherwise be interpreted as wildcard characters. See the description of the
'  LIKE operator in Section 12.5.1, “String Comparison Functions”. If you use “\%” or “\_” outside of
'  pattern-matching contexts, they evaluate to the strings “\%” and “\_”, not to “%” and “_”.
'
'  If the ANSI_QUOTES SQL mode is enabled, string literals can be quoted only within single quotation marks
'  because a string quoted within double quotation marks is interpreted as an identifier.
    
    strResult = pString
    
    If InStr(strResult, "'") Then
        strResult = Replace$(strResult, "'", "''")
    End If
    
    If InStr(strResult, "\") Then
        strResult = Replace$(strResult, "\", "\\")
    End If
    
    MySqlQ = "'" & strResult & "'"

End Function

Public Function MySqlQIdentifier(ByVal pFldName As String) As String
    MySqlQIdentifier = "`" & pFldName & "`"
End Function

Private Sub OpenCnn(ByRef pCnn As ADODB.Connection, ByRef pErrMsg As String, ByVal pDelBracketedErrPrefix As Boolean)
Dim strVBAErr As String
Dim strErrMsg As String

    On Error Resume Next
        pCnn.Open
        strVBAErr = VBA.Err.Description
    On Error GoTo 0

    If Len(strVBAErr) <> 0 Then
        If pDelBracketedErrPrefix Then
            strErrMsg = TrimWhiteSpace(StripBracketedPrefixes(strVBAErr))
        Else
            strErrMsg = strVBAErr
        End If
    End If

    If Not pCnn Is Nothing Then          '* Not sure if this IF Statement is needed
        If pCnn.Errors.Count <> 0 Then
        '   Currently only check and report the most recently added error {ie cn.Errors(0)} (27Sep2005)
            If strVBAErr <> pCnn.Errors(0).Description Then
            '   Disregard non-fatal, informational connection messages {ie where err number = 0}
            '   {eg [Microsoft][ODBC Driver Manager] Driver's SQLSetConnectAttr failed}
            '   Happens with Probe DSNs and I can't isolate which connection property it can't set/(is complaining about)
                If pCnn.Errors(0).Number <> 0 Then
                '   Somtimes VBA and connection errors are identical
                    If pDelBracketedErrPrefix Then
                        strErrMsg = strErrMsg & vbNewLine & TrimWhiteSpace(StripBracketedPrefixes(pCnn.Errors(0).Description))
                    Else
                        strErrMsg = strErrMsg & vbNewLine & pCnn.Errors(0).Description
                    End If
                End If
            End If
        End If
        If Len(strErrMsg) Then
            If pCnn.State = ADODB.adStateOpen Then
                pCnn.Close   ' Closing Cn b/c there were errors
            End If
            Set pCnn = Nothing
        End If
    End If

    pErrMsg = strErrMsg

End Sub

Public Function SqlDate(ByVal pDate As Date, ByVal pSqlType As SqlTypeEnum) As String
    Select Case pSqlType
        Case eJetSql
            SqlDate = JetSqlDate(pDate)
        Case eMySql
            SqlDate = MySqlDate(pDate)
        Case Else
            Err.Raise Number:=1, Source:="SqlDate()", Description:="Invalid pSqlType Parameter"
    End Select
End Function

Public Function SqlQ(ByVal pString As String) As String
' Same as SqlQuote() but shorter name (will migrate to this function)
' May be renamed to SqlQString()
'
' Function could also be modified to take a SqlType parameter
' to cater for various flavours of SQL (e.g. MySql with its backslash escape character sqequences etc.)
'
'   Works for ANSI SQL when there is a single quote in the literal string.
'   Tested for with Jet and ADO connections to BM and Access databases.
'   NB. Does NOT work for two adjacent single quotes in a literal string.
'       I don't know if/how the standard caters for this.

'   Note Access allows the use of double quotes in literal strings to get around
'   the problem, but ANSI SQL uses double quotes like Access uses square brackets
'   (ie. for database objects/identifiers eg. "WHERE [Field Name] = 'Fred'")

'   Replace$ is slow and is therefore only called when needed
'   InStrB is quicker than InStr and can be used when only checking for existence
'   of characters inside the range of 1-255 (presume that is Non Wide characters Kanji etc)

'   From Help: The InStrB function is used with byte data contained in a string.
'   Instead of returning the character position of the first occurrence of one
'   string within another, InStrB returns the byte position.
    
   If InStr(pString, "'") <> 0 Then
       SqlQ = "'" & Replace$(pString, "'", "''") & "'"
   Else
       SqlQ = "'" & pString & "'"
   End If

End Function

Public Function SqlQIdentifier(ByVal pFldName As String, ByVal pSqlType As SqlTypeEnum) As String
'   For quoting identifiers (cf quoting literal strings) such as table names, query names, ...
'   Alternate names: QSqlIdentifier, SqlQIdentifier

    Select Case pSqlType
        Case eJetSql
            SqlQIdentifier = JetSqlQIdentifier(pFldName)
        Case eMySql
            SqlQIdentifier = MySqlQIdentifier(pFldName)
        Case Else
            Err.Raise Number:=1, Source:="SqlQIdentifier()", Description:="Invalid pSqlType Parameter"
    End Select

End Function

Public Function SqlQuote(ByVal pString As String) As String
'   Works for ANSI SQL when there is a single quote in the literal string.
'   Tested for with Jet and ADO connections to BM and Access databases.
'   NB. Does NOT work for two adjacent single quotes in a literal string.
'       I don't know if/how the standard caters for this.

'   Note Access allows the use of double quotes in literal strings to get around
'   the problem, but ANSI SQL uses double quotes like Access uses square brackets
'   (ie. for database objects/identifiers eg. "WHERE [Field Name] = 'Fred'")

'   Replace$ is slow and is therefore only called when needed
'   InStrB is quicker than InStr and can be used when only checking for existence
'   of characters inside the range of 1-255 (presume that is Non Wide characters Kanji etc)

'   From Help: The InStrB function is used with byte data contained in a string.
'   Instead of returning the character position of the first occurrence of one
'   string within another, InStrB returns the byte position.
    
   If InStr(pString, "'") <> 0 Then
       SqlQuote = "'" & Replace$(pString, "'", "''") & "'"
   Else
       SqlQuote = "'" & pString & "'"
   End If

End Function

Public Sub Tx(ByRef pCnn As ADODB.Connection, ByVal pProcName As String, ByVal pTxAction As TxEnum)
' BeginTrans, CommitTrans, and RollbackTrans methods are not available on a client-side Connection object.
' (could put a test in the proc?) The test would need to give a WARNING code will not be functioning
'   as intended because the connection is ClientSide and Transaction functions are not supported for
'  ClientSide connections. The code won't break but the protection and or speed improvements from
'  using this transaction code will not be realised.
'
' Depending on the Connection object's Attributes property, calling either the CommitTrans or RollbackTrans
' methods may automatically start a new transaction. If the Attributes property is set to adXactCommitRetaining,
' the provider automatically starts a new transaction after a CommitTrans call. If the Attributes property is
' set to adXactAbortRetaining, the provider automatically starts a new transaction after a RollbackTrans call.

'   Should eventually keep a collection of cnns so I can act on each cnn (push/pop/trace/commit) independently
'   For the moment I can simply check that the Txn I am passed is the same as the one I have stored and if
'   I don't have one stored (e.g. I have committed, then I can create a reference to the new one.
'   When I commit on a Txn I will clear the reference I have to the connection
'   WITH MySQL I WON'T HAVE ANY NESTING OF TXNS (IE NO ISOLATION LEVELS)
'   MySQL DOES NOT SUPPORT THIS ON ANY LEVEL (EVEN NATIVELY)
'   (I could always return the isolation level as the result of the procedure. 0 to n with -1 being a failure?)

'   Good for this or related function to be queriable for most recent ProcCall
'    in that way it could be used in-line to determine how and if to RollBack etc.
'   All the fnality of this proc could be put in an object and that could be given the method
'   the object could delegate to a relevant cnns etc. Alternately there would be other
'   ways of getting around it like creating a Tx module

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' all this Tx stuff should probably be an object that you instantiate for a particular cnn
' this would be self documenting enough (eg g.cnnDwTx.Commit, g.cnnDwTx.BeginTx, g.cnnDwTx.Rollback
' could even go as far as creating a TCnn object that delegates mostly to a Cnn object
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim strErrMsg As String
Dim strErrSource As String
Dim strProcCall As String
Static scolTxStack As VBA.Collection
Static scnnTx As ADODB.Connection

    strProcCall = pProcName & "-> Tx(" & CnvTxEnumToStr(pTxAction) & ")"
'   Debug.Print strProcCall

    If pCnn Is Nothing Then
        strErrMsg = "Invalid Tx(" & CnvTxEnumToStr(pTxAction) & ") call from " & pProcName
    Else
    
        If scolTxStack Is Nothing Then              '
            Set scolTxStack = New VBA.Collection    '
        End If                                      '
    
        Select Case pTxAction
            Case TxEnum.eBeginTx
                If scnnTx Is Nothing Then
                    Set scnnTx = pCnn
                '   If scolTxStack Is Nothing Then
                '       Set scolTxStack = New VBA.Collection
                '   End If
                ElseIf scnnTx <> pCnn Then
                    strErrMsg = "Tx in process for other connection"
                End If
                If Len(strErrMsg) = 0 Then
                    scolTxStack.Add Item:=pProcName ' Push pProcName on to LIFO stack
                    If scolTxStack.Count = 1 Then
                        pCnn.Errors.Clear
                        pCnn.BeginTrans ' scnnTx.BeginTrans
                    End If
                End If
            
            Case TxEnum.eCommitTx, TxEnum.eRollbackTx
                If scnnTx Is Nothing Then
                    strErrMsg = "No Tx in process"
                ElseIf scnnTx <> pCnn Then
                    strErrMsg = "Tx in process for other connection"
                End If
                If Len(strErrMsg) = 0 Then
                    If scolTxStack Is Nothing Then
                        strErrMsg = "No matching TxBegin for " & strProcCall
                    ElseIf scolTxStack.Count = 0 Then
                        strErrMsg = "No matching TxBegin for " & strProcCall
                    ElseIf pProcName <> scolTxStack.Item(scolTxStack.Count) Then
                        strErrMsg = CnvTxEnumToStr(eBeginTx) & " without matching TxBegin - " & pProcName & "()"
                    ElseIf scolTxStack.Count = 1 Then
                            Select Case pTxAction
                                Case TxEnum.eCommitTx:  pCnn.CommitTrans    ' scnnTx.CommitTrans
                                Case TxEnum.eRollbackTx: pCnn.RollbackTrans ' scnnTx.eRollbackTx
                                Case Else
                                    strErrMsg = "Problem with procedure logic in procedure: Tx()"
                            End Select
                    End If
                    If Len(strErrMsg) = 0 Then
                        scolTxStack.Remove Index:=scolTxStack.Count ' Pop pProcName from LIFO stack
                        If scolTxStack.Count = 0 Then
                            Set scnnTx = Nothing
                        End If
                    End If
                End If
        End Select
    End If

    If Len(strErrMsg) Then
        strErrSource = "Tx(" & CnvTxEnumToStr(eBeginTx) & ")"
        Err.Raise Number:=1, Source:=strErrSource, Description:=strErrSource & ": " & strErrMsg
    End If

End Sub







