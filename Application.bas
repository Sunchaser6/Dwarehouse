Attribute VB_Name = "Application"
Option Explicit

' GLOBAL CONSTANTS
' ----------------
Public Const gkPromoADD As String = "PROMO"
Public Const gkPromoDELETE As String = "DELPROMO"
Public Const gkOPosFranType As Long = 100

' GLOBAL ENUMS
' ------------
' tblFranchisePromotion!TfrStatus Enum (FP!TfrStatus)
Public Enum FpTfrEnum
'   Use bit flag values to provide option of later storing histroy of states in a single field.
'   (Eg. A value of 4 would mean it promo has been precalled without ever having been uploaded.)
'   As at V386 tblFranchisePromoition!FpTfrStatus holds only the current state
    FpTfrRequested = 0          ' Default value on creation
    FpTfrCompleted = 1          ' Promotion succesfully transferred to Franchise
    FpRecallRequested = 2       ' Promotion Recall Requested via FP table in TsgDw db
    FpRecallRequestUploaded = 4 ' Fran RStats.mdb modified for promo recall b/c promo has been processed at Fran
                                ' (DOESN'T GUARANTEE RStats.exe IS RUNNING AT FRAN & HAS TAKEN APPROPRIATE ACTION)
                                ' (OF MODIFYING RMgr.mdb, OR WITH 3rd PARTY SYSTEMS NOTIFIED FRAN OPERATORS      )
                                ' (TsgMsgCentre CAN SHOW WHETHER RStats.exe HAS TAKEN APPROPRIATE ACTION         )
    FpRecalled = 8              ' Promo record deleted from RStats.mdb b/c promo has not been processed at Fran
                                ' (ie. was not applied to RMgr.mdb, or for 3rd party POS not notified Fran operator)
  ' fpuN1 = 16 etc
  ' fpuN2 = 32 etc
  ' fpuN3 = 64 etc
End Enum

Public Function GetAppDefaultsRst(ByRef pCnnAppDefaults As ADODB.Connection, _
                         Optional ByRef pErrMsg As String) As ADODB.Recordset
' TsgDw executes specific code in GetRst() to limit instances other than the Master to
' read-only recordsets. AppDefaultsRst (Defaults table in AppPath\Defaults.mdb) needs
' to be writeable for all instances of TsgDw so this procedure is written and used as
' the only exeption to using GetRst() in TsgDw to using GetRst to create recordsets.
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    Set rst = New ADODB.Recordset
    With rst
    '   Inherit CursorLocn from cnn (adUseClient, adUseServer)
        .ActiveConnection = pCnnAppDefaults
        .Source = "Defaults"
        .CursorType = adOpenForwardOnly
        .LockType = adLockOptimistic
    End With

    On Error Resume Next
        rst.Open Options:=adCmdTable
        strErrMsg = VBA.Err.Description
    On Error GoTo 0
    
    If Len(strErrMsg) Then
        Set rst = Nothing
    End If

    pErrMsg = strErrMsg
    Set GetAppDefaultsRst = rst
    Set rst = Nothing

End Function

Public Sub Main()
' Nb. When program is run the first lines executed are loading the custom custrol
'     (TDatePicker-displayed on main form. Even before this procedure starts.
Dim intPrevMousePointer As Integer
Dim strMsg As String
Dim strErrMsg As String
Dim strCnnString As String
Dim strPcName As String
Dim strDefaultsMdbFullname As String
Dim frmGetCnn As fdlgGetCnnMySQL
' To prevent two program instances under different logins on the one machine g.rstAppDefaults
' is kept open for the life of program and uses exclusively opened cnnAppDefaults as its cnn

Dim cnnAppDefaults As ADODB.Connection

    strDefaultsMdbFullname = App.Path & "\Defaults.mdb"
    
    If App.PrevInstance Then
        strMsg = App.Title & " is already running"
        MsgBox strMsg, vbInformation
        
    ElseIf Not IsDateFmtOk() Then   ''' Remove reliance on date format when time permits
        strMsg = "This program relies on a particular date setting." & vbNewLine & _
                 "Use regional option in control panelto ensure date format is set to " & SQ("dd/MM/yy")
        MsgBox strMsg, vbInformation
        
    ElseIf Len(Dir$(strDefaultsMdbFullname)) = 0 Then
        strMsg = "Defaults database not found" & vbNewLine & _
                 strDefaultsMdbFullname & vbNewLine & vbNewLine & _
                 vbNewLine & _
                 gconContactSystemAdministrator
        MsgBox strMsg, vbCritical, gconUnknownProductName
        
    Else
    '   Opening connection exclusively prevents two instances under different logins on same machine
    '   ([ effectively] as there would never be two installations on same machine)
    '   g.rstAppDefaults is open for lifetime of program, so it keeps the associated connection open
        Set cnnAppDefaults = GetCnn(pDataSource:=strDefaultsMdbFullname, _
                                    pCnnMode:=adModeShareExclusive, _
                                    pDataSourceType:=eMdb, _
                                    pCursorLocn:=adUseClient, _
                                    pErrMsg:=strErrMsg)
        If cnnAppDefaults Is Nothing Then
            strMsg = "Can't open defaults database: " & SQ(strDefaultsMdbFullname) & vbNewLine & _
                      strErrMsg & vbNewLine & vbNewLine & _
                      "Possible causes:" & vbNewLine & _
                      " - You are trying to run more than one instance of the program." & vbNewLine & _
                      " - Another program (possibly Access) has " & strDefaultsMdbFullname & " open."
            MsgBox strMsg, vbExclamation
        Else
            If Not UpdateAppDefaultsDbStructure(pCnn:=cnnAppDefaults, pErrMsg:=strErrMsg) Then
                strMsg = "Defaults database doesn't have correct structure for this version." & vbNewLine & _
                          Bracket(strDefaultsMdbFullname) & vbNewLine & vbNewLine & _
                         "Could not update structure." & vbNewLine & _
                          strErrMsg
                MsgBox strMsg, vbExclamation
            Else
                Set g.rstAppDefaults = GetAppDefaultsRst(pCnnAppDefaults:=cnnAppDefaults, pErrMsg:=strErrMsg)
            
            '   cnnAppDefaults variable cleared without closing cnn because
            '   cnn is referenced by g.rstAppDefaults and must remain open
                Set cnnAppDefaults = Nothing
                
                intPrevMousePointer = SetMousePointer(pMousePointer:=vbHourglass)
                glMaximumFranchises = gconReservedFranchiseID - 1
                g.bMaster = (UCase$(g.rstAppDefaults!NodeType) = "MASTER")
                g.bBusinessMgrPC = (UCase$(g.rstAppDefaults!NodeType) = "BUSINESS MANAGER PC")
                If g.bMaster Then
                    g.strNodeType = "Master"
                Else
                    g.strNodeType = g.rstAppDefaults!NodeType
                End If
                
               'GetComputerName fails if the input size is less than MAX_COMPUTERNAME_LENGTH + 1.
               ' MAX_COMPUTERNAME_LENGTH is 32 in Windows
                strPcName = Space$(64)
                GetComputerName strPcName, 64
            '   Trim trailing Null character. [vbNullChar = Chr(0)]
                g.strNodeName = Left$(strPcName, InStr(1, strPcName, vbNullChar, vbBinaryCompare) - 1)
                
                giTopSellers = gconTopSellersDefault
                gbEventLogRefreshIsEnabled = True
                gbEventLogRefreshIsNotAlreadyInProgress = True
                
                strCnnString = Cn(g.rstAppDefaults!MySqlCnnString, vbNullString) ''' V376 is handling Null necessarry!!
                Set g.cnnDW = GetCnnMySqlFromCnnString(pCnnString:=strCnnString, pErrMsg:=strErrMsg)

                If g.cnnDW Is Nothing Then
                    strMsg = "Can't connect to databse." & vbNewLine & "Error Message: " & strErrMsg
                    MsgBox strMsg, vbExclamation
                    SetMousePointer pMousePointer:=intPrevMousePointer
                    Set frmGetCnn = New fdlgGetCnnMySQL
                    Set g.cnnDW = frmGetCnn.GetCnn(pCnnString:=strCnnString)
                    Set frmGetCnn = Nothing
                    If Not g.cnnDW Is Nothing Then
                    '   Case sensitive comparison so user can change case to own preference
                        If strCnnString <> g.rstAppDefaults!MySqlCnnString Then
                            strMsg = "Save new connection settings"
                            If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
                                g.rstAppDefaults!MySqlCnnString = strCnnString
                                g.rstAppDefaults.Update
                            End If
                        End If
                    End If
                    SetMousePointer pMousePointer:=vbHourglass
                End If
                
                If Not g.cnnDW Is Nothing Then
                    UpdateDwDbStructure ''' Review: Not sure if should be in enclosed in "If g.bMaster ..."
                    SetGlobalDwRsts
                    Load frmTSGDataWarehouse
                End If
                
            End If
        End If
    
    End If

End Sub

Private Sub SetGlobalDwRsts()
Dim strErrMsg As String

''' ALL VALUES READ FROM THIS RST COULD BE READ HERE INTO A GLOBAL UDT AND
''' THERE WOULD BE NO NEED TO KEEP THE RST OPEN.
    Set g.rstDWDefaults = GetRst(pCnn:=g.cnnDW, _
                                 pSource:="Defaults", _
                                 pSourceType:=adCmdTable, _
                                 pRstType:=eEditableDynamic, _
                                 pErrMsg:=strErrMsg)
    
''' APART FROM SET DEFAULTS()!?, ONLY ONE PLACE DEFAULTS TABLE WRITTEN TO
    g.bAutoDataCapture = CBool(g.rstDWDefaults!AutoDataCaptureCycle)
''' Set g.rstEventLog = GetRstAddOnly(pCnn:=g.cnnDW, pSource:="EventLog", pErrMsg:=strErrMsg)   ''' V397

End Sub

Public Sub SetTableUpdateTime(ByVal pTableName As String, ByVal pTimeStamp As Date)
''' Review: This procedure should be replaced by Triggers in the TsgDw db
' pTimeStamp parameter b/c call to this procedure may be delayed while another call is made
' (eg. fetching auto-increment field before another row is added), but we want to pass the
' a time saved at table update and passed to this proc after subsequent calls

'  CALLS TO THIS PROCEDURE WOULD BE MUCH BETTER REPLACED WITH A SERIES OF DB TRIGGERS
'  IF THIS RPOCEDURE IS TO REMAIN IN USE IT SHOUDL BE REWRITTEN TO USE SQL RATHER THAN RST
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    strSQL = "SELECT * FROM tblTableTimeStamps " & vbNewLine & _
             "WHERE TableName = " & SqlQ(pTableName)
    Set rst = GetRst(pCnn:=g.cnnDW, _
                     pSource:=strSQL, _
                     pSourceType:=adCmdText, _
                     pRstType:=eEditableFwdOnly, _
                     pErrMsg:=strErrMsg)
    With rst
        If (.BOF And .EOF) Then
            .AddNew
            .Fields!TableName = pTableName
        End If
            .Fields!DataUpdated.Value = pTimeStamp
        .Update
        .Close
    End With
    Set rst = Nothing
    
End Sub

Public Function GetTableUpdateTime(ByVal pTableName As String) As Date
Dim dtmResult As Date
Dim strSQL As String

    strSQL = "SELECT DataUpdated FROM tblTableTimeStamps " & vbNewLine & _
             "WHERE TableName = " & SqlQ(pTableName)
    
    dtmResult = GetRstVal(pCnn:=g.cnnDW, _
                          pSource:=strSQL, _
                          pDefaultVal:=fdtmYesterday)
    
    GetTableUpdateTime = dtmResult

End Function

Private Function UpdateAppDefaultsDbStructure(ByRef pCnn As ADODB.Connection, _
                                              ByRef pErrMsg As String) As Boolean
'   Create/delete tables & fields as required for current version
'   Return True if required fields & tables already exist or were updated else return False
Dim bFieldExists As Boolean
Dim strErrMsg As String
Dim rst As ADODB.Recordset

'.  Add new fileds
'
    Set rst = GetRst(pCnn:=pCnn, _
                     pSource:="Defaults", _
                     pSourceType:=adCmdTable, _
                     pErrMsg:=strErrMsg)
    
    If Not rst Is Nothing Then
        bFieldExists = IsFldExists(pRst:=rst, pFldName:="MySqlCnnString")
        rst.Close
        Set rst = Nothing
        If Not bFieldExists Then
        '   use DDL?
        '   g.cnnDW.Execute "ALTER TABLE `defaults` ADD COLUMN `LiveDataUpdated` DATETIME NOT NULL DEFAULT '2000-01-01'"
            bFieldExists = AppendColumn(pCnn:=pCnn, _
                                        pTableName:="Defaults", _
                                        pColumnName:="MySqlCnnString", _
                                        pColumnType:=adVarWChar, _
                                        pColumnSize:=255, _
                                        pPropertyNamesArray:=Array("Nullable", _
                                                                   "Jet OLEDB:Allow Zero Length", _
                                                                   "Description"), _
                                        pPropertyValsArray:=Array(False, _
                                                                  True, _
                                                                  "Field added by for MySQL (by program code)"), _
                                        pErrMsg:=strErrMsg)
        End If
    End If
    
'''.  ZPrefix newly unused fields (but for now only in Development environment)
'''
'''    If IsIDE() Then
'''        Debug.Print "Decommission/Delete superseded fields and tables (for now only in IDE) - AppDefaultMdb"
'''    End If
    
    If Not bFieldExists Then
        pErrMsg = "Error in UpdateAppDefaultsDbStructure()" & vbNewLine & strErrMsg
    End If
    
    UpdateAppDefaultsDbStructure = bFieldExists

End Function

Public Sub UpdateDwDbStructure()    ' Public because also called when running with Access mdb
'   Procedure operates on global g.cnnDw object making changes to
'   Access Mdb or MySQL db depending on what db program is connected to
Dim strDDL As String
Dim strTblName As String
Dim strFldName As String
Dim strZPrefix As String
Dim strDataTypeMySQL As String
'Dim strErrMsg As String
'Dim rst As ADODB.Recordset

' ###############################################################################################
' FOR MySQL & may as well be for Access as well I will need to excecute DDL directly on the cnn #
' ###############################################################################################


'.   Creates/deletes tables & fields as required for current version
'

'.  Add new tables
'
    If Not (g.cnnDW Is Nothing) Then
    '   Add tblTableTimeStamps
        If Not IsTableExists(pCnn:=g.cnnDW, pTblName:="tblTableTimeStamps") Then
            strDDL = "CREATE  TABLE tblTableTimeStamps (" & _
                        "TableName VARCHAR(25) NOT NULL ," & _
                        "DataUpdated DATETIME NOT NULL ," & _
                        "PRIMARY KEY (TableName) )"
            CnnDwExecute pCommandText:=strDDL
        End If
    End If
    
    
    
''.  Add new fileds
''
''' V389 Start
'''' Neutered b/c Neil realieed we weren't getting full days data from POS stores
'''' They poll every two hours but not on shutdown, so the last sales for the day areen't
'''' included but the other sales of day will be reported as complete sales for the day to BATA
'If False = True Then
''' If Not IsColumnExist(pCnn:=g.cnnDW, pTblName:="defaults", pColName:="DaysOfPosLiveToTfr") Then
    If Not IsColumnExist(pCnn:=g.cnnDW, pTblName:="defaults", pColName:="DefaultDaysOfPosLiveToTfr") Then
        strDDL = "ALTER TABLE `defaults` " & _
                  "ADD COLUMN DefaultDaysOfPosLiveToTfr " & _
                   "INTEGER " & _
                   "NOT NULL " & _
                   "DEFAULT 3 " & _
                  "COMMENT " & _
                    "'By default TsgDw.exe will transfer data with Tx dates from this many days " & _
                    "before today to yesterday (including yesterday) from PosLive to LiveData.'" & _
                  "AFTER `AutoDataCaptureCycle` ;"
        CnnDwExecute strDDL
    End If
'End If
''' V389 End

'    Dim cnnAppDefaults As ADODB.Connection
'    If Not g.rstAppDefaults Is Nothing Then
'        If Not TsgTADO.IsFldExists(pRst:=g.rstAppDefaults, pFldName:="MySqlCnnString") Then
'            MsgBox "Blah!"
'            Set cnnAppDefaults = g.rstAppDefaults.ActiveConnection
'            g.rstAppDefaults.Close
'            g.cnnDW.Execute "ALTER TABLE `defaults` ADD COLUMN `MySqlCnnString` VARCHAR(255) NOT NULL DEFAULT ''"
'            MsgBox "g.cnnDW.Errors.Count: " & g.cnnDW.Errors.Count
''MySqlCnnString
''''        Else
''''            MsgBox "Field exists"
'        End If
'    End If
    

'    '   Add LiveDataUpdated to TsgDw Defaults table
'        Set rst = GetRst(pCnn:=g.cnnDW, _
'                         pSource:="Defaults", _
'                         pSourceType:=adCmdTable, _
'                         pRstType:=eReadOnlyFwdOnly, _
'                         pErrMsg:=strErrMsg)
'        If Not rst Is Nothing Then
'            If Not TsgTADO.IsFldExists(pRst:=rst, pFldName:="LiveDataUpdated") Then
'                rst.Close
'                Set rst = Nothing
'                g.cnnDW.Execute "ALTER TABLE `defaults` ADD COLUMN `LiveDataUpdated` DATETIME NOT NULL DEFAULT '2000-01-01'"
'            End If
'        '   Set value in new field in Access (already set in MySQL)
'        '   ALSO CODE BELOW DID NOT WORK ON FIRST ATTEMPT ON MySQL NOT SURE IF THIS WILL ALWAYS BE
'        '   THE CASE OR WHETHER DEV ENVIRONMENT IS SCREWIF. MAY NEED TO UPDATE WITH SQL DML
'        '   SO THAT DATE VALUES CAN BE APPROPRIATELY QUOTED ACCORDING TO THE TYPE OF DATABASE ATTACHED
'            If Not g.bMySQL Then
'                Set rst = GetRst(pCnn:=g.cnnDW, _
'                                 pSource:="Defaults", _
'                                 pSourceType:=adCmdTable, _
'                                 pRstType:=eEditableFwdOnly, _
'                                 pErrMsg:=strErrMsg)
'                rst!LiveDataUpdated = DateValue("1 January 2000")
'                rst.Update
'                rst.Close
'                Set rst = Nothing
'            End If
'        End If
'
'
'    '   Add PromoNonCompliantUpdated to TsgDw Defaults table
'        Set rst = GetRst(pCnn:=g.cnnDW, _
'                         pSource:="Defaults", _
'                         pSourceType:=adCmdTable, _
'                         pRstType:=eReadOnlyFwdOnly, _
'                         pErrMsg:=strErrMsg)
'        If Not rst Is Nothing Then
'            If Not TsgTADO.IsFldExists(pRst:=rst, pFldName:="PromoNonCompliantUpdated") Then
'                rst.Close
'                Set rst = Nothing
'                g.cnnDW.Execute "ALTER TABLE `defaults` ADD COLUMN `PromoNonCompliantUpdated` DATETIME NOT NULL DEFAULT '2000-01-01'"
'            End If
'        '   Set value in new field in Access (already set in MySQL)
'        '   ALSO CODE BELOW DID NOT WORK ON FIRST ATTEMPT ON MySQL NOT SURE IF THIS WILL ALWAYS BE
'        '   THE CASE OR WHETHER DEV ENVIRONMENT IS SCREWIF. MAY NEED TO UPDATE WITH SQL DML
'        '   SO THAT DATE VALUES CAN BE APPROPRIATELY QUOTED ACCORDING TO THE TYPE OF DATABASE ATTACHED
'            If Not g.bMySQL Then
'                Set rst = GetRst(pCnn:=g.cnnDW, _
'                                 pSource:="Defaults", _
'                                 pSourceType:=adCmdTable, _
'                                 pRstType:=eEditableFwdOnly, _
'                                 pErrMsg:=strErrMsg)
'                rst!PromoNonCompliantUpdated = DateValue("1 January 2000")
'                rst.Update
'                rst.Close
'                Set rst = Nothing
'            End If
'        End If


    '.  Decommission/Delete superseded fields and tables (for now only in development environment)
    '
    '   Only decommision fields and tables in development environment (at home) until going live
        If IsIDE() Then
            strZPrefix = Format$(Date, "Z_yyyy_mm_dd_")

    '''            Debug.Print "Decommission/Delete superseded fields and tables (for now only in IDE) - TsgDw.Mdb"
            '   Z_Prefix
            '   - Franchises table: FranchiseLastMsg, FranchisePriceUpdate, UseVpn, FranchiseRebate
            
            '   Construct a string with multiple alter field statements separated with ";" and vbNewLine
            '   Check fields exist first before add a line to ALTER them
    ' DwSqlQIdentifier
    'asdf '   DDL SYNTAX Varies between JetSQL & MySQL  (Could spend a brief period of time looking for a std that both can use
    'asdf '   but probably best to create a ZPrefix script to be run from inside MySQL (could have some conditional code here!"
    
    '**     THIS COULD ALL BE IN A LOOP     **'
    
            '.  Decommision TABLES
            '
        ''' If g.bMySQL Then    ''' V384
            '
                strTblName = "Summary"
                If IsTableExists(pCnn:=g.cnnDW, pTblName:=strTblName) Then
                    '** From Workbench ** ALTER TABLE `movedb_tsgdw`.`summary` RENAME TO  `movedb_tsgdw`.`ZUrbismos summary` ;
                ''' V384 Start
                ''' strDDL = "ALTER TABLE " & DwSqlQIdentifier(strTblName) & _
                '''          "RENAME TO " & DwSqlQIdentifier(strZPrefix & strTblName) ''' & _
                '''          " VARCHAR (50);" ''' NULL DEFAULT NULL  ;"
                    strDDL = "ALTER TABLE " & SqlQIdentifier(strTblName, eMySql) & _
                             "RENAME TO " & SqlQIdentifier(strZPrefix & strTblName, eMySql) ''' & _
                             " VARCHAR (50);" ''' NULL DEFAULT NULL  ;"
                ''' V384 End
                    CnnDwExecute pCommandText:=strDDL
                End If
                    
            '.  Decommision FIELDS
            '
                strTblName = "Franchises"
                strFldName = "FranchiseLastMsg"
                strDataTypeMySQL = "VARCHAR(50)"
                If IsColumnExist(pCnn:=g.cnnDW, pTblName:=strTblName, pColName:=strFldName) Then
                '** From Workbench ** ALTER TABLE `movedb_tsgdw`.`franchises` CHANGE COLUMN `Z_2013-07-11 FranchiseLastMsg` `FranchiseLastMsg` VARCHAR(50) NULL DEFAULT NULL  ;
                ''' V384 Start
                ''' strDDL = "ALTER TABLE " & DwSqlQIdentifier(strTblName) & _
                '''          "CHANGE COLUMN " & DwSqlQIdentifier(strFldName) & " " & DwSqlQIdentifier(strZPrefix & strFldName) & _
                '''          " " & strDataTypeMySQL & ";" ''' NULL DEFAULT NULL  ;"
                '''         ''' " VARCHAR (50);" ''' NULL DEFAULT NULL  ;"
                    strDDL = "ALTER TABLE " & SqlQIdentifier(strTblName, eMySql) & _
                             "CHANGE COLUMN " & SqlQIdentifier(strFldName, eMySql) & " " & SqlQIdentifier(strZPrefix & strFldName, eMySql) & _
                             " " & strDataTypeMySQL & ";" ''' NULL DEFAULT NULL  ;"
                            ''' " VARCHAR (50);" ''' NULL DEFAULT NULL  ;"
                ''' V384 End
                    CnnDwExecute pCommandText:=strDDL
                End If
        ''' Else                                    ''' V384
        ''' '   ACCESS MDB WILL NEED TO USE ADOX    ''' V384
        ''' End If                                  ''' V384
        
        End If

    '       ALTER TABLE `movedb_tsgdw`.`defaults` DROP COLUMN `LiveDataUpdated` ;
    '        DeleteColumn pCnn:=pCnnDw, pTableName:="Defaults", pColumnName:="Blah"             '
    
    '.  Delete superseded tables
    '

End Sub

Public Sub CnnDwExecute(ByVal pCommandText As String, _
               Optional ByRef pRecordsAffected As Long, _
               Optional ByVal pOptions As Long = ExecuteOptionEnum.adExecuteNoRecords)

' Wrapper procedure for execute method of g.cnnDw
'   1. Ensures only a MASTER instance of TsgDw makes data changes (via SQL) on g.cnnDW
'      (Relies on all calls to g.cnnDw.Execute being made through this wrapper function)
'      (cf changes made via rsts or SQL executed directly in code by g.cnnDw.Execute   )
'   2. Wraps pCommandText/SQL in a transaction so multiple SQL statements separated by a semi-colon
'      are executed as a group with only one trip/call to the database server.
'Parameters (From ADO Help with minor modifications)
' pCommandText
'   A String value that contains the SQL statement, stored procedure, a URL,
'   or provider-specific text to execute. Optionally, table names can be used
'   but only if the provider is SQL aware. For example if a table name of "Customers"
'   is used, ADO will automatically prepend the standard SQL Select syntax to form
'   and pass "SELECT * FROM Customers" as a T-SQL statement to the provider.
' pRecordsAffected
'   Optional. A Long variable to which the provider returns
'   the number of records that the operation affected.
' pOptions
'   Optional. A Long value that indicates how the provider should evaluate the CommandText
'   argument. Can be a bitmask of one or more CommandTypeEnum or ExecuteOptionEnum values.
'
' Note   Use the ExecuteOptionEnum value adExecuteNoRecords to improve performance by
' minimizing internal processing. Do not use the CommandTypeEnum values of adCmdFile or
' adCmdTableDirect with Execute. These values can only be used as options with the Open
' and Requery methods of a Recordset.
Const kProcName As String = "CnnDwExecute"
Const kSpace As String = " "
Dim bNonMasterWrite As Boolean
Dim dtmNow As Date
Dim lngError As Long
Dim strSQL As String
Dim strError As String
Dim strErrMsg As String
Dim strErrLogMsg As String
Dim strLogFile As String
Dim astrWords() As String

    dtmNow = Now
    
'   Replace vbNullChar chars. Strings retruned from [some] windows dlls are terminated with vbNullChar,
'   and will caused a bug when not trimmed and inadvertently embedded in a SQL string.
    strSQL = Replace$(Expression:=pCommandText, Find:=vbNullChar, Replace:=kSpace)

'   Only Master instance of TsgDw can manipulate db data
'   (still thinking about whether all instances can manipulate data definitions (tables, fields etc)
    If Not g.bMaster Then
        strSQL = Replace$(strSQL, Find:=vbNewLine, Replace:=" ", Compare:=vbTextCompare)
        If InStr(strSQL, vbCr) Then
            strSQL = Replace$(strSQL, Find:=vbCr, Replace:=" ", Compare:=vbTextCompare)
        End If
        If InStr(strSQL, vbLf) Then
            strSQL = Replace$(strSQL, Find:=vbLf, Replace:=" ", Compare:=vbTextCompare)
        End If
        If InStr(strSQL, vbTab) Then
            strSQL = Replace$(strSQL, Find:=vbTab, Replace:=" ", Compare:=vbTextCompare)
        End If
        astrWords = Split(strSQL, Delimiter:=" ")
    
    '   Raise error for any SQL statements with INSERT, UPDATE or DELETE
        Select Case True
            Case UBound(Filter(SourceArray:=astrWords, Match:="INSERT", Compare:=vbTextCompare)) > -1
                bNonMasterWrite = True
            Case UBound(Filter(SourceArray:=astrWords, Match:="UPDATE", Compare:=vbTextCompare)) > -1
                bNonMasterWrite = True
            Case UBound(Filter(SourceArray:=astrWords, Match:="DELETE", Compare:=vbTextCompare)) > -1
                bNonMasterWrite = True
        End Select

    End If

    If bNonMasterWrite Then
        Err.Raise Number:=666, _
                  Source:="CnnDwExecute()", _
                  Description:="Non MASTER instance of TsgDw.exe tried to alter data."
    Else
    '   TsgDw opens MySQL db with options to allow multiple SQL statements separated by a semi-colon
    '   Wrapping the Execute method in a transaction ensures groups of statements are executed in a
    '   single transaction. This may or may not be desirable. If a single SQL statement has
    '   been passed it may have a discernible impact on performance (on the downside - trip to server
    '   to begin Txn, trip to server to execute statement, trip to server to commit Txn), however for
    '   a large number of SQL statements concatenated into one, the performance impact will
    '   depend only on how the MySQL database implements transactions
        
        Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx
        On Error GoTo CnnExecuteFailed
            g.cnnDW.Execute strSQL, RecordsAffected:=pRecordsAffected, Options:=pOptions
        On Error GoTo 0 ' Disable error handler
        Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx

    End If


Procedure_Exit:
    Exit Sub

CnnExecuteFailed:
'   One alternative would be to have no error handling in this procedure and let VbWatch2
'   handle it. It would be simpler and possibly better however current code may provide
'   a bridge to handling some of the errors such as when MySQL server goes away.
    
'   We may as well re-raise the trapped error and get an exact line number
'   and local var values from the calling procedure rather than recreate
'   the error with a resume which would always happen on the .Execute line
'   in this procedure and not help locate the line in the calling procedure
'   (especially if the calling procedure had multiple calls to CnnDwExecute)

'   Perhaps should be code here attempting to address any problems executing against g.cnnDW
'   (e.g. 1.retries, 2.recovery of connection to g.cnnDW ['MySQL server has gone away'], ...)
'   Nb some calls to this proc implement their own retry attempts. eg gsubAddToLocalEventLog()
'   For the time being at least the Txn stack will be maintained by rolling back a failed executions by
'   popping the BeginTx from the stack. Without this retry attempts in calling procedures would not work.
'   The commented out lines below are a sample of ways to try to reestablish a lost cnn when this
'   exception is handled in this error handler
'   strCnnString = Cn(g.rstAppDefaults!MySqlCnnString, vbNullString)
'   Set g.cnnDW = GetCnnMySqlFromCnnString(pCnnString:=strCnnString, pErrMsg:=strErrMsg)

    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx

    lngError = Err.Number
    strError = Err.Description
    
    strErrMsg = strError
    If Len(strErrMsg) <> 0 Then
        strErrMsg = TrimWhiteSpace(StripBracketedPrefixes(strErrMsg))
    End If
    
    If g.cnnDW.Errors.Count <> 0 Then
        With g.cnnDW.Errors(0)
        '   Currently only check and report the most recently added error {ie .Errors(0)}
        '   Somtimes VBA and connection errors are identical
            If strError <> .Description Then
'''         ' Disregard non-fatal, informational connection messages {ie where err number = 0}
'''         ' {eg [Microsoft][ODBC Driver Manager] Driver's SQLSetConnectAttr failed}
'''             If .Number <> 0 Then
                    strErrMsg = strErrMsg & vbNewLine & _
                                TrimWhiteSpace(StripBracketedPrefixes(.Description))
'''             End If
            End If
        End With
    End If

    strErrLogMsg = dtmNow & " Error: " & SQ(strErrMsg) & vbNewLine & _
                   dtmNow & " CnnDwExecute() pCommandText = " & vbNewLine & SQ(pCommandText)
    
    strErrMsg = "Error: " & SQ(strErrMsg) & vbNewLine & vbNewLine & _
                "CnnDwExecute() pCommandText = " & vbNewLine & SQ(pCommandText)

'   Attempts to execute SQL against g.cnnDW have failed, so there is a
'   problem and use of a warning MsgBox before forcing program shutdown is Ok
    strLogFile = g.strLogFolder & "\SqlExecution_ERRORS.txt"
    AppendTextFile pFullName:=strLogFile, pText:=strErrLogMsg
    
    strErrMsg = "ERROR " & vbNewLine & _
                "See file " & strLogFile & _
                 vbNewLine & vbNewLine & _
                 strErrMsg
    MsgBox strErrMsg, vbCritical, App.Title
    
'   After a Resume statement a handler is deactivated BUT STILL ENABLED.
'=> Disable 'On Error Goto CnnExecuteFailed' handler with 'On Error Goto 0'
'   before raising error we want handled by VbWatch2 that will report
'   the error along with a call stack and numerous other diagnostic details
    On Error GoTo 0
    Err.Raise Number:=lngError, _
              Source:="CnnDwExecute()", _
              Description:="Error trapped and re-raised in CnnDwExecute()." & vbNewLine & _
                           "Match error report with entry in " & strLogFile & vbNewLine & _
                           strError
    End     ' Failsafe bailout!

End Sub

Public Sub TfrPosLiveToPreLive(ByVal pBwDatesSqlClause As String)
'''  Not called in subCaptureData() b/c Neil realieed we weren't getting full
'''  days data from POS stores. They poll every two hours but not on shutdown,
'''  so the last sales for the day aren't included but will the other sales of
'''  day will be reported as complete sales for the day to BATA
Const kSqlBase As String = _
    "INSERT INTO prelivedata(" & vbNewLine & _
        "FranchiseIDTSG, " & vbNewLine & _
        "Barcode, " & vbNewLine & _
        "TransactionDate, " & vbNewLine & _
        "Quantity, " & vbNewLine & _
        "TotalInc,  " & vbNewLine & _
        "NormalSellInc,  " & vbNewLine & _
        "CostInc,  " & vbNewLine & _
        "WholesaleQty, " & vbNewLine & _
        "WholesaleActualSell) " & vbNewLine & _
    "SELECT " & vbNewLine & _
        "FranchiseIDTSG, " & vbNewLine & _
        "Barcode, " & vbNewLine & _
        "Date(TransactionDate), " & vbNewLine & _
        "Sum(Quantity), " & vbNewLine & _
        "Sum(TotalInc), " & vbNewLine & _
        "NormalSellInc, " & vbNewLine & _
        "CostInc, " & vbNewLine & _
        "Sum(WholesaleQty),  " & vbNewLine & _
        "Sum(WholesaleActualSell) " & vbNewLine & _
    "FROM poslivedata "


' MINOR FLAW/BUG
' --------------
' OPos: NormalSellInc & CostInc vary slightly depending on the qty sold (rounding? OPos uses 4 dec places)
' Previously the following was written in these notes "NormalSellInc & CostInc remain the same during
' promotions and only change for price rises twice a year." Not sure whether this applies to OPos in that
' they remain essentialy the same except for small rounding differences depending on rounding

' If a price rise is implemented in the middle of a work day, grouping summarises pre and post price
' change records into a single summary record. This avoides eliminating sales data when transferring
' records from pre-live to live when duplicate checking of date, barcode pairs occurs.

Const kGroupBy As String = _
    "GROUP BY " & vbNewLine & _
        "FranchiseIDTSG, " & vbNewLine & _
        "Barcode," & vbNewLine & _
        "Date(TransactionDate)"
''' Version 399 Start
''' Remove grouping by NormalSellInc & CostInc b/c rounding differences with different
''' quantities give different prices. Unicenta OPos goes to 4 decimal points on currency
'''     "NormalSellInc, " & vbNewLine & _
'''     "CostInc"
''' Version 399 End

Dim intPrevMousePointer As Integer
Dim strSQL As String
Dim strWC As String
    
    intPrevMousePointer = SetMousePointer(vbHourglass)
    strWC = "WHERE TransactionDate " & pBwDatesSqlClause
    strSQL = kSqlBase & vbNewLine & strWC & vbNewLine & kGroupBy
    
On Error GoTo Procedure_Error
    CnnDwExecute strSQL
    
''' Review - Uinsg pBwDatesSqlClause for StatusBar msgS is not very human reader friendly
'''          Would be nice to review both the dialog form to get the dates (ADDNING
'''          another return value type) and this procedure to fix things up.
    StatusBar "Transferred PosLiveDate to PreLiveData for tx dates " & pBwDatesSqlClause
    
Procedure_Exit:
    SetMousePointer intPrevMousePointer
    Exit Sub

Procedure_Error:
'   It appears (at least from IDE) when you access g.cnnDW.Errors(0).Description from a MySQL cnn, the
'   g.cnnDW.Errors collection is cleared. -> always check (as is good practice) g.cnnDW.Errors.Count
'   before accessing g.cnnDW.Errors(0).Description. Also appears that if g.cnnDw.Errors was the last
'   error then this error will populate vba.Err.Description
    StatusBar UCase$("FAILED TRANSFER OF PosLiveDate to PreLiveData for tx dates ") & _
              pBwDatesSqlClause & ". " & VBA.Err.Description
    Resume Procedure_Exit
    Resume  ' Not executed but assists when debugging in IDE

End Sub

Public Sub OptimiseDb()
'Redirection to an output file works but only when executing a file that has the
'redirection rather than directly executing a command with a redirection In the command.
'If we want a copy of the output we have to create a batch file'and execute it. Shame then that I
'have to include the root user and pwd in the file that would be there for all to see while the
'command Is executing. I hide it a bit by writing both the bat and output to temporary file
'locations and deleting both of them once I am finished with them
'   This is how I should pass the user name and pwd parameters to the batch file and wait
'   I think it was something like Echo root pwd | mycheck blah user=%1, pwd=%2, but I found
'   problems doing this b/c Echo is not really providing stdio. maybe I need to use a filter like
'   more after the echo and do something else. Better off doing it in Windows Powershell
'   I could then write the batch file to any folder. Perhaps best to leave it in the app folder
'   for troubleshooting and diagnostic purposes.
'    If fso.FileExists(pFullName) Then
'        ExecuteAndWait pFullName & " " & pCommandLine
'    End If
Dim bPrevFormEnabled            As Boolean
Dim intPrevMousePointer         As Integer
Dim strTempFolder               As String
Dim strSupportBatchCmds         As String
Dim strTempBatchFileFullname    As String
Dim strMySqlBaseCmd             As String
Dim strUser                     As String
Dim strPwd                      As String
Dim strDb                       As String
Dim strExtendedProperties       As String
Dim strMySqlSupportCmd          As String
Dim strMySqlTempCmd             As String
Dim strSupportBatchFileFullname As String
Dim strSupportOutFullFilename   As String
Dim strTempBatchCmds            As String
Dim strLine                     As String
Dim fso                         As Scripting.FileSystemObject
Dim ts                          As Scripting.TextStream

    bPrevFormEnabled = SetFormEnabled(pForm:=Screen.ActiveForm, pEnabled:=False)
    intPrevMousePointer = SetMousePointer(vbHourglass)
    DoEvents    ' Give display a chance to update before executing time consuming batch file
    
    Set fso = New Scripting.FileSystemObject
    strTempFolder = fso.GetSpecialFolder(SpecialFolder:=TemporaryFolder)
    strTempBatchFileFullname = strTempFolder & "\" & fso.GetBaseName(fso.GetTempName) & ".bat"
    strSupportOutFullFilename = App.Path & "\OptimiseDwDatabaseOutfile.txt"
    strSupportBatchFileFullname = App.Path & "\Optimise_Dw_Database.bat"
    strExtendedProperties = g.cnnDW.Properties("Extended Properties")
    strUser = GetValStringValue(pVString:=strExtendedProperties, pVName:="UID")
    strPwd = GetValStringValue(pVString:=strExtendedProperties, pVName:="PWD")
    strDb = GetValStringValue(pVString:=strExtendedProperties, pVName:="DATABASE")
    strMySqlBaseCmd = "mysqlcheck " & _
                        "--databases " & strDb & " " & _
                        "--medium-check " & _
                        "--auto-repair " & _
                        "--force " & _
                        "--optimize " & _
                        "--analyze " & _
                        "--check-only-changed "
    strMySqlSupportCmd = strMySqlBaseCmd & "--user=%1 --password=%2"
    strMySqlTempCmd = strMySqlBaseCmd & "--user=" & strUser & " --password=" & strPwd
'   CLS used in support batch to remove usr & pwd entered on cmd line, a smart user would still be able to recall it
'   with history but it wouldn't be sitting on the screen for a couple of minutes for everyone who walked past to see.
    strSupportBatchCmds = "@ECHO Off > Nul" & vbNewLine & "CLS rem CLS used in support batch to remove usr and pwd entered on cmd line" & vbNewLine & _
                          "IF " & DQ("%~1") & "==" & DQ(vbNullString) & " (" & vbNewLine & " REM ECHO." & vbNewLine & " ECHO Please pass MySQL user name and password" & vbNewLine & " ECHO." & vbNewLine & " ECHO e.g. " & UCase$(fso.GetBaseName(strSupportBatchFileFullname)) & " MySql_User MySql_Password" & vbNewLine & " ECHO." & vbNewLine & ") else (" & vbNewLine & _
                          " ECHO Please Wait ..." & vbNewLine & " ECHO Executing -> " & strMySqlBaseCmd & " > " & strSupportOutFullFilename & vbNewLine & _
                          strMySqlSupportCmd & " > " & strSupportOutFullFilename & vbNewLine & _
                          " ECHO Command executed" & vbNewLine & ")"
    strTempBatchCmds = "@ECHO Off > Nul" & vbNewLine & _
                       " ECHO Please Wait ..." & vbNewLine & " ECHO Executing -> " & strMySqlBaseCmd & " > " & strSupportOutFullFilename & vbNewLine & _
                       strMySqlTempCmd & " > " & strSupportOutFullFilename & vbNewLine & _
                       " ECHO Command executed" & vbNewLine & ")"
    
    SaveTextFile strTempBatchFileFullname, pFileText:=strTempBatchCmds, pOverwrite:=True
    SaveTextFile strSupportBatchFileFullname, pFileText:=strSupportBatchCmds, pOverwrite:=True
    
    StatusBar "Optimising database ..."
    StatusBar strMySqlBaseCmd
    ExecuteAndWait cmdline:=strTempBatchFileFullname
    fso.DeleteFile strTempBatchFileFullname, Force:=True
    
    Set fso = New Scripting.FileSystemObject
    Set ts = fso.OpenTextFile(strSupportOutFullFilename, ForReading)
    Do While Not ts.AtEndOfStream
        strLine = ts.ReadLine
        StatusBar strLine, pRefreshEventLogDisplay:=False
    Loop
    ts.Close    ' Close ts and reclaim memory (must be closed before file can be moved)
    Set ts = Nothing
    Set fso = Nothing
    
    StatusBar vbNullString, pLog:=False, pRefreshEventLogDisplay:=False     ' Clear Status Bar
    SetMousePointer intPrevMousePointer
    SetFormEnabled Screen.ActiveForm, bPrevFormEnabled

End Sub
