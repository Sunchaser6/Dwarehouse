VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBataRpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const mkRemotePath As String = "\InBound"
Public Event OnRptLoad(ByRef oRpt As clsBataRpt)
Public Event AfterRptLoad(ByRef oRpt As clsBataRpt, ByVal Success As Boolean)
Public Event AfterRptUpload(ByRef oRpt As clsBataRpt, ByVal Success As Boolean, ByVal ErrMsg As String)
Public Event BeforeRptUpload(ByRef oRpt As clsBataRpt, ByVal UploadAttempt As Long)
Public Event AddUnsentComplete(ByVal Msg As String)
Private WithEvents moRpt As clsBataRpt   ' Module rpt variable declared so it can be declared WithEvents
Attribute moRpt.VB_VarHelpID = -1

Private Enum ReportingModeEnum
    eRmFtpUpload = 0
    eRmLocalFile = 1
End Enum

Private Type udt
    eReportingMode As ReportingModeEnum
    bAddedUnprocessed As Boolean
    bAddedUnsentRpts As Boolean
    strUploadSummary As String
    rpts As VBA.Collection
End Type
Dim m As udt

Public Function Add(ByVal pFranName As String, _
                    ByVal pFranID As Long, _
                    ByVal pBataFranID As Long, _
                    ByVal pTxDate As Date, _
                    ByVal pBataRptType As BataRptTypeEnum) As clsBataRpt
    
    Set moRpt = New clsBataRpt
    
    moRpt.Load pFranName:=pFranName, pFranID:=pFranID, pBataFranID:=pBataFranID, pTxDate:=pTxDate, pBataRptType:=pBataRptType
    If Not moRpt.HasData Then
        RaiseEvent AfterRptLoad(moRpt, False)
    Else
        If Not KeyExists(moRpt.Key) Then
            m.rpts.Add Item:=moRpt, Key:=moRpt.Key
            Set Add = moRpt
        End If
        RaiseEvent AfterRptLoad(moRpt, True)
    End If
    
    Set moRpt = Nothing

End Function

Public Sub AddRpts_FranAndDateSeln(ByRef pColFranIDs As VBA.Collection, _
                                   ByVal pDateFrom As Date, _
                                   ByVal pDateTo As Date)
Dim intPrevMousePointer As Integer
Dim lngFranID As Long
Dim lngBataFranID As Long
Dim dtmLoop As Date
Dim vntFranID As Variant
Dim strFranName As String
Dim strSQL As String
Dim strErrMsg As String
Dim rstFran As ADODB.Recordset

    intPrevMousePointer = SetMousePointer(vbHourglass)

    For Each vntFranID In pColFranIDs
        lngFranID = vntFranID
        strSQL = "SELECT FranchiseIDBATA, FranchiseBusinessName" & vbNewLine & _
                 "FROM qryFranchiseBata" & vbNewLine & _
                 "WHERE FranchiseIDTSG = " & lngFranID
                 
        Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
        If Not (rstFran.BOF And rstFran.EOF) Then
            lngBataFranID = rstFran!FranchiseIDBATA
            strFranName = rstFran!FranchiseBusinessName
        
        '   Loop from dtpFrom to dtpTo adding rpts [rpts are written to file when added to collection]
            For dtmLoop = pDateFrom To pDateTo
            '   Attempt to add both TotSales & Wholesale Rpt (somewhat inefficient as it may be a Sunday etc)
                If Not Me.Add(pFranName:=strFranName, _
                              pFranID:=lngFranID, _
                              pBataFranID:=lngBataFranID, _
                              pTxDate:=dtmLoop, _
                              pBataRptType:=eBataTotalSalesRpt) Is Nothing Then
                '   There were sales so may be some wholesale sales. Attempt to add Wholesale rpt
                    Me.Add pFranName:=strFranName, _
                           pFranID:=lngFranID, _
                           pBataFranID:=lngBataFranID, _
                           pTxDate:=dtmLoop, _
                           pBataRptType:=eBataWholesaleRpt
                End If
            Next dtmLoop
            
        End If
        
        rstFran.Close
        Set rstFran = Nothing
    Next vntFranID
    
    SetMousePointer intPrevMousePointer

End Sub

Private Sub AddRpts_UnProcessed()
' V389 Reinstate 'rstFran.MoveNext' inadvertantly deleted in V386 and subsequently causing a bug
''' V386 Start - Not sure if I can extract any meaning from comments below, although they once
'''              had meaning. May at a later stage be a reason to add an optional thorough search
''''   -> pThorough would check for WS if there was TotSale to rpt on and no Upload record for WS
''''ADDING [UNPROCESSED] BATA REPORTS Is THE TIME CONSUMING TASK/PROCESS In OVERNIGHT CAPUTRE  (~6HOURS In Nov12 WITHOUT PROMOS OR MSGS)
''''WHEREAS UPLOADING THE DATA ONLY TOOK AROUND 20 MINUTES ON THE SAME NIGHT. THIS PROCEDURE HOLDS THE KEY TO EVERYTHING
''''tblBataUploads HAS NOT BEEN PURGED SINCE IT BEGAN IN APRIL 2007 27/05/07
'''
''''   NEXT OPTIMISATION WOULD BE TO OUTER LOOP BY DATE (cf BY FRAN) AND CREATE rstUploads BY LOOP DATE
''''   THIS WOULD LIMIT RSTUPDATE TO A SMALLER RST. THE INNER LOOP WOULD ITERATE THROUGH A
''''   FRANCHISE COLLECTION. THIS WOULD ALLOW FOR THE FINAL OPTIMISATION OF ITERATING THE DATE LOOP
''''   FROM THE MOST RECENT DATE TO THE OLDEST DATE AND REMOVING THE FRANCHISE FROM THE FRAN COLLECTION
''''   WHEN YOU FIRST COME ACROSS A DATE THAT THE FRANCHISE ALREADY HAS UPLOADED DATA FOR.
''''   A THOROUGH SEARCH TAHT DIDN'T REMOVE THEN FRAN FROM THE FRAN COLLECTION COULD BE
''''   RUN AS REQUIRED (selected via interface, when flag for last CaptureAll indicates the
''''   process did not complete, even every Saturday night when timing is not critical just
''''   to ensure there is no missing data and we have caught up on any funny buggers)
''' V386 End
Dim lngFranID As Long
Dim lngBataFranID As Long
Dim dtmLoop As Date
Dim dtmDefaultMinDate As Date
Dim vntDateTo As Variant
Dim strErrMsg As String
Dim strSQL As String
Dim strFranName As String
Dim rstFran As ADODB.Recordset
    
''' If Not m.bAddedUnsentRpts Then  ' V401
    If Not m.bAddedUnprocessed Then ' V401
'   Most likely the only times we add rpts for unprocessed TxDates up to a date other than yesterday is when this
'   procedure is called as part of O/N processing and prior to midnight (very rare) and we collect up to today
'   We can only report on downloaded data -> check db for most recent TransactionDate in downloaded data

        strSQL = "SELECT Max(TransactionDate) FROM LiveData"                        ''' speed up by later getting date individually for each franid?
        vntDateTo = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=vbEmpty) ''' speed up by later getting date individually for each franid?
                              
        If Not IsEmpty(vntDateTo) Then
        '   We have LiveData
        '   Min date for (re-)creating unprocessed BATA reports has BIG effect
        '   on processing time for first time BATA reporting for a franchise
            dtmDefaultMinDate = DateAdd(Interval:="d", Number:=-121, Date:=Date) ''' Review *** HARD CODED LIMIT OF HOW FAR TO GO BACK
                                                                                 ''' 121 b/c loop start date is day after Max of 121 & tblUpload!UploadDate
            If dtmDefaultMinDate < g.dtmLiveDataStart Then
                dtmDefaultMinDate = g.dtmLiveDataStart
            End If
            
        '   Add Live to WHERE Clause - @Apr2013 is 35 less Bata frans out of 333 (~10%)
            strSQL = "SELECT FranchiseIDTSG, FranchiseIDBATA, FranchiseBusinessName " & vbNewLine & _
                     "FROM qryFranchiseBata " & vbNewLine & _
                     "WHERE Live " & vbNewLine & _
                     "ORDER BY FranchiseBusinessName"
                     
            Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            
        '   Loop through each franchise adding rpts not previously uploaded. Originally performed
        '   thorough search looping through past 3 months dates of each Fran picking up unprocessed data
        '   as determined by tblBataUploads, but now loop from day after last upload for Fran to
        '   MaxTxDate in LiveData assuming no missing patches of uplodaded data.
        '   Loop from FranMaxTxDateUploaded (or dtmDefaultMinDate if none uploaded) to vntDateTo
        '   [rpts are written to a text file when added to collection]

''' Review 'AN OPTIMISATION WORTH TRYING MIGHT BE TO GET THE INDIVIDUAL MAX DATA DATES FOR EACH FRAN
'AND THUS HAVE AN INDIVIDUAL vntToDate FOR EACH FRAN LOOP AS WELL AS AN INDIVIDUAL FROM
'DATE COULD EVEN HAVE A JOINED QUERY FOR BOTH TO LIMIT CALLS TO THE DATABASE SERVER!
'JOIN MAY BE B/W qryFranchiseBata & LiveData, tblBataUploads & LiveData, ...
'ALTERNATELY THE MAXDATE QUERY RST COULD BE COLLECTED AS Rst AND THEN DISCONNECTED
            
            Do While Not rstFran.EOF
                lngFranID = rstFran!FranchiseIDTSG
                lngBataFranID = rstFran!FranchiseIDBATA
                strFranName = rstFran!FranchiseBusinessName
                strSQL = "SELECT MAX(TxDate) FROM tblBataUploads " & vbNewLine & _
                         "WHERE FranID = " & lngFranID
                dtmLoop = DateAdd(Interval:="d", _
                                  Number:=1, _
                                  Date:=GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=dtmDefaultMinDate))
                                    
'LIMITING rstUploads BY Date KEEPS THE RST DOWN TO AROUND 440 ROWS AND BELOW ON CURRENT LIVE DATA OF TWO DATES
'LIMITING BY FRANID GAVE NUMBERS OF 1916 AND 3966 ROWS ON CURRENT LIVE DATA FOR THE TWO FRANS I LOOKED AT
'IF SEARCHING FROM LAST UPLOAD DATE THEN LIMITING BY FRANID AND DATE SHOULD REDUCE RST TO THE MINIMUM
' AS WELL AS SEARCHING THROUGH MARKEDLY FEWER DATES
'1. NEED TO HAVE AN AUTOMATIC PURGING OF tblBataUploads and tblBataREUploads (keep one year of data)
'   AND UNPURGED RECORDS SHOULD NOT EXTEND EARLIER THAN LIVE DATA.
'2. Need to Add a new Field for indicating how far back to G0 looking for un-uploaded BATA DATA because
'   how far you go back hugely affects the efficiency of loading unprocessed data for a new franchise
'   when there may be no upload data and we start from the earliest date
            
               Do While dtmLoop <= vntDateTo
                    '   Attempt to add both (somewhat inefficient as it may be a Sunday etc)
                        If Not Me.Add(pFranName:=strFranName, _
                                      pFranID:=lngFranID, _
                                      pBataFranID:=lngBataFranID, _
                                      pTxDate:=dtmLoop, _
                                      pBataRptType:=eBataTotalSalesRpt) Is Nothing Then
                        '   There were sales so may be some wholesale sales. Attempt to add Wholesale rpt
                            Me.Add pFranName:=strFranName, _
                                   pFranID:=lngFranID, _
                                   pBataFranID:=lngBataFranID, _
                                   pTxDate:=dtmLoop, _
                                   pBataRptType:=eBataWholesaleRpt
                         End If
                    dtmLoop = DateAdd(Interval:="d", Number:=1, Date:=dtmLoop)
'If Me.Count >= 50 Then
'Stop
'End If
                Loop
                rstFran.MoveNext     ' Reinstated V389
            Loop
            rstFran.Close
            Set rstFran = Nothing
            
        ''' m.bAddedUnsentRpts = True   ' V401
            m.bAddedUnprocessed = True  ' V401
        End If
    End If
    
    RaiseEvent AddUnsentComplete(m.rpts.Count & " reports added.")
    
End Sub

Private Sub Class_Initialize()
    Set m.rpts = New VBA.Collection

'   As object already violates OO design by using g.cnnDw
'   (application global data cnn variable), then we may as well
'   use the connection to initialise bBataFtpEnabled variable.
''' m.bBataFtpEnabled = g.rstDWDefaults!BataFtpEnabled
    g.rstDWDefaults.Requery
    If g.rstDWDefaults!BataFtpEnabled Then
        m.eReportingMode = eRmFtpUpload
    Else
        m.eReportingMode = eRmLocalFile
    End If
End Sub

Private Sub Class_Terminate()
'   Clear references to module object variables
    If m.rpts.Count Then
        Do While m.rpts.Count
            m.rpts.Remove 1
        Loop
    End If
    Set m.rpts = Nothing
End Sub

Public Property Get Count() As Long
    Count = m.rpts.Count
End Property

Public Function Item(ByVal pIndex As Variant) As clsBataRpt
Attribute Item.VB_UserMemId = 0
    Set Item = m.rpts.Item(pIndex)
End Function

Private Function KeyExists(pKey As Variant) As Boolean
Dim oRpt As clsBataRpt

    For Each oRpt In m.rpts
        If oRpt.Key = pKey Then
            KeyExists = True
            Exit For
        End If
    Next oRpt

End Function

Private Sub LogBataRptProcessed(ByRef pBataRpt As clsBataRpt)
Dim eRptType As BataRptTypeEnum
Dim lngFranID As Long
Dim dtmTxDate As Date
Dim strSQL As String
Dim strErrMsg As String
Dim strWC As String
Dim rst As ADODB.Recordset
    
    With pBataRpt
        lngFranID = .FranID
        dtmTxDate = .TxDate
        eRptType = .RptType
    End With

    strWC = "(FranID = " & lngFranID & ")" & _
       " AND (TxDate = " & MySqlDate(dtmTxDate) & ")" & _
       " AND (RptType = " & eRptType & ")"
    strSQL = "SELECT * FROM tblBataUploads WHERE " & strWC
    
    Set rst = GetRst(pCnn:=g.cnnDW, _
                     pSource:=strSQL, _
                     pSourceType:=adCmdText, _
                     pRstType:=eEditableFwdOnly, _
                     pErrMsg:=strErrMsg)
    
    If Not (rst.BOF And rst.EOF) Then
    '   Rpt already uploaded
    '   -> Close tblBataUploads rst and open tblBataREUploads rst for update
        rst.Close
        Set rst = GetRstAddOnly(pCnn:=g.cnnDW, pSource:="tblBataREUploads", pErrMsg:=strErrMsg)
    End If
        
    With rst
        .AddNew
            .Fields!TxDate = dtmTxDate
            .Fields!FranID = lngFranID
            .Fields!RptType = eRptType
            .Fields!UploadDate = Now
            .Fields!RptMode = m.eReportingMode
        .Update
        .Close
    End With
    Set rst = Nothing

End Sub

Private Sub moRpt_OnLoad(oRpt As clsBataRpt)
    RaiseEvent OnRptLoad(oRpt)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'   NewEnum must return the IUnknown interface of a collection's enumerator.
    Set NewEnum = m.rpts.[_NewEnum]
End Function

Public Sub Process(ByVal pAddUnProcessed As Boolean)
 ' **** ONLY Process BataRpts VIA THIS METHOD ****
'   By limiting uploads to this method we have the option of ensuring
'   unsent data is uploaded by non-conditionally calling AddRpts_Unsent
' NB. AddUnsentRptsQuick works backwards in time until it finds a sent report and assumes data previous
'     to the first sent report it finds has been sent. It thus relies on all unsent data being sent.
'------------------------------------------
' Extra retry code may be added if required
'Const kMaxRptCollectionUploadAttempts As Long = 2
'Dim lngRptCollectionUploadAttempts As Long
'------------------------------------------
Const kMaxRptProcessAttempts As Long = 3
Dim bProcessed As Boolean
Dim bLoggedAsProcessed As Boolean
Dim lngRptsBeforeUpload As Long
Dim lngRptsAfterUpload As Long
Dim lngRptProcessAttempt As Long
Dim strErr As String
Dim strUploadSummary As String
Dim strRemoteFullname As String
'''Dim strLocalPath As String
Dim strLocalFullname As String
Dim fso As Scripting.FileSystemObject
Dim oSFTP As clsSFTP
Dim oBataRpt As clsBataRpt
    
    If pAddUnProcessed Then
        If Not m.bAddedUnprocessed Then
            AddRpts_UnProcessed
        End If
    End If
    
    lngRptsBeforeUpload = Me.Count
    If lngRptsBeforeUpload > 0 Then
    
If m.eReportingMode = eRmFtpUpload Then
    '   Create and initialise SecureFTP wrapper object
        Set oSFTP = New clsSFTP
        With oSFTP
        '   BATA uses SFTP Protocol (default mode for clsSFTP)
            .HostAddress = g.rstDWDefaults!BataFtpHostAddress
            .Login = g.rstDWDefaults!BataFtpUser
            .Password = g.rstDWDefaults!BataFtpPwd
        End With
Else
    '   m.eReportingMode = eRmLocalFile
        Set fso = New Scripting.FileSystemObject
'''     strLocalPath = g.strBataRptsFolder & "\" & Format$(Now, "yyyy-mm-dd")
'''     If Not fso.FolderExists(strLocalPath) Then
'''         fso.CreateFolder (strLocalPath)
'''     End If
End If
'       '-----------------------------------------------------------------------------------------------------------
'       '   Retry remaining rpts in collection (even after possible kMaxRptProcessAttempts on previous attempt)
'       '   Would probably want to conditionally raise AfterRptUpload() event only for kMaxRptCollectionUploadAttempt
'           Do While (Me.Count > 0) And (lngRptCollectionUploadAttempts < kMaxRptCollectionUploadAttempts)
'       '-----------------------------------------------------------------------------------------------------------
                 For Each oBataRpt In Me
                    bLoggedAsProcessed = False
                    lngRptProcessAttempt = 0

If m.eReportingMode = eRmFtpUpload Then
                    strRemoteFullname = mkRemotePath & "\" & oBataRpt.Name
Else
'   m.eReportingMode = eRmLocalFile
'''                 strLocalFullname = strLocalPath & "\" & oBataRpt.Name
                    strLocalFullname = g.strBataRptsFolder & "\" & oBataRpt.Name
End If
                    
                    Do While (Not bLoggedAsProcessed) And (lngRptProcessAttempt < kMaxRptProcessAttempts)
                        lngRptProcessAttempt = lngRptProcessAttempt + 1
                        RaiseEvent BeforeRptUpload(oBataRpt, lngRptProcessAttempt)
                    '   Save Process success (bProcessed) for passing when raising AfterRptUpload event
                    ''' bProcessed = .Process(pLocalName:=oBataRpt.FullName, pRemoteName:=strRemoteFullname, pErrMsg:=strErr)   ''' 401
If m.eReportingMode = eRmFtpUpload Then
                        bProcessed = oSFTP.Upload(pLocalName:=oBataRpt.FullName, pRemoteName:=strRemoteFullname, pErrMsg:=strErr)
Else
'   m.eReportingMode = eRmLocalFile
                        On Error Resume Next
                            fso.CopyFile Source:=oBataRpt.FullName, Destination:=strLocalFullname, OverWriteFiles:=False
                            strErr = Err.Description
                            bProcessed = (Err.Number = 0)
                        On Error GoTo 0
End If
                        If bProcessed Then
                            LogBataRptProcessed oBataRpt    ' Passing oBataRpt provides a range of rpt properties
                            bLoggedAsProcessed = True
                        Else
If m.eReportingMode = eRmFtpUpload Then
                            
                            If oSFTP.RemoteExists(pRemoteName:=strRemoteFullname) Then
                            '   File already on remote server. Unusual situation may occur if previous
                            '   processing interupted after file uploaded but before UploadTabes updated
                            '   LogBataRptUploaded so we don't retry uploading files already present on remote server
                            'NB Logged Process failure in EventLog will assist reconciling recorded Process times (if required)
                                LogBataRptProcessed oBataRpt    ' Passing oBataRpt provides a range of rpt properties
                                strErr = "Remote File Exists " & strRemoteFullname & " - " & strErr
                                bLoggedAsProcessed = True
                            End If
Else
                            If fso.FileExists(strLocalFullname) Then
                            '   File already written. Unusual situation as it would be for other clause of If stmt
                                LogBataRptProcessed oBataRpt ' Passing oBataRpt provides a range of rpt properties  ''' V401
                                strErr = "Local Report File Exists " & strLocalFullname & " - " & strErr
                                bLoggedAsProcessed = True
                            End If
End If
                        End If
                    Loop
                    
                    If bLoggedAsProcessed Then
                    '   Remove successfully uploaded rpt. Remaining could be logged or uploads retried
                        Me.Remove oBataRpt.Key
                    Else
                        strErr = strErr & " (" & Plural(pQty:=lngRptProcessAttempt, pNounSingular:="attempt") & ")"
                    End If
                    
                '   AfterRptUpload event raised passing bUpload status event handling
                '   code can decide whether log or not log success or failures
                    RaiseEvent AfterRptUpload(oBataRpt, bProcessed, strErr)
                Next oBataRpt
'       '-----------------------------------------------------------------------------------------------------------
'       '   Retry remaining rpts in collection (even after possible kMaxRptProcessAttempts on previous attempt)
'       '   Would probably want to conditionally raise AfterRptUpload() event only for kMaxRptCollectionUploadAttempt
'           Loop    ' Do While (Me.Count > 0) And (lngRptCollectionUploadAttempts < kMaxRptCollectionUploadAttempts)
'       '-----------------------------------------------------------------------------------------------------------
'''     End With
            
        Set oBataRpt = Nothing
        Set oSFTP = Nothing
        Set fso = Nothing
        
    '   Set UploadSummary property
        lngRptsAfterUpload = Me.Count
        Select Case True
            Case lngRptsBeforeUpload = 0
                strUploadSummary = "WARNING: NO BATA REPORTS FOR PROCESSING"
            Case (lngRptsBeforeUpload <> 0) And (lngRptsAfterUpload = 0)
                strUploadSummary = "All BATA reports processed" '& " (" & Plural(pQty:=lngRptsBeforeUpload, pNounSingular:="report") & ")"
            Case (lngRptsBeforeUpload <> 0) And (lngRptsAfterUpload <> 0) And (lngRptsBeforeUpload = lngRptsAfterUpload)
                strUploadSummary = "WARNING: FAILED PROCESSING ALL BATA REPORTS" '& " (" & Plural(pQty:=lngRptsBeforeUpload, pNounSingular:="report") & ")"
            Case (lngRptsBeforeUpload <> 0) And (lngRptsAfterUpload <> 0)
                strUploadSummary = "WARNING: FAILED PROCESSING SOME BATA REPORTS - (" & lngRptsAfterUpload & " OF " & lngRptsBeforeUpload & " FAILED)"
        End Select
        m.strUploadSummary = strUploadSummary
    
    End If
    
End Sub

Public Sub Remove(ByVal pIndex As Variant)
    If KeyExists(pIndex) Then
        m.rpts.Remove pIndex
    End If
End Sub

Public Property Get UploadSummary() As String
    UploadSummary = m.strUploadSummary
End Property

