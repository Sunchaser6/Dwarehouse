VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBataRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event OnLoad(ByRef oRpt As clsBataRpt)

Public Enum BataRptTypeEnum
    eBataTotalSalesRpt = 1
    eBataWholesaleRpt = 2
End Enum

Private Type udt
    eRptType As BataRptTypeEnum
    bRptHasData As Boolean
    lngFranID As Long
    lngFranID_BATA As Long
    datTxDate As Date
    strFranName As String
    strName As String
    strFullname As String
    strRpt As String
    fso As Scripting.FileSystemObject
End Type
Private m As udt

Public Property Get HasData() As Boolean
    HasData = m.bRptHasData
End Property

Private Function IsSalesData(ByVal pFranID As Long, _
                             ByVal pTxDate As Date, _
                             pBataRptType As BataRptTypeEnum) As Boolean    ''' Review - should be passed ByVal
' Returns True if there is sales data for the nominated report otherwise returns False
' Quicker to initially test for data using the simple SELECT query
' here rather than unconditionally calling complex rpt queries
Dim bResult As Boolean
Dim strSQL As String
Dim strErrMsg As String
'''Dim strDataTable As String   ''' V386
Dim rst As ADODB.Recordset

''' V386 Start
''' If pTxDate < g.dtmLiveDataStart Then    ''' Review - No longer have an Archive LiveData table
'''     strDataTable = "LiveDataArchive" ' Linked to LiveData table in archive database
''' Else
'''     strDataTable = "LiveData"        ' LiveData table in in main database
''' End If
'''
''' strSQL = "SELECT FranchiseIDTSG FROM " & strDataTable & vbNewLine & _
'''          "WHERE (FranchiseIDTSG = " & pFranID & ")" & _
'''          " AND (TransactionDate = " & MySqlDate(pTxDate) & ")" & _
'''          " AND (Quantity <> 0)"
             
    strSQL = "SELECT FranchiseIDTSG FROM LiveData " & vbNewLine & _
             "WHERE (FranchiseIDTSG = " & pFranID & ")" & _
             " AND (TransactionDate = " & MySqlDate(pTxDate) & ")" & _
             " AND (Quantity <> 0)"
''' V386 End
             
    If pBataRptType = eBataWholesaleRpt Then
        strSQL = strSQL & " AND (WholesaleQty <> 0)"
    End If
    
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    bResult = Not (rst.BOF And rst.EOF)
    rst.Close
    Set rst = Nothing

    IsSalesData = bResult

End Function

Public Property Get Key() As String
    Key = m.lngFranID & "_" & Format$(m.datTxDate, "ddmmmyyyy") & "_" & m.eRptType
End Function

Public Property Get RptType() As BataRptTypeEnum
    RptType = m.eRptType
End Property

Public Property Get TxDate() As Date
    TxDate = m.datTxDate
End Property

Public Property Get FranID() As Long
    FranID = m.lngFranID
End Property

Sub Load(ByVal pFranName As String, pFranID As Long, ByVal pBataFranID As Long, ByVal pTxDate As Date, pBataRptType As BataRptTypeEnum)
' Outputs:  Assigns report contents to class wide string variable (m.strRpt), if no data assigns an empty string
'           Report particulars determined by passed parameters
'
'  Inputs:  pFranID     - Franchise ID
'           pTxDate     - Transaction Date (date on which sales occurred)
'           pBataRptType- Enumerated type signifying either 'Total Sales Rpt' or 'Wholesale Sales Rpt'
'
'   FranchiseIDTSG is used rather than FranchiseIDBATA as it's a Primary key and although
'   FranchiseIDBATA is a unique index it allows Nulls which dramatically reduces performance
Const kFldSep As String = ","
Dim lngRecordCount As Long
Dim strRpt As String
Dim strRptHdr As String
Dim strRptDtl As String
Dim strRptTrl As String
Dim strBataFranID As String
Dim strBataRptDate As String
Dim astrLine(1 To 9) As String
Dim ts As Scripting.TextStream
Dim strComText As String
Dim com As ADODB.Command
Dim rst As ADODB.Recordset

'   Set property variables from passed parameters. Properties made available in all events by passing rpt object
'   Both Franchise IDs (Tsg & Bata) required so data required for Name property is available
'   FranName facilitates meaningful UI messages via Events passing rpt object
'   As retrieving snapshot is time consuming the properties are loaded and OnLoad event raised immediately
    m.strFranName = pFranName
    m.lngFranID = pFranID
    m.lngFranID_BATA = pBataFranID
    m.datTxDate = pTxDate
    m.eRptType = pBataRptType
        
    RaiseEvent OnLoad(Me)

    If IsSalesData(pFranID:=pFranID, pTxDate:=pTxDate, pBataRptType:=pBataRptType) Then
''' V386 Start
''' If pTxDate < g.dtmLiveDataStart Then
'''     Select Case pBataRptType   ' Queries use LiveDataArchive table (linked to LiveData table in archive database)
'''         Case eBataTotalSalesRpt: strComText = "qrptBataTotalSalesArchive"
'''             Case eBataWholesaleRpt:  strComText = "qrptBataWSSalesArchive"
'''     End Select
''' Else
''' V386 End
        Select Case pBataRptType   ' Queries use LiveData table in main database
            Case eBataTotalSalesRpt: strComText = "qrptBataTotalSales"
            Case eBataWholesaleRpt:  strComText = "qrptBataWSSales"
        End Select
''' End If  ''' V386
    
    '   Create command and append parameters with appropriate values
        Set com = New ADODB.Command
        With com
            .ActiveConnection = g.cnnDW
            .CommandText = strComText
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter(Name:="FranID", Type:=adInteger, Direction:=adParamInput, Value:=pFranID)
            .Parameters.Append .CreateParameter(Name:="TxDate", Type:=adDate, Direction:=adParamInput, Value:=pTxDate)
        End With
        
    '   Only way to open rst with a command in ADO and specify CursorType and Locking Scheme
    '   even though Default Values are CursorType:=adOpenForwardOnly & LockType:=adLockReadOnly
    '   Also, could perhaps alter GetRst() or write another similar proc for opening a rst
    '   with a command as it would be helpful for trapping any errors in a standardised way
        Set rst = New ADODB.Recordset
        With rst
            .LockType = adLockReadOnly      ' Should by default open as ReadOnly
            .CursorType = adOpenForwardOnly ' Should by default open as ForwardOnly
            .Open Source:=com, Options:=adCmdStoredProc
            If Not (.BOF And .EOF) Then
            '*  m.lngFranID_BATA = rstRpt!FranID_BATA           '-----------------------------------------------------------'
                strBataFranID = m.lngFranID_BATA                ' From 'BATScan Push FTP - Technical Requirements' document '
                strBataRptDate = Format$(pTxDate, "dd/mm/yyyy") '-----------------------------------------------------------'
            '   astrLine(6) = ""                                ' Order code - leave blank ("" = uninitialised value)
                astrLine(7) = strBataFranID                     ' Store Identification Number
                astrLine(8) = "0"                               ' Please make this "0" unless advised otherwise
                astrLine(9) = "0"                               ' Total sales value at actual sell price Ex GST in cents.
                                                                ' (Optional, please leave as "0" if no value is inclueed)
                                                                
                Do While Not .EOF
                    lngRecordCount = lngRecordCount + 1
                    astrLine(1) = rst!Barcode                ' EAN/PLU Code
                    astrLine(2) = rst!Qty                    ' Number of individual transactions
                    astrLine(3) = rst!CostInCents            ' Total cost of sales Inc GST in cents
                    astrLine(4) = rst!TotalInCents           ' Total sales value at actual sell price Inc GST in cents
                    astrLine(5) = rst!NormalSellInCents      ' Current sell price (undiscounted list price) Inc GST
                    strRptDtl = strRptDtl & Join(astrLine, kFldSep) & vbNewLine
                    .MoveNext
                Loop
                
                strRptHdr = "*" & kFldSep & strBataFranID & kFldSep & strBataRptDate & kFldSep & "Day" & kFldSep
            '   Carriage Return (& Line Feed) added to Trailer. Requested by BATA 04Jul2008 (post live implentation)
                strRptTrl = "$" & kFldSep & strBataFranID & kFldSep & strBataRptDate & kFldSep & lngRecordCount & kFldSep & vbNewLine
                strRpt = strRptHdr & vbNewLine & strRptDtl & strRptTrl  ' NB strRptDtl already has trailing vbNewLine
            End If
        End With
        
        rst.Close: Set rst = Nothing
        Set com = Nothing
    End If
    
    m.strRpt = strRpt
    m.bRptHasData = Len(strRpt)
    
    If Me.HasData Then
    '   Save report as a text file
        Set ts = m.fso.CreateTextFile(FileName:=Me.FullName, Overwrite:=True)
        ts.Write m.strRpt
        ts.Close
        Set ts = Nothing
    End If

End Sub

Private Sub Class_Initialize()
    Set m.fso = New Scripting.FileSystemObject
End Sub

Private Sub Class_Terminate()
    If Me.HasData Then
        If m.fso.FileExists(Me.FullName) Then
            m.fso.DeleteFile Me.FullName, Force:=True
        End If
    End If
    Set m.fso = Nothing
End Sub

Public Property Get Name() As String
Dim strSuffix As String

    If Len(m.strName) = 0 Then
        Select Case m.eRptType
            Case eBataTotalSalesRpt: strSuffix = ".txt"
            Case eBataWholesaleRpt:  strSuffix = ".whs"
        End Select
        
    '   BataFranchiseID in filename must be left padded to 3 digits with zeros. (BATA - Samuel Le Hoang 31Aug2007)
        m.strName = "TSPS" & Format(m.lngFranID_BATA, "00#") & "_" & Format$(m.datTxDate, "yyyymmdd") & strSuffix
                  
    End If

    Name = m.strName
    
End Property

Public Property Get FullName() As String
    
    If Len(m.strFullname) = 0 Then
        m.strFullname = m.fso.GetSpecialFolder(TemporaryFolder) & "\" & Me.Name
    End If
    
    FullName = m.strFullname

End Property

Public Property Get FranName() As String
    FranName = m.strFranName
End Property

