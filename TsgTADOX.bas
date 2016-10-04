Attribute VB_Name = "TsgTADOX"
Option Explicit
' ADOX
' Microsoft® ActiveX® Data Objects Extensions for Data Definition Language and Security (ADOX) is an
' extension to the ADO objects and programming model. ADOX includes objects for schema creation and
' modification, as well as security. Because it is an object-based approach to schema manipulation, you can
' write code that will work against various data sources regardless of differences in their native syntaxes.
' Requires reference for ADOX type library:
'  *Microsoft ADO Ext. for DDL and Security: [ADOX library file: Msadox.dll, Program ID (ProgID) is "ADOX]

Public Function AppendColumn(ByRef pCnn As ADODB.Connection, _
                             ByVal pTableName As String, _
                             ByVal pColumnName As String, _
                             ByVal pColumnType As ADODB.DataTypeEnum, _
                    Optional ByVal pColumnSize As Long = 0, _
                    Optional ByVal pPropertyNamesArray As Variant, _
                    Optional ByVal pPropertyValsArray As Variant, _
                    Optional ByRef pErrMsg As String) As Boolean
' Parameters
'-----------
' pCnn
'   Enables access to catalog and provider specific properties
' pTableName
' pColumnType
'   Optional in ADOX.Tables.Append method but made mandatory for this wrapper procedure
' pColumnSize         (Optional)
'   Mapped to 'Defined Size'. Default is zero
' pPropertyNamesArray (Optional)
'   Array of property names for new fields in the new record.
' pPropertyValsArray  (Optional)
'   Array of property values for new fields.
'   Array items should correspond directly to array items in pPropertyNamesArray

' Sample Usage
'-------------
' 1.    If Not AppendColumn(pCnn:=cnn, pTableName:="Promotions", pColumnName:=vntNewFldName) Then
'
' 2.    If Not AppendColumn(pCnn:=cnn, _
'                            pTableName:="Promotions", _
'                            pColumnName:=vntNewFldName, _
'                            pColumnType:=adDate, _
'                            pPropertyNamesArray:=Array("Nullable"), _
'                            pPropertyValsArray:=Array(True)) Then

' Jet Database field properties
'------------------------------
'  0 - AutoIncrement
'  1 - Default
'  2 - Description
'  3 - Nullable
'  4 - Fixed Length
'  5 - Seed
'  6 - Increment
'  7 - Jet OLEDB:Column Validation Text
'  8 - Jet OLEDB:Column Validation Rule
'  9 - Jet OLEDB:IISAM Not Last Column
' 10 - Jet OLEDB: AutoGenerate
' 11 - Jet OLEDB:One BLOB per Page
' 12 - Jet OLEDB:Compressed UNICODE Strings
' 13 - Jet OLEDB:Allow Zero Length
' 14 - Jet OLEDB: Hyperlink
'-------------------------------------------

'   Could call table Append method directly if "Not IsArray(pPropertyNamesArray)"
'   (eg tbl.Columns.Append Item:=pColumnName, Type:=pColumnType, DefinedSize:=pColumnSize)
'   but simpler, neater & not perceptibly slower to use same code for both cases.

Dim lngPrpLoop As Long
Dim strErrMsg As String
Dim col As ADOX.Column
Dim tbl As ADOX.Table
Dim cat As ADOX.Catalog

    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = pCnn
    Set tbl = cat.Tables(pTableName)
    Set col = New ADOX.Column
    With col
        .Name = pColumnName
        .Type = pColumnType
        .DefinedSize = pColumnSize
        If IsArray(pPropertyNamesArray) Then
        '   If pPropertyNamesArray is an array assume matching items in pPropertyValsArray
        '    - incorrect parameter passing should turn up early in testing
            Set .ParentCatalog = cat    ' Enable access to provider specific properties
            For lngPrpLoop = LBound(pPropertyNamesArray, 1) To UBound(pPropertyNamesArray, 1)
                .Properties(pPropertyNamesArray(lngPrpLoop)) = pPropertyValsArray(lngPrpLoop)
            Next lngPrpLoop
        End If
    End With

    On Error Resume Next
        tbl.Columns.Append col
        strErrMsg = VBA.Err.Description ' See GetCn for ways of returning more error info (eg from ADO.Errors collection)
    On Error GoTo 0

    Set col = Nothing
    Set tbl = Nothing
    Set cat.ActiveConnection = Nothing
    Set cat = Nothing

    pErrMsg = strErrMsg
    AppendColumn = (Len(strErrMsg) = 0)

End Function

Public Function IsColumnExist(ByRef pCnn As ADODB.Connection, _
                              ByVal pTblName As String, _
                              ByVal pColName As String) As Boolean
' Could be renamed to IsColumnExist() no trailing s like other fn
Dim bResult As Boolean
Dim tbl As ADOX.Table
Dim cat As ADOX.Catalog

    If Not pCnn Is Nothing Then
        Set cat = New ADOX.Catalog
        Set cat.ActiveConnection = pCnn
        If IsTableExists(pCnn, pTblName) Then
            Set tbl = cat.Tables(pTblName)
            bResult = IsColumnExists(pTbl:=tbl, pColName:=pColName) 'Then
            Set tbl = Nothing
            Set cat = Nothing
        End If
    End If

    IsColumnExist = bResult

End Function

Public Function IsColumnExists(ByRef pTbl As ADOX.Table, ByVal pColName As String) As Boolean
Dim col As ADOX.Column

    On Error Resume Next
        Set col = pTbl.Columns(pColName)
        IsColumnExists = (Err.Number = 0)
    On Error GoTo 0

    Set col = Nothing

End Function

Public Function IsTableExists(ByRef pCnn As ADODB.Connection, ByVal pTblName As String) As Boolean
Dim cat As ADOX.Catalog
Dim tbl As ADOX.Table

    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = pCnn
    On Error Resume Next
        Set tbl = cat.Tables(pTblName)
        IsTableExists = (Err.Number = 0)
    On Error GoTo 0

    Set tbl = Nothing
    Set cat = Nothing

End Function

