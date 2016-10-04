Attribute VB_Name = "FlexGrid"
Option Explicit
' VSFlexGrid Function Module
' Requires reference for VSFlex8Ctl
'---------------------------------------------------------------------------
'   Code Snippet is the type of code you would put in a Form_Load event    |
'   -----------------------------------------------------------            |
'   If .FixedRows = 1 Then .RowHeight(0) = .RowHeight(0) * 1.5             |
'   .MergeCells = flexMergeFixedOnly                                       |
'   ------------------------------------------------------------------------
'   Check out the following properties
''    Me.grd.RowStatus
''    Me.grd.RightCol

Public Function GridGetCollection(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, ByVal pColKey As String, ByVal pSelected As Boolean) As VBA.Collection
Dim lngLoop As Long
Dim strValue As String
Dim colResult As VBA.Collection
    
    Set colResult = New VBA.Collection
    With pGrid
        If pSelected Then
            For lngLoop = 0 To .SelectedRows - 1
                strValue = .TextMatrix(Row:=.SelectedRow(lngLoop), Col:=.ColIndex(pColKey))
                On Error Resume Next    ' Ignore error when adding duplicates: May be called where many rows share same value being collected
                    colResult.Add Item:=strValue, Key:=strValue
                On Error GoTo 0
            Next lngLoop
        Else
            For lngLoop = .FixedRows To .Rows - 1
                strValue = .TextMatrix(Row:=lngLoop, Col:=.ColIndex(pColKey))
                On Error Resume Next    ' Ignore error when adding duplicates: May be called where many rows share same value being collected
                    colResult.Add Item:=strValue, Key:=strValue
                On Error GoTo 0
            Next lngLoop
        End If
    End With
 
    Set GridGetCollection = colResult

End Function

Public Function GridFormatRow(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, _
                              ByVal pRow As Long, _
                              ByVal pProperty As CellPropertySettings, _
                              ByVal pValue)
    With pGrid
        .Cell(pProperty, pRow, .FixedCols, , .Cols - 1) = pValue
    End With
    
End Function

Public Function GridIsDataSelected(ByRef pGrid As VSFlex8Ctl.VSFlexGrid) As Boolean
' Returns True if data row is selected and False otherwise (*Use when Selection Mode = flexSelectionListBox*)
' Is a workaround for bug in VSFlexGrid V8 (Not 7) where default property values indicate FixedRows are selected.
' eg Default values for a grid with 1 fixed row are: SelectedRows = 1; SelectedRow(0) = 0
'    Prior to V8 you could use grd.SelectedRows > 0 to determine if any rows were selected

Dim lngRow As Long

    With pGrid
    '   Unselect any selected fixed rows
        For lngRow = 0 To .FixedRows - 1
            .IsSelected(lngRow) = False
        Next lngRow
        
    '   SelectedRows: returns the number of selected rows when SelectionMode is set to flexSelectionListBox.
        GridIsDataSelected = (pGrid.SelectedRows > 0)
    End With

End Function

Public Sub GridAutoSizeCol(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, _
                           ByVal pCol As Long, _
                  Optional ByVal pHideWhenEmpty As Boolean = False)
'   Uses AutoSizeMode which will not set ColWidth outside the
'   range set by the ColWidthMin and ColWidthMax properties

Dim bIsData As Boolean
Dim lngRow As Long
Dim OriginalAutoSizeMode As AutoSizeSettings

    With pGrid
        If (.Rows > 0) And (-1 < pCol < .Cols) Then
        '   Check whether pCol has data and set bIsData flag accordingly
            For lngRow = .FixedRows To .Rows - 1
                If LenB(Trim$(.TextMatrix(lngRow, pCol))) <> 0 Then
                    bIsData = True
                    Exit For
                End If
            Next lngRow
            
            .ColHidden(pCol) = pHideWhenEmpty And (Not bIsData)
            
            If Not pGrid.ColHidden(pCol) Then
            '   Save original AutoSizeMode
                OriginalAutoSizeMode = .AutoSizeMode
                
            '   1. AutoSize column by ColWidth ( ColWidthMin <= ColWidth <= ColWidthMax)
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize pCol  ' Will Autosize resize > ColWidthMax? (It appears not)
                
            '   2. AutoSize column by RowHeight
            '   WordWrap set True but not restored b/c if was False resotring
            '   will leaves cells resized vertically but with unwrapped text.
                .WordWrap = True
                .AutoSizeMode = flexAutoSizeRowHeight
                .AutoSize pCol
                
                .AutoSizeMode = OriginalAutoSizeMode
            End If
        End If
    End With

End Sub

Public Sub GridSetToolTip(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, _
                 Optional ByVal pTextFromCellRow As Long = -1, _
                 Optional ByVal pTextFromCellCol As Long = -1, _
                 Optional ByVal pFixedRowTooltip As String = "")
' Tonda Perhaps have a pTooltipText override which if provided sets the grids ToolTipText
'   regardless of all other parameters. You could use the grid's ToolTipText property
'   directly for this, but the counter argument is that if you always use this function
'   you only need to search only for the function to find how you are setting grid tooltips
Dim lngFixedRow As Long
Dim strToolTipText As String

    With pGrid
        Select Case .MouseRow
            Case -1                             ' Over blank space (below fixed rows and data in grid)
                .ToolTipText = vbNullString
            Case Is < .FixedRows                ' Over fixed row
                If LenB(pFixedRowTooltip) <> 0 Then
                    .ToolTipText = pFixedRowTooltip
                Else
                '   Dynamically created tooltip for fixed rows (Sort by: Col Heading)
                '   If Not over blank space (rhs of grid area)
                    If .MouseCol > -1 Then               ' Not over blank space (rhs of grid area)
                        If Not .MergeRow(.MouseRow) Then ' Not over a merged row: in-built sorting only works on un-merged col headings
                            strToolTipText = "Sort by:"
                            For lngFixedRow = 0 To .FixedRows - 1
                                If (lngFixedRow = 0) Or Not .MergeCol(.MouseCol) Then
                                    strToolTipText = strToolTipText & " " & .TextMatrix(Row:=lngFixedRow, Col:=.MouseCol)
                                Else
                                    If .TextMatrix(Row:=lngFixedRow - 1, Col:=.MouseCol) <> .TextMatrix(Row:=lngFixedRow, Col:=.MouseCol) Then
                                        strToolTipText = strToolTipText & " " & .TextMatrix(Row:=lngFixedRow, Col:=.MouseCol)
                                    End If
                                End If
                            Next
                        End If
                    End If
                    .ToolTipText = strToolTipText ' Outside If stmnt to clear Tooltip text when moving over blank space
                End If
                
            Case Else                           ' Over data row
                If (pTextFromCellRow = -1) Or (pTextFromCellCol = -1) Then
                    .ToolTipText = vbNullString
                Else
                    .ToolTipText = .TextMatrix(pTextFromCellRow, pTextFromCellCol)
                End If
        End Select
        
    End With

End Sub

Public Function GridFormatEditWindow(ByRef pGrid As VSFlex8Ctl.VSFlexGrid)
'   The EditWindow takes it's formatting from the formatting for the underlying cell
'   My validations effect cell formatting, and I use this function to force the EditWindow to take
'   the formatting of the underlying cell while the edit is in progress. (ie between keystrokes)
Dim lngSelTextStart As Long

    With pGrid
        lngSelTextStart = .EditSelStart
        .EditText = .EditText
        .EditSelStart = lngSelTextStart
    End With

End Function

Public Function GridClearSelections(ByRef pGrid As VSFlex8Ctl.VSFlexGrid)
Dim lngRowLoop As Long

    With pGrid
        If .SelectionMode = flexSelectionListBox Then
            For lngRowLoop = 0 To .SelectedRows - 1
                .IsSelected(.SelectedRow(lngRowLoop)) = False
            Next lngRowLoop
        End If
    End With
    
End Function

Public Function GridSelChangeAssociatedRows(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, ByVal pAssociateByColWithKey As String) As Boolean
'   Apply SelChange in current row to rows with same value in column(pGroupColKey)
'   Returns True on completion of primary call (as distinct from assumed cascading event calls)
'   (Handy in calling to prevent un-necessarily repeating code execution from cascading events)

Static sbIgnoreCascadingEvent As Boolean ' To ignore cascading events triggered by setting IsSelected(n)
Dim bSelect As Boolean
Dim bPreviouslyEnabled As Boolean
Dim lngLoop As Long
Dim lngGroupColIdx As Long
Dim strGroupColVal As String
        
    If Not sbIgnoreCascadingEvent Then
        sbIgnoreCascadingEvent = True
        bPreviouslyEnabled = SetCtlEnabled(pGrid, False) ' Prevent another click/selection
        
        With pGrid
            bSelect = .IsSelected(.Row)
            lngGroupColIdx = .ColIndex(pAssociateByColWithKey)
            strGroupColVal = .TextMatrix(.Row, lngGroupColIdx)
            For lngLoop = .FixedRows To .Rows - 1   ' Allow for heading row(s)
                If strGroupColVal = .TextMatrix(lngLoop, lngGroupColIdx) Then
                'NB Setting .IsSelected triggers pGrid_SelChange event which is probably the calling event
                '   Calling code should handle the prolbem of cascading deletes
                    .IsSelected(lngLoop) = bSelect  ' Triggers pGrid_SelChange event (probable calling event)
                End If
            Next
        End With
        
        SetCtlEnabled pGrid, bPreviouslyEnabled         ' Restore original Enabled value
        sbIgnoreCascadingEvent = False
    End If
    
    GridSelChangeAssociatedRows = (Not sbIgnoreCascadingEvent)
  
End Function

Public Function EnbableMoveRowsDown(ByRef pGrid As VSFlex8Ctl.VSFlexGrid) As Boolean
' Returns True if there is space for at least one of the selected rows to be move DOWN
Dim bEnable As Boolean
Dim lngRow As Long

    With pGrid
        If pGrid.SelectedRows Then
            For lngRow = .SelectedRow(0) To .Rows - 1
                If Not .IsSelected(lngRow) Then
                    bEnable = True
                    Exit For
                End If
                
            Next lngRow
        End If
    End With

    EnbableMoveRowsDown = bEnable

End Function

Public Function EnbableMoveRowsUp(ByRef pGrid As VSFlex8Ctl.VSFlexGrid) As Boolean
' Returns True if there is space for at least one of the selected rows to be move UP
Dim bEnable As Boolean
Dim lngRow As Long

    With pGrid
        If pGrid.SelectedRows Then
            For lngRow = .SelectedRow(.SelectedRows - 1) To .FixedRows Step -1
                If Not .IsSelected(lngRow) Then
                    bEnable = True
                    Exit For
                End If
                
            Next lngRow
        End If
    End With

    EnbableMoveRowsUp = bEnable

End Function

Public Sub GridAutoSizeCols(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, Optional ByVal pHideWhenEmpty As Boolean = False)
''' Under development - originally being developed with intended use in TsgMsgCentre
''   In majority of cases can be replaced by:- Grid.AutoSize Col1 := 0, col2 := Grid.Cols - 1
'Dim lngLoop As Long
'
'    For lngLoop = 0 To pGrid.Cols - 1
'        If Not pGrid.ColHidden(lngLoop) Then
'            GridAutoSizeCol pGrid:=pGrid, pCol:=lngLoop
'        End If
'    Next lngLoop
'
End Sub

'Public Function GridBoundTotal(ByRef pGrid As VSFlex8Ctl.VSFlexGrid, ByVal pCol As Long) As Double
'' Aggregate method of VSFlexGrid will perform the same function (Probably quicker with less maintenance)
'Dim lngRow As Long
'Dim dblResult As Double
'
'    With pGrid
'        For lngRow = .FixedRows To .Rows - 1
'            dblResult = dblResult + Cn(.ValueMatrix(lngRow, pCol), 0)
'        Next lngRow
'    End With
'
'    GridBoundTotal = dblResult
'
'End Function

'Private Sub FillFlexGridComboList(ByRef pGrid As VSFlex7Ctl.VSFlexGrid, _
'                                  ByVal pGridColumn As Long, _
'                                  ByVal pRst As ADODB.Recordset, _
'                                  ByVal pDisplayField As String, _
'                         Optional ByVal pKeyField As String = "", _
'                         Optional ByVal pIncludeBlank As Boolean = False, _
'                         Optional pLimitToList As Boolean = True)
''   Single drop downs can act as either simple drop down lists or column combo boxes
''   (ie. Can be limit to list or not) This does not apply to multi column combo
''   boxes because the key field would not be know (as per MS Access limit to list problem)
'Dim strColComboList As String
'
'    With pGrid
'        If Len(pKeyField) = 0 Then
'            strColComboList = .BuildComboList(pRst, pDisplayField)
'            If pIncludeBlank Then
'                If Len(strColComboList) = 0 Then
'                    strColComboList = " |"
'                Else
'                    strColComboList = " |" & strColComboList
'                End If
'            End If
'        '   pLimitToList = False can only apply here.
'        '   Cannot apply when there are different display and key fields (LimitToList problem)
'            If Not pLimitToList Then
'                strColComboList = "|" & strColComboList
'            End If
'        Else
'            strColComboList = .BuildComboList(rstLookup, pDisplayField, pKeyField)
'            If pIncludeBlank Then
'                If Len(strColComboList) = 0 Then
'                    strColComboList = "#-1; "
'                Else
'                    strColComboList = "#-1; |" & strColComboList
'                End If
'            End If
'        End If
'
'        .ColComboList(pGridColumn) = strColComboList
'
'    End With
'
'End Sub
