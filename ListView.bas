Attribute VB_Name = "ListView"
Option Explicit
' ListView Function Module
' (ListView ActiveX control is part of 'Microsoft Windows Common Controls 5.0' and must be added to Project Components)
'
' Distribution Note:
'   The ListView control is part of a group of ActiveX controls that are found in the COMCTL32.OCX file.
'   To use the ListView control in your application, you must add the COMCTL32.OCX file to the project.
'   When distributing your application, install the COMCTL32.OCX file in the user's Microsoft Windows System or System32 directory
'
'   There are two versions of the Microsoft Windows Common Controls.
'   Comctl32.ocx contains Windows Common Controls 5.0 and was included with Microsoft Visual Studio 5.0.
'   Mscomctl.ocx contains Windows Common Controls 6.0 and was included with Visual Studio 6.0.

Public Function LvwGetItemCollection(ByRef pListView As ComctlLib.ListView, _
                            Optional ByVal pSelected As Boolean = False) As VBA.Collection
Dim colResult As VBA.Collection
Dim lvwItem As ListItem
         
    Set colResult = New VBA.Collection
    
    For Each lvwItem In pListView.ListItems
        If (Not pSelected) Or (pSelected And lvwItem.Selected) Then
            colResult.Add Item:=lvwItem
        End If
    Next lvwItem
    
    Set LvwGetItemCollection = colResult

End Function

Public Function LvwGetSubItemCollection(ByRef pListView As ComctlLib.ListView, _
                                        ByVal pSubItemIdx As Long, _
                               Optional ByVal pSelected As Boolean = False) As VBA.Collection
Dim strValue As String
Dim colResult As VBA.Collection
Dim lvwItem As ComctlLib.ListItem
    
    Set colResult = New VBA.Collection
    With pListView
        
        For Each lvwItem In pListView.ListItems
            If (Not pSelected) Or (pSelected And lvwItem.Selected) Then
                strValue = CStr(lvwItem.SubItems(pSubItemIdx))
                On Error Resume Next    ' Ignore error when adding duplicates: May be called where many rows share same value being collected
                    colResult.Add Item:=strValue, Key:=strValue
                On Error GoTo 0
            End If
        Next lvwItem
        
        If colResult.Count = 0 Then
            Set colResult = Nothing
        End If
    
    End With
    
    Set LvwGetSubItemCollection = colResult

End Function

Public Function LvwIsDataSelected(ByRef pListView As ComctlLib.ListView) As Boolean
Dim bResult As Boolean
Dim lvwItem As ListItem
   
    For Each lvwItem In pListView.ListItems
        If lvwItem.Selected Then
            bResult = True
            Exit For
        End If
    Next lvwItem
    
    LvwIsDataSelected = bResult

End Function

Public Sub LvwSelChangeAssociatedRows(ByRef pListView As ComctlLib.ListView, ByVal pAssociateBySubItemWithIdx As Long)
'   Apply selection in current row to rows with same value in SubItem(pAssociateBySubItemWithIdx)
Dim bSelect As Boolean
Dim bPreviouslyEnabled As Boolean
Dim strSubItem As String
Dim lvwItem As ListItem
    
    bPreviouslyEnabled = SetCtlEnabled(pListView, False)    ' Prevent another click/selection
    
    With pListView
        bSelect = .SelectedItem.Selected
        strSubItem = .SelectedItem.SubItems(pAssociateBySubItemWithIdx)
        For Each lvwItem In pListView.ListItems
            If (lvwItem.SubItems(pAssociateBySubItemWithIdx) = strSubItem) Then
                lvwItem.Selected = bSelect
            End If
        Next lvwItem
    End With
    
    SetCtlEnabled pListView, bPreviouslyEnabled             ' Restore original Enabled value

End Sub
