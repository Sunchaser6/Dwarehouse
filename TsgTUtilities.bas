Attribute VB_Name = "TsgTUtilities"
Option Explicit

' Utility Module
' Requires references for:
'  *Microsoft Scripting Runtime

Public Sub AppendTextFile(ByVal pFullName As String, ByVal pText As String)
'   Could be an idea to reverse the order of naming to TextFileAppend, TextFileSave, TextFile...
' FROM Application module of TslRental
Dim ts As Scripting.TextStream
Dim fso As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject

    If Not fso.FileExists(pFullName) Then
        SaveTextFile pFilename:=pFullName, pFileText:=pText & vbNewLine
    Else
        Set ts = fso.OpenTextFile(FileName:=pFullName, IOMode:=ForAppending)
        ts.WriteLine pText
        ts.Close
        Set ts = Nothing
    End If

    Set fso = Nothing

End Sub

Public Function BoolToChkBox(ByVal pBooleanValue As Boolean) As Integer
'   One of two functions (ChkBoxToBool and BoolToChkBox)
'   for moving values between a database boolean value and a checkbox control
'   CheckBox control — 0 is Unchecked (default), 1 is Checked, and 2 is Grayed (dimmed).
    
    If pBooleanValue Then
        BoolToChkBox = 1
    Else
        BoolToChkBox = 0      ' Not giving option of greyed (2)
    End If

End Function

Public Function Bracket(ByRef pString As String) As String
'   Args passed ByRef for speed, small proc can avoid side effects with care
    Bracket = "(" & pString & ")"
End Function

Public Function ChkBoxToBool(ByRef pCheckbox As CheckBox) As Boolean
'   One of two functions (ChkBoxToBool and BoolToChkBox)
'   for moving values between a database boolean value and a checkbox control
'   CheckBox control — 0 is Unchecked (default), 1 is Checked, and 2 is Grayed (dimmed).

    If pCheckbox.Value = vbChecked Then
        ChkBoxToBool = True
    Else
        ChkBoxToBool = False
    End If
    
End Function

Public Function Cn(ByRef pValue As Variant, ByRef pReplaceWith As Variant) As Variant
'   Args passed ByRef for speed, small proc can avoid side effects with care

    If IsNull(pValue) Then
        Cn = pReplaceWith
    Else
        Cn = pValue
    End If

End Function

Public Function CnvDdMmmYyToDate(ByVal pDdMmmYy As String) As Date
'   Function presumes pDdMmmYy is a valid date
'   IsValidDate_ddmmmyy() should be prior to calling this fn to the value being passed
Dim intYear As Integer
Dim strDay As String
Dim strMonth As String
Dim strYear As String

    strDay = Left$(pDdMmmYy, Length:=2)
    strMonth = Mid$(pDdMmmYy, start:=3, Length:=3)
    strYear = Right$(pDdMmmYy, 2)

'   As per DateSerial Function - For the year argument, values between 0 and 29, inclusive, are
'   interpreted as the years 2000–2029. Values between 30 and 99 are interpreted as the years 1930–1999.
    intYear = CInt(strYear)
    If intYear < 30 Then
        intYear = 2000 + intYear
    Else
        intYear = 1900 + intYear
    End If
    
    CnvDdMmmYyToDate = DateValue(strDay & " " & strMonth & " " & intYear)

End Function

Public Sub CtlMove(ByVal bForward As Boolean, ByRef ctlMoveFrom As Control)
'--------------------------------------------------------------------------------------'
' Sets focus to the next/prev ctl in the Tab order                                     '
' If next/prev ctl can't receive focus, calls itself with next/prev ctl as current ctl '
' A static counter (sintNoRecursions) is maintained to limit the number of recursions  '
' Max recursions is set in constant kMaxNumRecursions                                  '
'--------------------------------------------------------------------------------------'
' Modified AUrban 23Jan06                                                              '
'   Passed parameters explicilty ByVal/ByRef                                           '
'   Removed Call Statements                                                            '
'   Cleard ctlMoveTo and f variables                                                   '
' Modified AUrban 31Aug03                                                              '
'   ctlMoveFrom passed as a control                                                    '
' Modified AUrban 08Jan03                                                              '
'   Tidy up. Replace data type suffixes in Dim statements, change constant prefix.     '
' Modified AUrban 28Jul00                                                              '
'   Don't move to controls that have 'TabStop = False'                                 '
'   Note that for readability (and perhaps speed) the Do while Loop could be replaced  '
'   by a less technically correct For Each Loop with an Exit For                       '
' New Procedure AUrban 22Jun98                                                         '
'--------------------------------------------------------------------------------------'
Const kMaxNumRecursions As Integer = 20
Static sintNoRecursions As Integer
Dim f As Form, ctlMoveTo As Control
Dim bTabStop As Boolean, bFound As Boolean, bError As Integer
Dim intCnt As Integer, intNumCtls As Integer
Dim intCurrentTabIdx As Integer, intLoopTabIdx As Integer, intMaxTabIdx As Integer, intMoveToTabIdx As Integer
Dim intMoveToCtlIdx As Integer, intMinTabCtlIdx As Integer, intMaxTabCtlIdx As Integer
 
    If sintNoRecursions < kMaxNumRecursions Then
    '   Proceed with normal processing
        Set f = ctlMoveFrom.Parent
        intNumCtls = f.Count
        intCurrentTabIdx = ctlMoveFrom.TabIndex
        If bForward Then
        '   won't be found if ctlMoveFrom is LAST in Tab Index
            intMoveToTabIdx = intCurrentTabIdx + 1
        Else
        '   won't be found if ctlMoveFrom is FIRST in Tab Index
            intMoveToTabIdx = intCurrentTabIdx - 1
        End If
        
        Do
            On Error Resume Next    ' For ctls without Tab Index property
                intLoopTabIdx = f(intCnt).TabIndex
                bError = (Err <> 0)
            On Error GoTo 0         ' Reset Error Handling
            
            If Not bError Then
                If intLoopTabIdx = intMoveToTabIdx Then
                    bFound = True
                    intMoveToCtlIdx = intCnt
                '   If Not bFound, will only need to find one of intMaxTabCtlIdx
                '   or intMinTabCtlIdx therefore could be optimised
                ElseIf intLoopTabIdx > intMaxTabIdx Then
                    intMaxTabIdx = intLoopTabIdx
                    intMaxTabCtlIdx = intCnt
                ElseIf intLoopTabIdx = 0 Then
                    intMinTabCtlIdx = intCnt
                End If
            End If
            intCnt = intCnt + 1
        Loop Until bFound Or (intCnt = intNumCtls)
        
        If bFound Then
            Set ctlMoveTo = f(intMoveToCtlIdx)
        ElseIf bForward Then
            Set ctlMoveTo = f(intMinTabCtlIdx)
        Else
            Set ctlMoveTo = f(intMaxTabCtlIdx)
        End If
        
    '   May not want next/prev control to accept focus
    '   for example it's TabStop property may be false
    '   or it does not have a TabStop property
        On Error Resume Next        ' For ctls without TabStop property
            bTabStop = ctlMoveTo.TabStop
            bError = (Err <> 0)
        On Error GoTo 0             ' Reset Error Handler
 
        If bError Or Not bTabStop Then
            sintNoRecursions = sintNoRecursions + 1
            CtlMove bForward, ctlMoveTo
        Else
        '   next/prev control may be unable to accept focus (eg. disabled)
        '   therefore move on to next/prev control able to accept focus
            On Error Resume Next        ' For ctls that can't accept focus
                ctlMoveTo.SetFocus
                bError = (Err <> 0)
            On Error GoTo 0             ' Reset Error Handler
            If bError Then
                sintNoRecursions = sintNoRecursions + 1
                CtlMove bForward, ctlMoveTo
            End If
        End If
        Set ctlMoveTo = Nothing
        Set f = Nothing
        
    '   Reset recursion counter
        sintNoRecursions = 0
    End If
 
End Sub

Public Sub CtlMoveNext(Optional ByRef ctlMoveFrom As Control = Nothing)
    If ctlMoveFrom Is Nothing Then
        Set ctlMoveFrom = Screen.ActiveControl
    End If
    CtlMove True, ctlMoveFrom
End Sub

Public Sub CtlMovePrevious(Optional ByRef ctlMoveFrom As Control = Nothing)
    If ctlMoveFrom Is Nothing Then
        Set ctlMoveFrom = Screen.ActiveControl
    End If
    CtlMove False, ctlMoveFrom
End Sub

Public Function Czls(ByRef pVar As String, ByRef pSubs As Variant) As Variant
'   Args passed ByRef for speed, small proc can avoid side effects with care
'   Czls = (ConvertZeroLengthString)
    
    If Len(pVar) = 0 Then
        Czls = pSubs
    Else
        Czls = pVar
    End If

End Function

Public Function DQ(ByRef pVnt As Variant) As String
'   Args passed ByRef for speed, small proc can avoid side effects with care
    DQ = """" & pVnt & """" ' Double Quote
End Function

Public Function FileExists(ByVal pFullName As String) As Boolean
    FileExists = (Len(Dir$(pFullName, vbNormal)) <> 0)
End Function

Public Function GetTitleBar(Optional ByVal pFormSubject As String = "") As String
' Same prefix across all form titles in App so that AppActivate sets focus to App.
' Originally prefixed with a space so AppActivate doesn't activate an Explorer
' instance using a folder of same name (Name in TitleBar) etc., but would then need to
' supply the prefixed title to all MsgBoxs to be absolutely correct.
' CONCLUSION WAS TO INSTALL MY APPLICATIONS IN TO DIRECTORIES WITH A DIFFERENT
' NAME THAN THE EXE IF ONLY BY A TRAILING ALT(255)/CHR$(255) WONKY SPACE CHARACTER.

'   NB. Using caption of an MDI form sets the focus TO
'   the application regardless of which forms it has open
    
''   ----------------------------
''   Behaviour of VBA.AppActivate (from testing with an MDI form)
''   ----------------------------
''   A form named "Form1" can be activated by VBA.AppActivate "Form1" regardless
''   of its caption if its caption has not been altered at run-time
    
    
    If LenB(pFormSubject) = 0 Then
        GetTitleBar = " " & App.Title
    Else
        GetTitleBar = " " & App.Title & " - " & pFormSubject
    End If

End Function

Public Function GetValStringValue(ByVal pVString As String, ByVal pVName As String) As String
Dim lngLoop As Long
Dim lngPosOfEqual As Long
Dim strLoop As String
Dim strResult As String
Dim astr() As String

    astr = Split(pVString, ";")
    For lngLoop = LBound(astr) To UBound(astr)
        strLoop = astr(lngLoop)
        lngPosOfEqual = InStr(1, strLoop, "=", vbBinaryCompare)
        If lngPosOfEqual Then
            If StrComp(Trim$(Left$(strLoop, lngPosOfEqual - 1)), pVName, vbTextCompare) = 0 Then
                strResult = Trim$(Right$(strLoop, Len(strLoop) - lngPosOfEqual))
                Exit For
            End If
        End If
    Next
    
    GetValStringValue = strResult

End Function

Public Function GetVersionString() As String
' Use GetAbsolutePathName to get case of file name as seen in Explorer
Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    GetVersionString = "Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & _
                       "Program file: " & fso.GetAbsolutePathName(App.Path & "\" & App.EXEName & ".exe")
    Set fso = Nothing

End Function

Public Function GetWcListFromColn(ByRef pCollection As VBA.Collection, _
                                  ByVal pSqlQuoteVals As Boolean) As String
''' pSqlQuoteVals may later be given a better name
'''                           Optional byval pIsQuoteValues as Boolean = False ) As String

' Returns an empty string if pCollection is empty -> calling code can test for empty string.
' A SQL statement with an empty Value List in Where Clause {eg. SELECT ... WHERE id IN ()}
' is invalid SQL syntax.

Const kSep As String = ", "
Dim strList As String
Dim vntItem As Variant

    For Each vntItem In pCollection
        If pSqlQuoteVals Then
            strList = strList & SqlQ(vntItem) & kSep
        Else
            strList = strList & vntItem & kSep
        End If
    Next vntItem
    
    If Len(strList) Then
        strList = Left$(strList, Len(strList) - Len(kSep))
        strList = Bracket(strList)
    End If

    GetWcListFromColn = strList
    
End Function

Public Function GetWcListItemCount(ByVal pWcValueList As String) As Long
'  Function assumes correctly formatted WcValueList i.e. "(a, b, c)"
Dim lngResult As Long
Dim strTemp As String
Dim astrTemp() As String

    strTemp = Trim$(pWcValueList)
    If (InStr(strTemp, "(") = 1) And (InStrRev(strTemp, ")") = Len(strTemp)) Then
        strTemp = Right$(strTemp, Len(strTemp) - 1)
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    '   Split() - Returns a zero-based, one-dimensionalarray containing a specified number of substrings.
        astrTemp() = Split(strTemp, ",")    ' split returns a zero based one dimesional
        lngResult = UBound(astrTemp) - LBound(astrTemp) + 1
    End If
    
    GetWcListItemCount = lngResult

End Function

Public Function IsEachCharADigit(ByVal pString As String) As Boolean
Dim bResult As Boolean
Dim lngPos As Long

    If Len(pString) Then
        bResult = True
    End If
    
    For lngPos = 1 To Len(pString)
        If Not IsNumeric(Mid$(pString, lngPos, 1)) Then
            bResult = False
            Exit For
        End If
    Next lngPos

    IsEachCharADigit = bResult
    
End Function

Public Function IsEmptyArray(pArray As Variant) As Boolean
'   Assumptions:    Assumes that pArray is an array
'   pArray is defined as a variant so that Arrays of different types can be passed
Dim lngLowerBound As Long

    If Not IsArray(pArray) Then
    '   Not an array variable
        Err.Clear
        Err.Raise Number:=513, _
                  Source:="User defined Function: IsEmptyArray", _
                  Description:="Arguement was not an array variable."
    Else
        On Error Resume Next
            lngLowerBound = LBound(pArray, 1)
            IsEmptyArray = (Err <> 0)
        On Error GoTo 0
    '   There are cases when an empty array variable returns a value
    '   from LBound and UBound functions without generating an error.
    '   I have observed this when interrogating (1) return values from the
    '   split function or (2) ParamArrays when no parameters are passed.
    '   I have only observed it with one dimensional arrays.
    '   In each case the UBound is less than the LBound.

    '   This can be reproduced in the immediate window with:
    '   (1) ?LBound(split("","A")) -> 0
    '   (2) ?UBound(split("","A")) -> -1
        If Not IsEmptyArray Then
            IsEmptyArray = (UBound(pArray) < lngLowerBound)
        End If
    End If

End Function

Public Function IsIDE() As Boolean
    On Error Resume Next
        Debug.Print 1 / 0
        IsIDE = (Err.Number <> 0)
    On Error GoTo 0 ' Clear Err object (Could have used Err.Clear)
End Function

Public Function IsIPAddress(ByVal zvIPAddressID As String) As Boolean
' Validating IP Addresses by Miles Forrest - Independent Developer
' (Australian Visual Developers Forum - Aug 98)
'This little routine has only one purpose in life - to check whether an IP Address is valid.
'It checks two things - that the address is in the form xxx.xxx.xxx.xxx and that each xxx is between 1 and 3 numeric digits, and its numeric value is in the range 0-255.
'It looks a bit hairy (and is a bit hairy!) but it's fast - which was important because when I first developed it I had to process thousands of IP Addresses.

Dim ipLengthOfGroup As Integer
Dim ipPosition As Integer
Dim ipNoOfGroups As Integer

    IsIPAddress = False
    ipNoOfGroups = 0
    ipLengthOfGroup = 0
    For ipPosition = 1 To Len(zvIPAddressID)
        Select Case Asc(Mid$(zvIPAddressID, ipPosition, 1))
            Case 48 To 57
                ipLengthOfGroup = ipLengthOfGroup + 1
                If ipLengthOfGroup = 4 Then
                    Exit Function
                End If
            Case 46
                ipNoOfGroups = ipNoOfGroups + 1
                If ((ipNoOfGroups = 4) Or (ipLengthOfGroup = 0)) Or _
                    (Val(Mid$(zvIPAddressID, ipPosition - ipLengthOfGroup, ipLengthOfGroup)) > 256) Then
                    Exit Function
                End If
                ipLengthOfGroup = 0
            Case Else
                Exit Function
        End Select
    Next ipPosition

    IsIPAddress = (ipNoOfGroups = 3) And (ipLengthOfGroup > 0) And _
                  (Val(Mid$(zvIPAddressID, ipPosition - ipLengthOfGroup, ipLengthOfGroup)) < 256)
    
End Function

Public Function IsKeyInCollection(ByRef pCollection As VBA.Collection, ByVal pKey As String) As Boolean
'   Note we don't assign the collection member to a variable as it may be an object,
'   (requiring "Set varname = ...") or a more simple type held in a variant (requiring "varname = ...")
Dim bResult As Boolean
Dim strTypeName As String

    On Error Resume Next
        strTypeName = TypeName(pCollection(Index:=pKey))
        bResult = (Err.Number = 0)
    On Error GoTo 0

    IsKeyInCollection = bResult

End Function

Public Function ListBoxGetCollection(ByRef pListBox As VB.ListBox, ByVal pItemData As Boolean, ByVal pSelected As Boolean) As VBA.Collection
Dim bAdd As Boolean
Dim lngLoop As Long
Dim colResult As VBA.Collection
    
    Set colResult = New VBA.Collection
    With pListBox
        If (Not pSelected) Or (pSelected And .SelCount) Then
            For lngLoop = 0 To .ListCount - 1
                bAdd = (Not pSelected) Or (pSelected And .Selected(lngLoop))
                If bAdd Then
                    If pItemData Then
                    '   Ignore error when adding duplicates: May be called where many rows share same value being collected
                        On Error Resume Next
                            colResult.Add Item:=.ItemData(lngLoop), Key:=CStr(.ItemData(lngLoop))
                        On Error GoTo 0
                    Else
                    '   Ignore error when adding duplicates: May be called where many rows share same value being collected
                        On Error Resume Next
                            colResult.Add Item:=.List(lngLoop), Key:=.List(lngLoop)
                        On Error GoTo 0
                    End If
                End If
            Next lngLoop
        End If
    End With
 
    Set ListBoxGetCollection = colResult

End Function

Function Plural(ByVal pQty As Variant, ByVal pNounSingular As String) As String
''' REQUIRES AN OPTIONAL pIncludeQty PARAMETER, AND OPTIONAL pSuffix PARAMETER SO TAHT A NEWLINE
''' OR SPACE CAN BE APPENEDED PROBLEM WITH THIS IS THAT IT WOULD ONLY BE USEFUL IF IT WAS CONDITIONAL
''' ON, ALSO PERAHAPS THE pIncludeQty SHOULD REALLY BE INCLUDE_ZERO_QTY OR INCLUDE_QTY_OF_ONE

'   Fractional amounts and zero amounts are given the plural version of a noun
'   eg I shot 0 foxes, an average of 1.3 foxes a year
'   Note that pQty is a variant so it can take all of the numeric data types
'   It should probably include a test that the pQty parameter passes IsNumeric()
Dim strAdjustedNoun As String

    If (Abs(pQty) = 1) Then
        strAdjustedNoun = pNounSingular
    Else
        Select Case pNounSingular
        '   Case Exceptions
        '       sheep -> sheep
        '       ox -> oxen
        '   Case Rule
        '       ends in ox  -> oxes (BUT BEWARE OX -> OXEN)
        '       ends in eaf -> eaves
        '       ends in ily -> ilies (eg family -> families, homily -> homilies)
        '   Case Else like record where an s is added.
            Case Else
                strAdjustedNoun = pNounSingular & "s"
        End Select
    End If
    
    Plural = pQty & " " & strAdjustedNoun

End Function

Public Sub SaveTextFile(ByVal pFilename As String, ByRef pFileText As String, Optional pOverwrite As Boolean = True)
'   No error trapping. Simply a pass through function.
'   If the overwrite argument is False for a filename that already exists an error occurs.

'   It is up to the calling code to handle errors
'   (eg. File exists and you are passing pOverwrite = False)
'   (Calling code will determine whether to prompt for a new filename etc.)
Dim fso As Scripting.FileSystemObject
Dim ts As Scripting.TextStream

    Set fso = New Scripting.FileSystemObject
    Set ts = fso.CreateTextFile(FileName:=pFilename, Overwrite:=pOverwrite)
    ts.Write pFileText
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

End Sub

Public Function SetCtlEnabled(ByRef pCtl As Control, ByVal pEnabled As Boolean) As Boolean
' Purpose: Set Ctl.Enabled to pEnabled
' Returns: Value of Ctl.Enabled at the time of calling the function.
'          Return value can be stored and used to restore the original value.
'
' Example Usage-
'   bPreviouslyEnabled = SetCtlEnabled(Ctl, False)
'   .
'   .
'   Ctl.Enabled = bPreviouslyEnabled

    SetCtlEnabled = pCtl.Enabled
    If pEnabled <> pCtl.Enabled Then
        pCtl.Enabled = pEnabled
    End If
    
End Function

Public Function SetFormEnabled(ByRef pForm As Form, ByVal pEnabled As Boolean) As Boolean
' Purpose: Set pForm.Enabled property to pEnabled
' Returns: Value of pForm.Enabled at the time of calling the function.
'          Return value can be stored and used to restore the original value
'
'   bPreviousMeEnabled = SetFormEnabled(Me, False)
'   .
'   .
'   SetFormEnabled Me, bPreviousMeEnabled

    SetFormEnabled = pForm.Enabled
    If pForm.Enabled <> pEnabled Then
        pForm.Enabled = pEnabled
    End If
    
End Function

Public Function SetMousePointer(ByVal pMousePointer As Integer) As Integer
' Purpose: Set Screen.MousePointer to pMousePointer
' Returns: Value of Screen.MousePointer at the time of calling the function.
'          Return value can be stored and used to restore the original value.
'
'   The Screen object is the entire Windows desktop. Using the Screen object, you can set the
'   MousePointer property of the Screen object to the hourglass pointer while a modal form is displayed.
'
' Example Usage-
'   intPreviousMousePointer = SetMousePointer(vbHourglass)
'   .
'   .
'   Screen.MousePointer = intPreviousMousePointer

    SetMousePointer = Screen.MousePointer
    If pMousePointer <> Screen.MousePointer Then
        Screen.MousePointer = pMousePointer
    End If
    
End Function

Public Function SetValStringValue(ByRef pVString As String, ByVal pVName As String, ByVal pValue As String) As String
Dim bFound As Boolean
Dim lngLoop As Long
Dim lngPosOfEqual As Long
Dim strLoop As String
Dim strNewValue As String
Dim astr() As String

'   Split(expression[, delimiter[, count[, compare]]])
'   If expression is a zero-length string(""), Split returns an empty array, that is, an array
'   with no elements and no data.-> IsArray(Arr) = True, LBound(astr) = 0 and UBound(Arr) = -1
'   This behaviour used so that For Loop doesn't iteratie & ReDim Preserve creates a single
'   element array when this proc is passed a zero-length strin in pVString
    strNewValue = pVName & "=" & pValue
    astr = Split(pVString, ";")
    For lngLoop = LBound(astr) To UBound(astr)
        strLoop = astr(lngLoop)
        lngPosOfEqual = InStr(1, strLoop, "=", vbBinaryCompare)
        If lngPosOfEqual Then
            If StrComp(Trim$(Left$(strLoop, lngPosOfEqual - 1)), pVName, vbTextCompare) = 0 Then
                astr(lngLoop) = strNewValue
                bFound = True
                Exit For
            End If
        End If
    Next
   
    If Not bFound Then
        ReDim Preserve astr(0 To UBound(astr) + 1) ' Append array element for new Name=Value pair
        astr(UBound(astr)) = strNewValue    ' Assign passed Name=Value pair to new array element
    End If

    pVString = Join(astr, ";")    ' Return result in ByRef parameter for abbreviated calling
    SetValStringValue = pVString  ' Return result as value of function

End Function

Public Function SQ(ByRef pVnt As Variant) As String
'   Args passed ByRef for speed, small proc can avoid side effects with care
    SQ = "'" & pVnt & "'"   ' Single Quote
End Function

Public Function StripBracketedPrefixes(ByVal pString As String) As String
'   Tonda NEEDS to be improved with optional parameters as per StripOutBracketedSubStrings
Dim strResult As String
Dim lngPosRHBracket As Long
    
    strResult = pString
    
    Do While Left$(strResult, 1) = "["
        lngPosRHBracket = InStr(strResult, "]")
        If lngPosRHBracket = 0 Then
            Exit Do ' Loop tests for 1st char of "[" which remains the same in this case
        Else
            strResult = Right$(strResult, Len(strResult) - lngPosRHBracket)
        End If
    Loop

    StripBracketedPrefixes = strResult
    
End Function

Public Function Substitute(ByVal pValToTest As Variant, ByVal pIfThis As Variant, ByVal pSubstituteWith As Variant) As Variant
'   Passed ByRef for speed only.
'   For this size procedure we can avoid side effects with care

'   Note that "-1" <> -1 because one comes in as a string and the other as a number
Dim bEquivalent As Boolean

    If IsNull(pValToTest) Then
        bEquivalent = IsNull(pIfThis)
    ElseIf IsNumeric(pIfThis) Then
        If IsNumeric(pValToTest) Then
            bEquivalent = (CDbl(pValToTest) = CDbl(pIfThis))
        End If
        
'   ElseIf Empty Then (Null, blah, blah ),vbEmpty, Datevalue
'   Lookup help for the VarType function and cater for all the types when I want
'   to fix this procedure to a state ready for library of utilities
    Else
        bEquivalent = (pValToTest = pIfThis)
    End If
    
    If bEquivalent Then
        Substitute = pSubstituteWith
    Else
        Substitute = pValToTest
    End If

End Function

Public Function TReplace(ByRef pString As String, ByRef pFind As String, ByRef pReplace As String) As String
'  Args passed ByRef for speed, small proc can avoid side effects with care

'   Replace$ is slow and is therefore only called when needed
'   InStrB is quicker than InStr and can be used when only checking for existence
'   of characters inside the range of 1-255

'   From Help: The InStrB function is used with byte data contained in a string.
'   Instead of returning the character position of the first occurrence of one
'   string within another, InStrB returns the byte position.

    If InStr(pString, pFind) <> 0 Then
        TReplace = Replace$(pString, pFind, pReplace)
    Else
        TReplace = pString
    End If

End Function

Public Function TrimWhiteSpace(ByVal pString As String) As String
'   Trims white space in manner analgous to Trim$ trimming spaces
'   Leading and trailing spaces are trimmed but embedded multiple spaces
'   are not trimmed to single spaces
    TrimWhiteSpace = Trim$(TReplace$(pString, vbTab, " "))
End Function

Public Function TRound(ByRef pNumber As Variant, Optional ByRef pDecimalPlaces As Long = 0) As Variant
'   Args passed ByRef for speed, small proc can avoid side effects with care

'   Analagous to VBA.Round introduced in VB6 but rounds predictably - unlike VBA.Round
'   Note that ROUND function in an Excel spreadsheet behaves like this function,
'   but VBA.Round function in code behind a spreadsheet behaves unpredictably.

'   Arguments
'   ---------
'   pNumber:        Numeric expression being rounded. (Accepted as variant to maintain it's sub-type)
'   pDecimalPlaces: (Optional) Number indicating how many places to the right
'                   of the decimal are included in the rounding.
'                   If omitted, integers are returned by the Round function

'   Rounds pNumber to pDecimalPlaces.
'   If pDecimalPlaces not supplied it rounds to a whole number.
'   Rounds to smallest absolute value for 4 and below, and largest absolute value for 5 and above.
'   Signed numbers: tRound (-1.25, 1) = -1.3, tRound (-1.24, 1) = -1.2
'                   tRound (1.25, 1) = 1.3,   tRound (1.24, 1) = 1.2
'                   (as you would expect for currency transactions, owed money etc)

    Dim dblTemp As Double
    dblTemp = Fix(pNumber * (10 ^ (pDecimalPlaces + 1))) + (Sgn(pNumber) * 5)
    TRound = Fix(dblTemp / 10) / (10 ^ pDecimalPlaces)

End Function



