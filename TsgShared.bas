Attribute VB_Name = "TsgShared"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const gkLocalDriveFranchiseFolder As String = "C:\Statistics"
Public Const gkRStatsMdbRMgrVerFld As String = "RetailManagerVersion"
Public Const gkFmtDateUnambiguous As String = "dd mmm yyyy" ' possible new candidate -> "d mmm yyyy"    '''

Public Const gkCAT_CigCtn As String = "CGCTN"
Public Const gkCAT_CigPkt As String = "CGPKT"
Public Const gkCAT_Tobac As String = "TOBAC"
Public Const gkCAT_Cigar As String = "CIGAR"
Public Const gkCATList_TSG As String = "('" & gkCAT_CigCtn & "', '" & _
                                              gkCAT_CigPkt & "', '" & _
                                              gkCAT_Tobac & "', '" & _
                                              gkCAT_Cigar & "')"

Public Function GetStkPkgFullFilename(ByVal pStkFullFilename As String) As String
Dim strParentFolder As String
Dim fso As Scripting.FileSystemObject

    Set fso = New Scripting.FileSystemObject
    
    strParentFolder = fso.GetParentFolderName(pStkFullFilename)
    
    If Len(strParentFolder) Then
        strParentFolder = strParentFolder & "\"
    End If
    
    GetStkPkgFullFilename = strParentFolder & _
                            fso.GetBaseName(pStkFullFilename) & "Pkg." & _
                            fso.GetExtensionName(pStkFullFilename)

End Function

Public Function IsUseLocalDriveFranFolder() As Boolean
    IsUseLocalDriveFranFolder = InStr(Command$, "UseLocalDriveFranFolder")
End Function

Public Function fsVersion() As String
'   Error handling b/c reading version info from App object sometimes caused following error in TsgDW program
'   'Run-time error 326 - resource with identifier version not found.'
Dim lngErr As Long
Dim strResult As String

    On Error Resume Next
        strResult = App.Major & App.Minor & App.Revision
        lngErr = Err.Number
    On Error GoTo 0

    If lngErr Then
    '   Trims down to 5 char string of "? Err" when used to populate Version field in defaults table of RemoteStatistics.mdb
        strResult = "? Err: " & lngErr
    End If

    fsVersion = strResult

End Function

Function breaktext(strMessage As String) As String
    Dim ndx As Long, start_of_line As Long
    Dim ch As String
    Const LINEWIDTH = 39
    
    start_of_line = 1
    ndx = start_of_line + LINEWIDTH

    Do While ndx < Len(strMessage)
        ch = Mid(strMessage, ndx, 1)
        Do While (ndx > start_of_line) And (ch <> " ") And (ch <> vbLf)
            ndx = ndx - 1
            ch = Mid(strMessage, ndx, 1)
            DoEvents
        Loop
    ''' Mid(strMessage, ndx, 1) = vbCrLf
    ''' start_of_line = ndx + 2 ''' AUrban Think of giving it two characters so we get vbCrLf and don't truncate it to vbCr
                                ''' which will be displayed by some installations of notepad as a square.
        strMessage = Left$(String:=strMessage, length:=ndx) & vbNewLine & Right$(strMessage, Len(strMessage) - ndx)
        start_of_line = ndx + 2
        ndx = start_of_line + LINEWIDTH
        DoEvents
    Loop
    breaktext = strMessage
End Function

Sub subOpenFile(ByVal pFilename As String)
' Open a limited number of file types with associated program.
' ACROSS ALL WINDOWS VERSIONS REGARDLESS OF CONFIGURATION AND SOFTWARE INSTALLED
'      (This procedure is used across the range of TSG programs     )
'      (Therefore only open txt - don't open zip, mdb or rtf files. )
'-----------------------------------------------------------------------------
'   NB UNfortunately 'Shell Start [app] [file]' does not work for Windows XP |
'-----------------------------------------------------------------------------
    
    If Dir$(pFilename) = "" Then
        MsgBox "Could not locate '" & pFilename & "'. ", vbExclamation
    Else
        Select Case LCase$(Right$(pFilename, 4))
            Case ".txt"
            '   If the file is too large for notepad then Windows/Notepad will prompt
            '   user as tow whether they would like to open the file with Wordpad
                Shell "notepad " & DQ(pFilename), vbNormalNoFocus
        '    Case ".rtf"
        '   '   Start required for Win98 because wordpad is not installed in
        '   '   C:\Windows or a folder included in the default search paths.
        '   '   Start causes flicker - it works for Win98 but not WindowsXP
        '        Shell "start wordpad " & DQ(pFilename), vbNormalNoFocus ' Works in Windows 98
        '   Case ".zip": Shell "pkzip " & DQ(pFilename), vbNormalNoFocus
        '   Case ".mdb": Shell "msaccess " & DQ(pFilename), vbNormalNoFocus
            Case ".exe"
            '   Doubtful you should be able to execute a program file from a list box
                MsgBox "You cannot execute programs from a list.", vbExclamation
            Case Else
             '  Do Nothing. Not even a message box to say we don't open this sort of file
                MsgBox "You cannot open this type of file from a list.", vbInformation
        End Select
    End If

End Sub

