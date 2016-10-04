VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fdlgCommon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Populated in Display Method"
   ClientHeight    =   525
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   510
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlg 
      Left            =   30
      Top             =   -15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "fdlgCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
             
Public Enum ShowOpenOrSaveEnum
    eShowSave = 1
    eShowOpen = 2
End Enum

Public Function GetFullFileName(ByVal pMethod As ShowOpenOrSaveEnum, _
                       Optional ByVal pFilename As String = vbNullString, _
                       Optional ByVal pInitDir As String = vbNullString, _
                       Optional ByVal pFilter As String = vbNullString, _
                       Optional ByVal pFilterDescription As String, _
                       Optional ByVal pDefaultExtension As String, _
                       Optional ByVal pDialogTitle As String = vbNullString) As String

'   Returns AbsoluteFileName as a string when a selection has been made otherwise returns vbNullString

' ---------------------
' Common Dialog Control
' ---------------------
' Flags
'   cdlOFNHideReadOnly:  Hides the Read Onlycheck box.
'   cdlOFNFileMustExist: Specifies that the user can enter only names of existing files in the File Name text box.
'                        If this flag is set and the user enters an invalid filename, a warning is displayed.
'                        (APU Flag has no effect when using ShowSave method rather than ShowOpen)
'   When combining flags (cdlOFNFileMustExist + cdlOFNHideReadOnly) is equivalent to (cdlOFNFileMustExist Or cdlOFNHideReadOnly)

' Properties
'   DefaultExt - When a file with no extension is saved, the extension specified by this property is automatically appended to the filename.
'   FilterIndex -Specifies the default filter when you use the Filter property to specify filters for an Open or Save As dialog box.
'                The index for the first defined filter is 1.
'   Filename -   Returns or sets the path and filename of a selected file.
'   Filter -     A filter specifies the type of files that are displayed in the dialog box's file list box.
'                You can do this by setting the Filter property using the following format:
'                   description1 | filter1 | description2 | filter2...
'               Description is the string displayed in the list box — for example, "Text Files (*.txt)."
'               Filter is the actual file filter — for example, "*.txt." or "*.txt;*.rtf *.doc"
'               Each description | filter set must be separated by a pipe symbol (|).
'               For example, selecting the filter *.txt displays all text files.
'               Example of multiple filters and multiple file extensions witin a filter
'                -> .Filter = "Text (*.txt)|*.txt|Pictures (*.bmp;*.ico)|*.bmp;*.ico"
'               When you specify more than one filter for a dialog box, use the FilterIndex
'               property to determine which filter is displayed as the default.
Dim lngErr As Long
Dim bCancelled As Boolean

    With dlg
        .CancelError = True ' Error generated when user chooses Cancel button
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        If Len(pInitDir) Then .InitDir = pInitDir
        If Len(pDialogTitle) Then .DialogTitle = pDialogTitle
        If Len(pFilename) Then .FileName = pFilename
        If Len(pDefaultExtension) Then .DefaultExt = "." & pDefaultExtension
        If Len(pFilter) Then
            If Len(pFilterDescription) Then
                .Filter = pFilterDescription & "|" & pFilter
            Else
                .Filter = "(" & pFilter & ")|" & pFilter
            End If
        End If

    '   Trap error of user selecting cancel button
        On Error Resume Next
            Select Case pMethod
                Case eShowOpen: .ShowOpen
                Case eShowSave: .ShowSave
            End Select
            lngErr = Err.Number
        On Error GoTo 0
        
        If lngErr <> 0 Then
            If lngErr = cdlCancel Then
                bCancelled = True
            Else
                Err.Raise lngErr
            End If
        End If
        
        If Not bCancelled Then
            GetFullFileName = .FileName
        End If
        
    End With

'   Allow Screen Refresh to remove dialog form
    DoEvents
    Unload Me

End Function

Public Function GetFont(Optional pDialogTitle As String = vbNullString, Optional pFont As StdFont = Nothing) As StdFont

'   Returns font object with associated selections made via common dialog box

' ---------------------
' Common Dialog Control
' ---------------------
'   Fonts Dialog Box Flags
'    cdlCFANSIOnly &H400 Specifies that the dialog box allows only a selection of the fonts that use the Windows character set. If this flag is set, the user won't be able to select a font that contains only symbols.
'    cdlCFApply &H200 Enables the Apply button on the dialog box.
'    cdlCFBoth &H3 Causes the dialog box to list the available printer and screen fonts. The hDC property identifies thedevice context associated with the printer.
'    cdlCFEffects &H100 Specifies that the dialog box enables strikethrough, underline, and color effects.
'    cdlCFFixedPitchOnly &H4000 Specifies that the dialog box selects only fixed-pitch fonts.
'    cdlCFForceFontExist &H10000 Specifies that an error message box is displayed if the user attempts to select a font or style that doesn't exist.
'    cdlCFHelpButton &H4 Causes the dialog box to display a Help button.
'    cdlCFLimitSize &H2000 Specifies that the dialog box selects only font sizes within the range specified by the Min and Max properties.
'    cdlCFNoFaceSel &H80000 No font name selected.
'    cdlCFNoSimulations &H1000 Specifies that the dialog box doesn't allow graphic device interface (GDI) font simulations.
'    cdlCFNoSizeSel &H200000 No font size selected.
'    cdlCFNoStyleSel &H100000 No style was selected.
'    cdlCFNoVectorFonts &H800 Specifies that the dialog box doesn't allow vector-font selections.
'    cdlCFPrinterFonts &H2 Causes the dialog box to list only the fonts supported by the printer, specified by the hDC property.
'    cdlCFScalableOnly &H20000 Specifies that the dialog box allows only the selection of fonts that can be scaled.
'    cdlCFScreenFonts &H1 Causes the dialog box to list only the screen fonts supported by the system.
'    cdlCFTTOnly &H40000 Specifies that the dialog box allows only the selection of TrueType fonts.
'    cdlCFWYSIWYG &H8000 Specifies that the dialog box allows only the selection of fonts that are available on both the printer and on screen. If this flag is set, the cdlCFBoth and cdlCFScalableOnly flags should also be set.

Dim bCancelled As Boolean
Dim lngErr As Long
    
    With dlg
        If Len(pDialogTitle) Then
            .DialogTitle = pDialogTitle
        End If
        .Flags = cdlCFANSIOnly Or _
                 cdlCFForceFontExist Or _
                 cdlCFScreenFonts Or _
                 cdlCFTTOnly
            '    cdlCCRGBInit Or cdlCFEffects Or cdlCCFullOpen   (Don't include colour)

    '   Set FontSize limiting properties (don't allow a fontsize < 8)
        .Flags = .Flags Or cdlCFLimitSize:  .Min = 8: .Max = 72
                
        If Not pFont Is Nothing Then
            .FontName = pFont.Name
            .FontSize = pFont.Size
            .FontBold = pFont.Bold
            .FontItalic = pFont.Italic
        '   .FontUnderline = pFont.Underline
        '   .FontStrikethru = pFont.Strikethrough
        '   .hDC
        End If
        
    '   Trap error of user selecting cancel button
        On Error Resume Next
            .ShowFont
            lngErr = Err.Number
        On Error GoTo 0
        
        If lngErr <> 0 Then
            If lngErr = cdlCancel Then
                bCancelled = True
            Else
                Err.Raise lngErr
            End If
        End If
        
        If Not bCancelled Then
            Dim font As StdFont
            Set font = New StdFont
            With dlg
                font.Name = dlg.FontName
                font.Size = dlg.FontSize
                font.Bold = dlg.FontBold
                font.Italic = dlg.FontItalic
            '   Font.Underline = dlg.FontUnderline
            '   Font.Strikethrough = dlg.FontStrikethru
            End With
            Set GetFont = font
            Set font = Nothing
        End If

    End With
    
    DoEvents    '   Allow Screen Refresh to remove dialog form
    Unload Me

End Function

Public Function GetPrinter(Optional pDialogTitle As String = vbNullString, Optional pFont As StdFont = Nothing) As String
' ******************************************************************************'
' UNDER DEVELOPMENT - UNDER DEVELOPMENT - UNDER DEVELOPMENT - UNDER DEVELOPMENT '
' ******************************************************************************'
' Is possible there may be no use for using the common dialog printer options other than for setting the Windows default printer

' ---------------------
' Common Dialog Control
' ---------------------
'Printer Dialog Box Flags
'   Constant Value Description
'   cdlPDAllPages &H0 Returns or sets the state of the All Pagesoption button.
'   cdlPDCollate &H10 Returns or sets the state of the Collatecheck box.
'   cdlPDDisablePrintToFile &H80000 Disables the Print To File check box.
'   cdlPDHelpButton &H800 Causes the dialog box to display the Help button.
'   cdlPDHidePrintToFile &H100000 Hides the Print To File check box.
'   cdlPDNoPageNums &H8 Disables the Pages option button and the associated edit control.
'   cdlPDNoSelection &H4 Disables the Selection option button.
'   cdlPDNoWarning &H80 Prevents a warning message from being displayed when there is no default printer.
'   cdlPDPageNums &H2 Returns or sets the state of the Pages option button.
'   cdlPDPrintSetup &H40 Causes the system to display the Print Setup dialog box rather than the Print dialog box.
'   cdlPDPrintToFile &H20 Returns or sets the state of the Print To File check box.
'   cdlPDReturnDC &H100 Returns adevice context for the printer selection made in the dialog box. The device context is returned in the dialog box's hDC property.
'   cdlPDReturnDefault &H400 Returns default printer name.
'   cdlPDReturnIC &H200 Returns an information context for the printer selection made in the dialog box. An information context provides a fast way to get information about the device without creating a device context. The information context is returned in the dialog box's hDC property.
'   cdlPDSelection &H1 Returns or sets the state of the Selection option button. If neither cdlPDPageNums nor cdlPDSelection is specified, the All option button is in the selected state.
'   cdlPDUseDevModeCopies &H40000 If a printer driver doesn't support multiple copies, setting this flag disables the copies edit control. If a driver does support multiple copies, setting this flag indicates that the dialog box stores the requested number of copies in the Copies property.

Dim bCancelled As Boolean
Dim lngErr As Long
Dim prn As VB.Printer
Dim s As String
    For Each prn In Printers
        s = s & prn.DeviceName & vbNewLine
    Next prn
    'MsgBox "Printers List: " & vbNewLine & s & vbNewLine & vbNewLine & Printers.Count & "printers"
    
'    'cdlPDPrintToFile - Returns or sets the state of the Print To File check box.
'    'and cdlPDPrintToFile cdlPDDisablePrintToFile
'    dlg.Flags = cdlPDReturnDC Or cdlPDDisablePrintToFile    ' <- this is how you set the flags
'    Me.dlg.ShowPrinter
'    MsgBox Me.dlg.hDC
'    Unload Me
'
'Exit Function
'----------------------------------------------------------------------------------------------------
    
    With dlg
        If Len(pDialogTitle) Then
            .DialogTitle = pDialogTitle
        End If
        .Flags = cdlPDReturnDC Or _
                 cdlPDDisablePrintToFile Or _
                 cdlPDNoPageNums Or _
                 cdlPDUseDevModeCopies
            '    cdlCCRGBInit Or cdlCFEffects Or cdlCCFullOpen   (Don't include colour)

'    '   Set FontSize limiting properties (don't allow a fontsize < 8)
'        .Flags = .Flags Or cdlCFLimitSize:  .Min = 8: .Max = 72
'
'        If Not pFont Is Nothing Then
'            .FontName = pFont.Name
'            .FontSize = pFont.Size
'            .FontBold = pFont.Bold
'            .FontItalic = pFont.Italic
'        '   .FontUnderline = pFont.Underline
'        '   .FontStrikethru = pFont.Strikethrough
'        '   .hDC
'        End If
        
    '   Trap error of user selecting cancel button
        On Error Resume Next
            .ShowPrinter
            lngErr = Err.Number
        On Error GoTo 0
        
        If lngErr <> 0 Then
            If lngErr = cdlCancel Then
                bCancelled = True
            Else
                Err.Raise lngErr
            End If
        End If
        
        If Not bCancelled Then
'            Dim font As StdFont
'            Set font = New StdFont
'            With dlg
'                font.Name = dlg.FontName
'                font.Size = dlg.FontSize
'                font.Bold = dlg.FontBold
'                font.Italic = dlg.FontItalic
'            '   Font.Underline = dlg.FontUnderline
'            '   Font.Strikethrough = dlg.FontStrikethru
'            End With
'            Set GetFont = font
'            Set font = Nothing
        End If

    End With
    
    DoEvents    '   Allow Screen Refresh to remove dialog form
    Unload Me


End Function
