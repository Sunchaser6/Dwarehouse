VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUploadBataRpts 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload Bata Reports"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Transaction Dates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3510
      TabIndex        =   4
      Top             =   195
      Width           =   2595
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   285
         Left            =   555
         TabIndex        =   5
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "d MMMM yyyy"
         Format          =   17694723
         CurrentDate     =   40124
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   285
         Left            =   555
         TabIndex        =   6
         Top             =   705
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "d MMMM yyyy"
         Format          =   17694723
         CurrentDate     =   40124
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "From"
         Height          =   180
         Index           =   1
         Left            =   15
         TabIndex        =   8
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "To"
         Height          =   180
         Index           =   2
         Left            =   30
         TabIndex        =   7
         Top             =   750
         Width           =   450
      End
   End
   Begin VB.ListBox lstFranchise 
      Height          =   7080
      Left            =   210
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   285
      Width           =   3150
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   3510
      TabIndex        =   2
      Top             =   2565
      Width           =   2625
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "&Upload Reports to BATA"
      Height          =   360
      Left            =   3510
      TabIndex        =   1
      Top             =   1620
      Width           =   2625
   End
   Begin VB.Label lbl 
      Caption         =   "Selected Franchises"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   1770
   End
End
Attribute VB_Name = "frmUploadBataRpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents moBataRpts As clsBataRpts
Attribute moBataRpts.VB_VarHelpID = -1

Public Sub LoadLstFranchise()
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    lstFranchise.Clear
    
    strSQL = "SELECT FranchiseIDTSG, FranchiseBusinessName" & vbNewLine & _
             "FROM qryFranchiseBata" & vbNewLine & _
             "ORDER BY FranchiseBusinessName"
    
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    
    If Not rst Is Nothing Then
        With rst
            Do While Not .EOF
                lstFranchise.AddItem .Fields!FranchiseBusinessName.Value
                lstFranchise.ItemData(lstFranchise.NewIndex) = .Fields!FranchiseIDTSG.Value
                .MoveNext
            Loop
        End With
        
        rst.Close
        Set rst = Nothing
        
    End If

End Sub

Private Sub ConfigureButtons()
Dim bEnabled As Boolean

    If lstFranchise.SelCount <> 0 Then
        If dtpFrom.Value <= dtpTo.Value Then
            If dtpFrom <= Date Then
                bEnabled = True
            End If
        End If
    End If
    
    cmdUpload.Enabled = bEnabled
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpload_Click()
Dim intPrevMousePointer As Integer
Dim strFranSeln As String
Dim strDateSeln As String
Dim strMsg As String
Dim colFranIDs As VBA.Collection

    If lstFranchise.SelCount Then
        Set colFranIDs = ListBoxGetCollection(pListBox:=lstFranchise, pItemData:=True, pSelected:=True)
        
        strDateSeln = Format$(dtpFrom.Value, gkFmtDateUnambiguous) & " to " & Format$(dtpTo.Value, gkFmtDateUnambiguous)
        If colFranIDs.Count = lstFranchise.ListCount Then
            strFranSeln = "All Franchises"
        Else
            strFranSeln = Plural(pQty:=colFranIDs.Count, pNounSingular:="franchise") & " selected"
        End If
        
        strMsg = "Upload Bata Reports?" & vbNewLine & strFranSeln & vbNewLine & strDateSeln
                  
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Me.Enabled = False
            Me.Hide
            intPrevMousePointer = SetMousePointer(vbHourglass)
            
            strMsg = "Manual Bata Report Upload: " & strFranSeln & " for " & strDateSeln
            StatusBar pMsg:=strMsg & " Started"
        
            Set moBataRpts = New clsBataRpts
            moBataRpts.AddRpts_FranAndDateSeln pColFranIDs:=colFranIDs, _
                                               pDateFrom:=dtpFrom.Value, _
                                               pDateTo:=dtpTo.Value
            moBataRpts.Upload pAddUnsent:=False ' Previous line added reports selected for upload
            Set moBataRpts = Nothing
            
            StatusBar pMsg:=strMsg & " Completed"
            gsubRefreshEventLogDisplay
        
            SetMousePointer intPrevMousePointer
            Unload Me
        End If
        
    End If

End Sub

Private Sub Form_Load()
Dim dtmToday As Date

    LoadLstFranchise
    ConfigureButtons
    
    dtmToday = Date
    With dtpFrom
        .Value = dtmToday
    ''' .MinDate = DateSerial(Year:=2000, Month:=1, Day:=1) ' Can go back into archive but give a reasonable limit  ''' V386
        .MinDate = g.dtmLiveDataStart ' Go back as far available data in LiveData table                             ''' V386
        .MaxDate = dtmToday
    End With
    
    With dtpTo
        .Value = dtmToday
        .MinDate = dtpFrom.Value
        .MaxDate = dtmToday
    End With

End Sub

Private Sub dtpFrom_Change()
    dtpTo.MinDate = dtpFrom.Value
End Sub

Private Sub dtpTo_Change()
    dtpFrom.MaxDate = dtpTo.Value
End Sub

Private Sub lstFranchise_Click()
    ConfigureButtons
End Sub

Private Sub moBataRpts_AfterRptLoad(oRpt As clsBataRpt, ByVal Success As Boolean)
Dim strMsg As String

    strMsg = "Loading " & oRpt.FranName & " - " & oRpt.Name
    If Not Success Then
        If oRpt.HasData Then
            strMsg = strMsg & " FAILED."
        Else
            strMsg = strMsg & " NO DATA."
        End If
    End If
    
    StatusBar pMsg:=strMsg, pLog:=False

End Sub

Private Sub moBataRpts_AfterRptUpload(oRpt As clsBataRpt, ByVal Success As Boolean, ByVal ErrMsg As String)
Dim strMsg As String

    strMsg = "Uploading " & oRpt.FranName & " - " & oRpt.Name
    If Success Then
        strMsg = strMsg & " succeeded"
    Else
    '   Use uppercase for first part of message to draw attention to problem, but maintain mixed case of
    '   returned ErrMsg to preserve diagnostic details and also place ErrMsg on new line to improve readability
        strMsg = UCase$(strMsg & " FAILED. ") & vbNewLine & ErrMsg
    End If
    
    StatusBar pMsg:=strMsg

End Sub

Private Sub moBataRpts_BeforeRptUpload(oRpt As clsBataRpt, ByVal UploadAttempt As Long)
    StatusBar pMsg:="Uploading " & oRpt.FranName & " - " & oRpt.Name & " (attempt " & UploadAttempt & ")", pLog:=False
End Sub

Private Sub moBataRpts_OnRptLoad(oRpt As clsBataRpt)
    StatusBar pMsg:="Loading " & oRpt.FranName & " - " & oRpt.Name, pLog:=False
End Sub

