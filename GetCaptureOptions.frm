VERSION 5.00
Begin VB.Form fdlgGetCaptureOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Populated in Display Method"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4590
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
   ScaleHeight     =   2385
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkUploadBataRpts 
      BackColor       =   &H0000FF00&
      Caption         =   "Upload BATA reports"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2355
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.ListBox lstFranchises 
      Height          =   2010
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   255
      Width           =   2040
   End
   Begin VB.CheckBox chkUpdateNonCompliants 
      Caption         =   "Update Non Compliants"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2355
      TabIndex        =   3
      Top             =   1815
      Width           =   2010
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2280
      TabIndex        =   1
      Top             =   285
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Franchises"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   45
      Width           =   810
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "fdlgGetCaptureOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type udt
    Result As clsDataCaptureOptions
End Type
Private m As udt

Public Enum DataCaptureEnum
'   eCaptureALL
    eCaptureSelected
End Enum

Private Sub cmdCancel_Click()
    m.Result.CancelCapture = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m.Result.CancelCapture = False ' required because code part of form seems to stay loaded
    Me.Hide
End Sub

Public Function GetOptions(ByVal pDataCapture As DataCaptureEnum, _
                  Optional ByRef pColFranNames As VBA.Collection = Nothing) As clsDataCaptureOptions
'   Returns selected data capture options
Dim strTitle As String
Dim vntFranName As Variant

    Select Case pDataCapture
        Case DataCaptureEnum.eCaptureSelected
            strTitle = "Capture SELECTED Franchises"
    '   Case DataCaptureEnum.eCaptureALL
    '       strTitle = "Capture ALL Franchises"
    End Select
 
    Me.Caption = strTitle
        
    If pColFranNames Is Nothing Then
        Me.lstFranchises.AddItem "ALL"
    Else
        For Each vntFranName In pColFranNames
            lstFranchises.AddItem vntFranName
        Next vntFranName
    End If
    
    Me.Show vbModal
    
'   Assign members of m.Result
    m.Result.UpdateNonCompliants = ChkBoxToBool(chkUpdateNonCompliants)
    
    Set GetOptions = m.Result

    Unload Me
    
End Function

Private Sub Form_Initialize()
    Set m.Result = New clsDataCaptureOptions
End Sub

Private Sub lstFranchises_Click()
'   Disable selections of franchise while leaving list
'   enables so user can scroll through list
    lstFranchises.ListIndex = -1
End Sub
