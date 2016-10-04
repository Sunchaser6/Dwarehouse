VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form fdlgGetDateRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Populated in Display Method"
   ClientHeight    =   1590
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4665
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
   ScaleHeight     =   1590
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3570
      TabIndex        =   1
      Top             =   990
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   300
      Left            =   3570
      TabIndex        =   0
      Top             =   600
      Width           =   900
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   570
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "d MMMM yyyy"
      Format          =   17760259
      CurrentDate     =   37862
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   975
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "d MMMM yyyy"
      Format          =   17760259
      CurrentDate     =   37862
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input propmt"
      Height          =   435
      Left            =   225
      TabIndex        =   8
      Top             =   90
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFromCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   615
      Width           =   345
   End
   Begin VB.Label lblToCaption 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1020
      Width           =   345
   End
   Begin VB.Label lblToDayOfWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " (Monday, ...)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2430
      TabIndex        =   5
      Top             =   1035
      Width           =   1140
   End
   Begin VB.Label lblFromDayOfWeek 
      AutoSize        =   -1  'True
      Caption         =   " (Monday, ...)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2430
      TabIndex        =   3
      Top             =   630
      Width           =   1140
   End
End
Attribute VB_Name = "fdlgGetDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' GetDateRange dialog form
'   - requires reference for DTPicker

Public Enum DateRangeRtnEnum
    eMySql_BetweenClause
End Enum

Private Type udt
    bCancelled As Boolean
End Type

Private m As udt

Private Sub cmdCancel_Click()
    m.bCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m.bCancelled = False ' required b/c form code module seems to stay loaded
    Me.Hide
End Sub

Private Sub dtpFrom_Change()
    dtpTo.MinDate = dtpFrom.Value
    lblFromDayOfWeek.Caption = Format$(dtpFrom.Value, "dddd")
End Sub

Public Function GetDateRange(pPrompt As String, _
                             pFromDate As Variant, _
                             pToDate As Variant, _
                             pReturnType As DateRangeRtnEnum, _
                    Optional pMinFromDate As Date = 0, _
                    Optional pMaxToDate As Date = 0, _
                    Optional pTitleBar As String = "Select Date Range") As Variant

Dim dtmTo As Date
Dim vntResult As Variant

'   Returns selected date or Empty to flag dialog was cancelled
    Me.Caption = App.Title & " - " & pTitleBar
    lblPrompt.Caption = pPrompt
    
    If pMinFromDate <> 0 Then dtpFrom.MinDate = pMinFromDate
    If pMaxToDate <> 0 Then dtpTo.MaxDate = pMaxToDate
    dtpFrom.Value = pFromDate
    dtpTo.Value = pToDate
    
'   Update DayOfWeek labels - change event not triggered by Value change via code
    Call dtpFrom_Change
    Call dtpTo_Change
    
    Me.Show vbModal
    
    If Not m.bCancelled Then
        Select Case pReturnType
            Case DateRangeRtnEnum.eMySql_BetweenClause
                dtmTo = DateAdd("s", -1, DateAdd("d", 1, dtpTo.Value))
                vntResult = "between " & Format$(dtpFrom.Value, "\'yyyy-mm-dd'\") & _
                               " and " & Format$(dtmTo, "\'yyyy-mm-dd hh:nn:ss'\")
        '   Case ?
        '       vntResult
            Case Else
                Err.Raise Number:=99, _
                          Source:=App.Title & " fdlgGetDateRange.GetDateRange()", _
                          Description:="Invalid value passed in pReturnType pReturnType parameter"
        End Select
        GetDateRange = vntResult
    End If

    Unload Me
    
End Function

Private Sub dtpTo_Change()
    dtpFrom.MaxDate = dtpTo.Value
    lblToDayOfWeek.Caption = Format$(dtpTo.Value, "dddd")
End Sub
