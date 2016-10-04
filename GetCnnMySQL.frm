VERSION 5.00
Begin VB.Form fdlgGetCnnMySQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MySQL Connection Settings"
   ClientHeight    =   2340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5940
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowPwd 
      Caption         =   "&Show Password"
      Height          =   195
      Left            =   4380
      TabIndex        =   10
      Top             =   1785
      Width           =   1470
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1065
      PasswordChar    =   "l"
      TabIndex        =   9
      Top             =   1725
      Width           =   3210
   End
   Begin VB.TextBox txtUID 
      Height          =   285
      Left            =   1065
      TabIndex        =   7
      Top             =   1335
      Width           =   3210
   End
   Begin VB.TextBox txtDatabase 
      Height          =   285
      Left            =   1065
      TabIndex        =   5
      Top             =   960
      Width           =   3210
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1065
      TabIndex        =   3
      Top             =   570
      Width           =   3210
   End
   Begin VB.TextBox txtDriver 
      Height          =   285
      Left            =   1065
      TabIndex        =   1
      Top             =   195
      Width           =   3210
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4710
      TabIndex        =   12
      Top             =   855
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4710
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   270
      Width           =   855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   4
      Left            =   255
      TabIndex        =   8
      Top             =   1785
      Width           =   690
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      Height          =   195
      Index           =   3
      Left            =   615
      TabIndex        =   6
      Top             =   1380
      Width           =   330
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      Height          =   195
      Index           =   2
      Left            =   255
      TabIndex        =   4
      Top             =   1005
      Width           =   690
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   615
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Driver"
      Height          =   195
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   225
      Width           =   420
   End
End
Attribute VB_Name = "fdlgGetCnnMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type udt
'   bCancel As Boolean
    strCnnString As String
    cnn As ADODB.Connection
End Type
Private m As udt

Private Sub chkShowPwd_Click()
    If ChkBoxToBool(chkShowPwd) Then
        txtPwd.font = "Microsoft Sans Serif"
        txtPwd.PasswordChar = vbNullString
    Else
    '   Displayed as a bold filled cirle in Wingdings
    '   Unfortunately appears VB6 won't display Micorsoft San Serif unicode char that applies
        txtPwd.font = "Wingdings"
        txtPwd.PasswordChar = "l"
    End If
End Sub

Private Sub cmdCancel_Click()
'   m.bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim strErrMsg As String

'   Currently no validation enabling/disabling cmdOK, just going with what's in the text boxes
    
    SetValStringValue pVString:=m.strCnnString, pVName:="DRIVER", pValue:=Trim$(txtDriver)
    SetValStringValue pVString:=m.strCnnString, pVName:="SERVER", pValue:=txtServer
    SetValStringValue pVString:=m.strCnnString, pVName:="DATABASE", pValue:=txtDatabase
    SetValStringValue pVString:=m.strCnnString, pVName:="UID", pValue:=txtUID
    SetValStringValue pVString:=m.strCnnString, pVName:="PWD", pValue:=txtPwd
   
''Suspend timer when making a connection (put In cmdBrowse (never liked that flag anyway). )
'Group g.rstDwDefaults and g.rstEventLog with making a connection
'Ask user whether they want to write new connection string to the g.rstAppDefaults when a connection is successfully made.
   
    Set m.cnn = GetCnnMySqlFromCnnString(pCnnString:=m.strCnnString, pErrMsg:=strErrMsg)
    If m.cnn Is Nothing Then
        strErrMsg = "Can't connect to databse." & vbNewLine & "Error Message - " & strErrMsg
        MsgBox strErrMsg, vbExclamation
    Else
    '   m.bCancel = False ' required because code part of form seems to stay loaded -> ensure it is set
        Me.Hide
    End If
    
End Sub

Public Function GetCnn(ByRef pCnnString As String) As ADODB.Connection
'   Returns Connection to MySQL databse or NOTHING
'   Should this be renamed to avoid any ambiguity with the TsgTADO function

    m.strCnnString = pCnnString ''' Review IS IT NEEDED ?
    
    txtDriver = GetValStringValue(pVString:=pCnnString, pVName:="DRIVER")
    txtServer = GetValStringValue(pVString:=pCnnString, pVName:="SERVER")
    txtDatabase = GetValStringValue(pVString:=pCnnString, pVName:="DATABASE")
    txtUID = GetValStringValue(pVString:=pCnnString, pVName:="UID")
    txtPwd = GetValStringValue(pVString:=pCnnString, pVName:="PWD")
'   txtOption = GetValStringValue(pVString:=pCnnString, pVName:="OPTION")
    
    Me.Show vbModal
    
'   Set return values
'   IF m.bCancel THEN Blah ...
    pCnnString = m.strCnnString
    Set GetCnn = m.cnn

    Unload Me
    
End Function
