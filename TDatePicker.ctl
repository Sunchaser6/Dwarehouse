VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TDatePicker 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   1755
   Begin MSComCtl2.DTPicker dtp 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "d MMMM yyyy"
      Format          =   63111171
      CurrentDate     =   37862
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   0
   End
End
Attribute VB_Name = "TDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--------------
' Control Notes
'--------------
' DatePicker control may later allow any number of EditMode settings
' When this occurs developers will need to inspect the EditMode property to
' determine what actions to take for the various events (eg AfterEdit)
' If at this stage AfterEdit is triggered when no change has occured, then
' the event may pass a Changed parameter by Value for the developer to inspect
'-------------------------------------------------------------------------------------------
' Default Property Values
'Const kDefaultHeight As Long = 285
'Const kDefaultWidth As Long = 1755

Public Enum enmEditMode
'   For later use if allow various editing modes
    eDropDown   ' Editing only available through DropDown calendar (via Keyboard or Mouse)
End Enum

Private Type udt
    OnDropDownValue As Date     ' Compared on CloseUp to see if Value has changed
'   Property Variables
    ToolTipUC As String
    EditMode As enmEditMode
End Type
Private m As udt

Public Event Change()
Public Event AfterEdit()
'Public Event BeforeEdit(ByRef pValue As Date, ByRef Cancel As Boolean) ' Could be useful for developers

Public Sub dtp_Change()
    RaiseEvent Change
End Sub

Private Sub dtp_CloseUp()
'   Allow dtp_CloseUp event procedure to complete so that DropDown calendar closes, but if
'   date changed then enabled timer so AfterEdit event can be raised from timer event proc
    If dtp.Value <> m.OnDropDownValue Then
        tmr.Enabled = True
    End If
End Sub

Private Sub dtp_DropDown()
    m.OnDropDownValue = dtp.Value
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
' dtp_KeyDown and UserControl_KeyDown are used together so dtp can
' be dropped down (and edited once dropped down) via the keyboard,
' but it cannot be edited with the keyboard until dropped down
'
' DatePikcer KeyDown receives most keys, but setting KeyCode = 0 for numeric
' keys has no effect therefore numeric keys are handled in UserControl_KeyDown.
'
' UserControl_KeyDown does not receive arrow keys therefore arrow keys are handled in dtp_KeyDown.
'
'  When the DTPicker control is dropped down neither the form or control recieves any KeyDown events
'  (I presume the dropdown calender is a completely independent window on it's own)
'  therefore the user can use the keyboard to change date selections
'
' Neither Form or DTPicker KeyDown events receives Tab key which moves focus (No need to cater for tab key)
' The two events must be used in concert.
'--------------------------------------------------------------------------------------------------
'   Some controls (command buttons, option buttons, and check boxes) do not receive arrow-key events
'   Instead, arrow keys cause movement to another control.
'   I am emulating this behaviour below. It may be slightly confusing if you move to another ctl
'   with arrows but that control does not support moving (back) with arrows
'--------------------------------------------------------------------------------------------------
Dim intKeyCode As Integer

    intKeyCode = KeyCode
    KeyCode = 0

'   Form_KeyDown Handler only allows arrow key KeyCodes through (which it does not receive)
    Select Case intKeyCode
        Case KeyCodeConstants.vbKeyLeft, KeyCodeConstants.vbKeyUp
            CtlMovePrevious
        Case KeyCodeConstants.vbKeyRight
            CtlMoveNext
        Case KeyCodeConstants.vbKeyDown
        '   Allow Alt-DownArrow through to dropdown the editing box
            If (Shift = vbAltMask) Then
                KeyCode = intKeyCode    ' Reinstate down arrow to allow control dropdown
            Else
                CtlMoveNext
            End If
    End Select

''   Alternative code limits keyboard interaction to Alt-vbKeyDown.
''   It is not used because user may think app has hung if they use
''   arrows when the control has focus and appear to get no response.
'    If Not ((Shift = vbAltMask) And (KeyCode = KeyCodeConstants.vbKeyDown)) Then
'        KeyCode = 0
'    End If

End Sub

Private Sub tmr_Timer()
    tmr.Enabled = False
    RaiseEvent AfterEdit
End Sub

Private Sub UserControl_InitProperties()
'   Happens once in a lifetime of a custom control's instance:
'   When the developer actually creates a new instance of a control by clicking
'   on the control's toolbox icon and placing the new instance in a container
    
'   Have NOT been unable to successfully set a default height and width in
'   InitProperties for either the user control. The changing of the user control
'   width and height causes an immediate resize, but the user control width and
'   height properties remain as they were at design time and not what they
'   were changed to. There is a sort of cascading events with resizing happening
'   many times. Trying to use InitProperties to set the default size has been
'   a wheel spinning waste of time. Sizing the user control and letting the date
'   picker fill it is an effective alternative
    
'
'   UserControl.Height = kDefaultHeight
'   UserControl.Width = kDefaultWidth
    Me.EditMode = eDropDown
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
' dtp_KeyDown and UserControl_KeyDown are used together so dtp can
' be dropped down (and edited once dropped down) via the keyboard,
' but it cannot be edited with the keyboard until dropped down
'
' DatePikcer KeyDown receives most keys, but setting KeyCode = 0 for numeric
' keys has no effect therefore numeric keys are handled in UserControl_KeyDown.
'
' UserControl_KeyDown does not receive arrow keys therefore arrow keys are handled in dtp_KeyDown.
'
'  When the DTPicker control is dropped down neither the form or control recieves any KeyDown events
'  (I presume the dropdown calender is a completely independent window on it's own)
'  therefore the user can use the keyboard to change date selections
'
' Neither Form or DTPicker KeyDown events receives Tab key which moves focus (No need to cater for tab key)
' The two events must be used in concert.
'--------------------------------------------------------------------------------------------------

'   Dtp control is on top & fills user control, -> no need to verify Dtp has focus
    If Not ((Shift = vbAltMask) And (KeyCode = KeyCodeConstants.vbKeyDown)) Then
    '   Nb. Arrow key events are not received by this handler, BUT
    '       I can't unconditionally set KeyCode = 0 and expect Alt-KeyDown to get
    '       through because setting KeyCode = 0 in effect cancels the event and
    '       prevents the Alt key proceeding the pressing of vbKeyDown getting through
        KeyCode = 0
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'   Read values in to Let procedures which perform the required processing
    
    With PropBag
        Me.EditMode = .ReadProperty("EditMode", enmEditMode.eDropDown)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.MinDate = .ReadProperty("MinDate", DateValue("1 January 100"))
        Me.MaxDate = .ReadProperty("MaxDate", DateValue("31 December 9999"))
        Me.ToolTipUC = .ReadProperty("ToolTipUC", "")
        Me.Value = .ReadProperty("Value", Date)
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'   Write property values from Get procedures
    
    With PropBag
        .WriteProperty "EditMode", Me.EditMode
        .WriteProperty "Enabled", Me.Enabled
        .WriteProperty "MinDate", dtp.MinDate
        .WriteProperty "MaxDate", dtp.MaxDate
        .WriteProperty "ToolTipUC", Me.ToolTipUC
        .WriteProperty "Value", dtp.Value
    End With
    
End Sub

Private Sub UserControl_Resize()
'   Adjust DatePicker size to fill UserControl
    
    dtp.Height = UserControl.Height
    dtp.Width = UserControl.Width

End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal pEnabled As Boolean)
' Delegate property all constituent controls and UserControl
' Stored in corresponding UserControl Extender object property
' Disabling UserControl prevents constituent controls being used, but does not
' effect their appearance which is why Enabled is delegated to all constituent controls

    UserControl.Enabled = pEnabled
    dtp.Enabled = pEnabled
    
    PropertyChanged "Enabled"
    
End Property

Public Property Get ToolTipUC() As String
' When property was named ToolTipText property procedures didn't execute
' (but UserControl.ToolTipText property was accessed)
' Property was renamed so constituent controls properties could be set

' Property applied to UserControl and all constituent controls otherwise
' tooltip only displays for areas of UC not covered by constituent controls
    
    ToolTipUC = m.ToolTipUC

End Property

Public Property Let ToolTipUC(ByVal pToolTipText As String)
' When property was named ToolTipText property procedures didn't execute
' (but UserControl.ToolTipText property was accessed)
' Property was renamed so constituent controls properties could be set

' Property applied to UserControl and all constituent controls otherwise
' tooltip only displays for areas of UC not covered by constituent controls

Dim ctl As VB.Control

    m.ToolTipUC = pToolTipText
    
    UserControl.Extender.ToolTipText = m.ToolTipUC
    On Error Resume Next    ' For controls without corresponding property
        For Each ctl In UserControl.Controls
            ctl.ToolTipText = m.ToolTipUC
        Next ctl
        Set ctl = Nothing
    On Error GoTo 0
    
    PropertyChanged "ToolTipUC"

End Property

Public Property Get EditMode() As enmEditMode
    EditMode = m.EditMode
End Property

Public Property Let EditMode(pEditMode As enmEditMode)
    m.EditMode = pEditMode
    PropertyChanged "EditMode"
End Property

Public Property Get MinDate() As Date
    MinDate = dtp.MinDate
End Property

Public Property Let MinDate(pMinDate As Date)
    dtp.MinDate = pMinDate
    PropertyChanged "MinDate"
End Property

Public Property Get MaxDate() As Date
    MaxDate = dtp.MaxDate
End Property

Public Property Let MaxDate(pMaxDate As Date)
    dtp.MaxDate = pMaxDate
    PropertyChanged "MaxDate"
End Property

Public Property Get Value() As Date
Attribute Value.VB_UserMemId = 0
    Value = dtp.Value
End Property

Public Property Let Value(ByVal pDate As Date)
    dtp.Value = pDate
    PropertyChanged "Value"
End Property
