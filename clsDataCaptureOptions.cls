VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataCaptureOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m As udt
Private Type udt
    bCaptureCancelled As Boolean
    bUpdateNonCompliants As Boolean
End Type

Public Property Let UpdateNonCompliants(pValue As Boolean)
    m.bUpdateNonCompliants = pValue
End Property

Public Property Get UpdateNonCompliants() As Boolean
    UpdateNonCompliants = m.bUpdateNonCompliants
End Property

Public Property Let CancelCapture(pValue As Boolean)
    m.bCaptureCancelled = pValue
End Property

Public Property Get CancelCapture() As Boolean
    CancelCapture = m.bCaptureCancelled
End Property

Private Sub Class_Initialize()
'   Set default values used when user is not prompted for Data Capture Options
    m.bCaptureCancelled = False
    m.bUpdateNonCompliants = True
End Sub
