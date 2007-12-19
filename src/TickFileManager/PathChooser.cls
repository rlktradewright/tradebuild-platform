VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForm As fPathChooser
Private mPath As String
Private mCancelled As Boolean

Private Sub Class_Initialize()
Set mForm = New fPathChooser
End Sub

Private Sub Class_Terminate()
Unload mForm
End Sub

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Public Property Let path(ByVal newvalue As String)
mPath = newvalue
End Property

Public Property Get path() As String
path = mPath
End Property

Public Sub choose()
If mPath <> "" Then mForm.path = mPath
mForm.show vbModal
mCancelled = mForm.cancelled
mPath = mForm.path
End Sub

