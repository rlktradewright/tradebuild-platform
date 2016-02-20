VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnhandledExcptnHndlr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mUEH As UnhandledErrorHandler
Attribute mUEH.VB_VarHelpID = -1

Private Sub Class_Initialize()
Set mUEH = UnhandledErrorHandler
End Sub

Private Sub mUEH_UnhandledError(ev As ErrorEventData)
'LogMessage "Unhandled error: " & ev.ErrorMessage & vbCrLf & _
'            "At:" & vbCrLf & _
'            ev.ErrorSource
mUEH.Handled = True
End Sub
