VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'/**
'   Holds a reference to an object that does not result in the object's
'   reference count being increased.
'*/


'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mObjectAddress As Long
Private mInitialised As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise(ByVal targetObject As Object)

If mInitialised Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "TradeBuild.WeakReference::initialise", _
                "Target already set"
End If

If targetObject Is Nothing Then
    mObjectAddress = 0
Else
    mObjectAddress = ObjPtr(targetObject)
End If

mInitialised = True

End Sub

Public Property Get target() As Object

Dim lObject As Object

If mObjectAddress = 0 Then
    Set target = Nothing
Else
    CopyMemory lObject, mObjectAddress, 4
    Set target = lObject
    CopyMemory lObject, 0&, 4
End If

End Property

'================================================================================
' Helper Functions
'================================================================================

