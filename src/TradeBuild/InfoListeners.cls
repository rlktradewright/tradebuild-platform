VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoListeners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Implemented Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

Public mLastKeyAdded As String  ' for debugging only

'================================================================================
' Member variables and constants
'================================================================================

Private mAllListeners As WeakCollection
    ' each entry in this collection is a listener that listens on all
    ' value types
Private mListenerCols As Collection
    ' each entry in this collection is itself a collection of listeners,
    ' keyed by requiredData
    
'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mAllListeners = New WeakCollection
Set mListenerCols = New Collection
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Function add(ByVal newListener As InfoListener, _
                Optional ByVal valueType As Long = TradeBuildListenValueTypes.VTAll) As String
Dim key As String
Dim listeners As WeakCollection

If valueType <> TradeBuildListenValueTypes.VTAll Then
    On Error Resume Next
    Set listeners = mListenerCols("V" & valueType)
    If listeners Is Nothing Then
        Set listeners = New WeakCollection
        mListenerCols.add listeners, "V" & valueType
    End If
    On Error GoTo 0

    key = valueType & Chr(0) & gGenerateGUIDString
    listeners.add newListener, key
Else
    key = gGenerateGUIDString
    mAllListeners.add newListener, key
End If

add = key
mLastKeyAdded = key
End Function

Public Function Item(ByRef key As String) As InfoListener
Dim listeners As WeakCollection

If InStr(1, key, Chr(0)) <> 0 Then
    Set listeners = mListenerCols("V" & Left(key, InStr(1, key, Chr(0)) - 1))
    Set Item = listeners(key)
Else
    Set Item = mAllListeners(key)
End If
End Function

Public Function Count() As Long
Dim listeners As WeakCollection

For Each listeners In mListenerCols
    Count = Count + listeners.Count
Next
Count = Count + mAllListeners.Count
End Function

Public Sub remove(ByVal key As String)
Dim listeners As WeakCollection
Dim listener As InfoListener

If InStr(1, key, Chr(0)) <> 0 Then
    Set listeners = mListenerCols("V" & Left(key, InStr(1, key, Chr(0)) - 1))
    Set listener = listeners(key)
    listeners.remove key
Else
    mAllListeners.remove key
End If
End Sub

Public Sub notify(value As Variant, valueType As Long, source As Object)
Dim listeners As WeakCollection
Dim listener As InfoListener
Dim i As Long
Dim ev As InfoEvent

Set ev.source = source
ev.data = value
ev.valueType = valueType

On Error Resume Next
Set listeners = mListenerCols("V" & valueType)
On Error GoTo 0

For i = 1 To mAllListeners.Count
    Set listener = mAllListeners(i)
    listener.notify ev
Next
If Not listeners Is Nothing Then
    For i = 1 To listeners.Count
        Set listener = listeners(i)
        listener.notify ev
    Next
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================





