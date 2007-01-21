VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueListeners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Type ListenerTableEntry
    theListener         As StudyValueListener
    valueName           As String
End Type

'================================================================================
' Member variables and constants
'================================================================================

Private mListenerTable() As ListenerTableEntry
Private mNextListenerTableIndex As Long

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mListenerTable(10) As ListenerTableEntry
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

Public Sub add( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String)
Dim tableEntry As ListenerTableEntry

If mNextListenerTableIndex > UBound(mListenerTable) Then
    ReDim Preserve mListenerTable(UBound(mListenerTable) + 10) As ListenerTableEntry
End If
Set tableEntry.theListener = listener
tableEntry.valueName = UCase$(valueName)
mListenerTable(mNextListenerTableIndex) = tableEntry
mNextListenerTableIndex = mNextListenerTableIndex + 1
End Sub

Public Sub notify( _
                ev As StudyValueEvent)
Dim i As Long
Dim listener As StudyValueListener

For i = 0 To mNextListenerTableIndex - 1
    Set listener = mListenerTable(i).theListener
    If Not listener Is Nothing Then
        If UCase$(ev.valueName) = mListenerTable(i).valueName Then
            listener.notify ev
        End If
    End If
Next
End Sub

Public Sub remove( _
                ByVal listener As StudyValueListener)
Dim i As Long
Dim lListener As StudyValueListener

For i = 0 To mNextListenerTableIndex - 1
    Set lListener = mListenerTable(i).theListener
    If lListener Is listener Then
        Set mListenerTable(i).theListener = Nothing
        mListenerTable(i).valueName = ""
    End If
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================










