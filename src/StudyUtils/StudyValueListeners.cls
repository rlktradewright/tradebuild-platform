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

'@================================================================================
' Implemented Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ListenerTableEntry
    theListener         As StudyValueListener
    ValueName           As String
End Type

'@================================================================================
' Member variables and constants
'@================================================================================

Private mListenerTable() As ListenerTableEntry
Private mNextListenerTableIndex As Long

Private mDefaultValuename As String

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mListenerTable(3) As ListenerTableEntry
End Sub

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let DefaultValueName(ByVal value As String)
mDefaultValuename = UCase$(value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                ByVal listener As StudyValueListener, _
                ByVal ValueName As String)
Dim tableEntry As ListenerTableEntry

If mNextListenerTableIndex > UBound(mListenerTable) Then
    ReDim Preserve mListenerTable(2 * (UBound(mListenerTable) + 1) - 1) As ListenerTableEntry
End If
Set tableEntry.theListener = listener
tableEntry.ValueName = UCase$(ValueName)
mListenerTable(mNextListenerTableIndex) = tableEntry
mNextListenerTableIndex = mNextListenerTableIndex + 1
End Sub

Public Sub Notify( _
                ev As StudyValueEvent)
Dim i As Long
Dim listener As StudyValueListener

ev.ValueName = UCase$(ev.ValueName)
For i = 0 To mNextListenerTableIndex - 1
    Set listener = mListenerTable(i).theListener
    If Not listener Is Nothing Then
        If ev.ValueName = mDefaultValuename Then
            If mListenerTable(i).ValueName = DefaultStudyValueNameStr Then
                listener.Notify ev
            End If
        End If
        If ev.ValueName = mListenerTable(i).ValueName Then
            listener.Notify ev
        End If
    End If
Next
End Sub

Public Sub Remove( _
                ByVal listener As StudyValueListener)
Dim i As Long
Dim lListener As StudyValueListener

For i = 0 To mNextListenerTableIndex - 1
    Set lListener = mListenerTable(i).theListener
    If lListener Is listener Then
        Set mListenerTable(i).theListener = Nothing
        mListenerTable(i).ValueName = ""
    End If
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================










