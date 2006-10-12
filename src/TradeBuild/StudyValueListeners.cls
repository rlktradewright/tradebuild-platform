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

Implements TradeBuildSP.IStudyValueListeners

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type ListenerTableEntry
    theListener         As TradeBuildSP.IStudyValueListener
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
' IStudyValueListeners Interface Members
'================================================================================

Private Sub IStudyValueListeners_Add( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String)
Dim tableEntry As ListenerTableEntry

If mNextListenerTableIndex > UBound(mListenerTable) Then
    ReDim Preserve mListenerTable(UBound(mListenerTable) + 10) As ListenerTableEntry
End If
Set tableEntry.theListener = listener
tableEntry.valueName = valueName
mListenerTable(mNextListenerTableIndex) = tableEntry
mNextListenerTableIndex = mNextListenerTableIndex + 1
End Sub

Private Function IStudyValueListeners_Count() As Long
IStudyValueListeners_Count = mNextListenerTableIndex
End Function

Private Function IStudyValueListeners_Item( _
                ByVal index As Long) As TradeBuildSP.IStudyValueListener
Set IStudyValueListeners_Item = mListenerTable(index).theListener
End Function

Private Sub IStudyValueListeners_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim i As Long
Dim listener As TradeBuildSP.IStudyValueListener

For i = 0 To mNextListenerTableIndex - 1
    Set listener = mListenerTable(i).theListener
    If Not listener Is Nothing Then
        If ev.valueName = mListenerTable(i).valueName Then
            listener.notify ev
        End If
    End If
Next
End Sub

Private Sub IStudyValueListeners_Remove( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
Dim i As Long
Dim lListener As TradeBuildSP.IStudyValueListener

For i = 0 To mNextListenerTableIndex - 1
    Set lListener = mListenerTable(i).theListener
    If lListener Is listener Then
        Set mListenerTable(i).theListener = Nothing
        mListenerTable(i).valueName = ""
    End If
Next
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

'================================================================================
' Helper Functions
'================================================================================








