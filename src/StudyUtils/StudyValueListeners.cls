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
    theListener         As IStudyValueListener
    ValueName           As String
End Type

'@================================================================================
' Member variables and constants
'@================================================================================

Private Const ModuleName                As String = "StudyValueListeners"

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

Public Property Let DefaultValueName(ByVal Value As String)
mDefaultValuename = UCase$(Value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                ByVal listener As IStudyValueListener, _
                ByVal ValueName As String)
Dim tableEntry As ListenerTableEntry

Const ProcName As String = "Add"
On Error GoTo Err

If mNextListenerTableIndex > UBound(mListenerTable) Then
    ReDim Preserve mListenerTable(2 * (UBound(mListenerTable) + 1) - 1) As ListenerTableEntry
End If
Set tableEntry.theListener = listener
tableEntry.ValueName = UCase$(ValueName)
mListenerTable(mNextListenerTableIndex) = tableEntry
mNextListenerTableIndex = mNextListenerTableIndex + 1

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Notify( _
                ev As StudyValueEventData)
Dim i As Long
Dim listener As IStudyValueListener

Const ProcName As String = "Notify"
On Error GoTo Err

ev.ValueName = UCase$(ev.ValueName)
For i = 0 To mNextListenerTableIndex - 1
    If mListenerTable(i).theListener Is Nothing Then
    Else
        If ev.ValueName = mDefaultValuename Then
            If mListenerTable(i).ValueName = DefaultStudyValueNameStr Then
                mListenerTable(i).theListener.Notify ev
            End If
        End If
        If ev.ValueName = mListenerTable(i).ValueName Then
            mListenerTable(i).theListener.Notify ev
        End If
    End If
Next

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Remove( _
                ByVal listener As IStudyValueListener)
Dim i As Long
Dim lListener As IStudyValueListener

Const ProcName As String = "Remove"
On Error GoTo Err

For i = 0 To mNextListenerTableIndex - 1
    Set lListener = mListenerTable(i).theListener
    If lListener Is listener Then
        Set mListenerTable(i).theListener = Nothing
        mListenerTable(i).ValueName = ""
    End If
Next

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================










