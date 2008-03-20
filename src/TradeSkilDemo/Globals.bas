Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const IncreasedValueColor As Long = &HB7E43
Public Const DecreasedValueColor As Long = &H4444EB

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Global object references
'================================================================================

Public gLogger As Logger

'================================================================================
' External function declarations
'================================================================================

'================================================================================
' Variables
'================================================================================

Private mStudyPickerForm As fStudyPicker

'================================================================================
' Procedures
'================================================================================

Public Sub gShowStudyPicker( _
                ByVal chartMgr As chartManager, _
                ByVal title As String)
    If mStudyPickerForm Is Nothing Then
        Set mStudyPickerForm = New fStudyPicker
    End If
    mStudyPickerForm.initialise chartMgr, title
    mStudyPickerForm.Show vbModeless
End Sub

Public Sub gSyncStudyPicker( _
                ByVal chartMgr As chartManager, _
                ByVal title As String)
    If mStudyPickerForm Is Nothing Then Exit Sub
    mStudyPickerForm.initialise chartMgr, title
End Sub

Public Sub gUnsyncStudyPicker()
    If mStudyPickerForm Is Nothing Then Exit Sub
    mStudyPickerForm.initialise Nothing, "Study picker"
End Sub


