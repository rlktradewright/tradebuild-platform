Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const RegionNameCustom As String = "$custom"
Public Const RegionNameDefault As String = "$default"
Public Const RegionNamePrice As String = "Price"
Public Const RegionNameVolume As String = "Volume"

Public Const LB_SETHORZEXTENT = &H194

Public Const TaskTypeStartStudy As Long = 1
Public Const TaskTypeReplayBars As Long = 2
Public Const TaskTypeAddValueListener As Long = 3

Public Const ErroredFieldColor As Long = &HD0CAFA

Public Const PositiveChangeBackColor As Long = &HB7E43
Public Const NegativeChangebackColor As Long = &H4444EB

Public Const PositiveProfitColor As Long = &HB7E43
Public Const NegativeProfitColor As Long = &H4444EB

Public Const IncreasedValueColor As Long = &HB7E43
Public Const DecreasedValueColor As Long = &H4444EB

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultStudyConfigurations As Collection

Private mStudyPickerForm As fStudyPicker

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'Public Sub cleverComboChanged( _
'                ByVal combo As ComboBox)
'Dim l As Long
'Dim selItem As Long
'
'l = Len(combo.Text)
'If combo.Text <> "" Then
'    selItem = SendMessageW(combo.hWnd, CB_FINDSTRING, -1, StrPtr(combo.Text))
'    If selItem <> CB_ERR Then
'        SendMessageW combo.hWnd, CB_SETCURSEL, selItem, 0
'        combo.SelStart = l
'        combo.SelLength = Len(combo.Text) - l
'    End If
'End If
'
'End Sub
'
'Public Sub cleverComboKeyDown( _
'                ByVal combo As ComboBox, _
'                KeyCode As Integer)
'Dim i As Long
'Dim posn As Long
'If KeyCode = vbKeyUp Then
'    posn = 0
'    If combo.ListIndex = -1 Then
'        For i = 0 To combo.ListCount - 1
'            If StrComp(combo.Text, combo.List(i), vbTextCompare) = 0 Then
'                combo.ListIndex = i
'                Exit Sub
'            End If
'            If StrComp(combo.List(i), combo.Text, vbTextCompare) < 0 Then posn = i
'        Next
'        If posn = combo.ListCount - 1 Then
'            KeyCode = 0
'        Else
'            posn = posn + 1
'        End If
'        combo.ListIndex = posn
'    End If
'ElseIf KeyCode = vbKeyDown Then
'    posn = combo.ListCount - 1
'    If combo.ListIndex = -1 Then
'        For i = combo.ListCount - 1 To 0 Step -1
'            If StrComp(combo.Text, combo.List(i), vbTextCompare) = 0 Then
'                combo.ListIndex = i
'                Exit Sub
'            End If
'            If StrComp(combo.Text, combo.List(i), vbTextCompare) < 0 Then posn = i
'        Next
'        If posn = 0 Then
'            KeyCode = 0
'        Else
'            posn = posn - 1
'        End If
'        combo.ListIndex = posn
'    End If
'End If
'End Sub
'
'Public Sub cleverComboKeyPress( _
'                ByVal combo As ComboBox, _
'                KeyAscii As Integer)
'
'If KeyAscii = vbKeyBack Then
'    If combo.SelStart <> 0 And combo.SelLength <> 0 Then
'        Dim l As Long
'        l = combo.SelLength
'        combo.SelStart = combo.SelStart - 1
'        combo.SelLength = l + 1
'    End If
'End If
'End Sub
'
'Public Sub cleverComboValidate( _
'                ByVal combo As ComboBox, _
'                Cancel As Boolean)
'Dim i As Long
'If combo.Text = "" Then Exit Sub
'If combo.ListIndex = -1 Then
'    For i = 0 To combo.ListCount - 1
'        If combo.Text = combo.List(i) Then
'            combo.ListIndex = i
'            Exit Sub
'        End If
'    Next
'    Cancel = True
'End If
'End Sub

Public Sub filterNonNumericKeyPress(ByRef KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
End If
End Sub

Public Function isPrice( _
                ByVal value As String, _
                ByVal ticksize As Double) As Boolean
Dim theVal As Double

On Error GoTo Err

If IsNumeric(value) Then
    theVal = value
    If theVal > 0 And _
        Int(theVal / ticksize) * ticksize = theVal _
    Then
        isPrice = True
    End If
End If

Exit Function

Err:
If Err.Number <> VBErrorCodes.VbErrOverflow Then Err.Raise Err.Number
End Function

Public Function loadDefaultStudyConfiguration( _
                ByVal name As String, _
                ByVal spName As String) As StudyConfiguration
Dim sc As StudyConfiguration
If mDefaultStudyConfigurations Is Nothing Then
    Set loadDefaultStudyConfiguration = Nothing
Else
    On Error Resume Next
    Set sc = mDefaultStudyConfigurations.item(calcDefaultStudyKey(name, spName))
    On Error GoTo 0
    If Not sc Is Nothing Then Set loadDefaultStudyConfiguration = sc.Clone
End If
End Function

Public Sub notImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

Public Sub showStudyPicker( _
                ByVal chartMgr As chartManager)
If mStudyPickerForm Is Nothing Then
    Set mStudyPickerForm = New fStudyPicker
End If
mStudyPickerForm.initialise chartMgr
mStudyPickerForm.Show vbModeless
End Sub

Public Sub syncStudyPicker( _
                ByVal chartMgr As chartManager)
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise chartMgr
End Sub

Public Sub unsyncStudyPicker()
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise Nothing
End Sub

Public Sub updateDefaultStudyConfiguration( _
                ByVal value As StudyConfiguration)
Dim sc As StudyConfiguration

If mDefaultStudyConfigurations Is Nothing Then
    Set mDefaultStudyConfigurations = New Collection
End If
On Error Resume Next
mDefaultStudyConfigurations.remove calcDefaultStudyKey(value.name, value.StudyLibraryName)
On Error GoTo 0

Set sc = value.Clone
sc.underlyingStudy = Nothing
mDefaultStudyConfigurations.add sc, calcDefaultStudyKey(value.name, value.StudyLibraryName)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcDefaultStudyKey( _
                ByVal studyName As String, _
                ByVal StudyLibraryName As String) As String
calcDefaultStudyKey = "$$" & studyName & "$$" & StudyLibraryName & "$$"
End Function

