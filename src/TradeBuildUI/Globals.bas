Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const CustomRegionName As String = "$custom"
Public Const PriceRegionName As String = "$price"
Public Const VolumeRegionName As String = "$volume"

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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Declare Function SendMessageByNum Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'================================================================================
' Member variables
'================================================================================

Private mDefaultStudyConfigurations As Collection

Private mTaskManager As TradeBuild.TaskManager

Private mStudyPickerForm As fStudyPicker

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

Public Sub filterNonNumericKeyPress(ByRef KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
End If
End Sub

Public Function isInteger( _
                ByVal value As String, _
                Optional ByVal minValue As Long = 0, _
                Optional ByVal maxValue As Long = &H7FFFFFFF) As Boolean
Dim quantity As Long

On Error GoTo err

If IsNumeric(value) Then
    quantity = CLng(value)
    If CDbl(value) - quantity = 0 Then
        If quantity >= minValue And quantity <= maxValue Then
            isInteger = True
        End If
    End If
End If
                
Exit Function

err:
If err.Number <> TradeBuild.ErrorCodes.ErrOverflow Then err.Raise err.Number
End Function

Public Function isPrice( _
                ByVal value As String, _
                ByVal ticksize As Double) As Boolean
Dim theVal As Double

On Error GoTo err

If IsNumeric(value) Then
    theVal = value
    If theVal > 0 And _
        Int(theVal / ticksize) * ticksize = theVal _
    Then
        isPrice = True
    End If
End If

Exit Function

err:
If err.Number <> TradeBuild.ErrorCodes.ErrOverflow Then err.Raise err.Number
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
    If Not sc Is Nothing Then Set loadDefaultStudyConfiguration = sc.clone
End If
End Function

Public Sub notImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

Public Sub showStudyPicker( _
                ByVal pTicker As TradeBuild.ticker, _
                ByVal chart As TradeBuildChart)
If mStudyPickerForm Is Nothing Then
    Set mStudyPickerForm = New fStudyPicker
End If
mStudyPickerForm.initialise chart, pTicker
mStudyPickerForm.Show vbModeless
End Sub

Public Sub syncStudyPicker( _
                ByVal pTicker As TradeBuild.ticker, _
                ByVal chart As TradeBuildChart)
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise chart, pTicker
End Sub

Public Function startTask( _
                ByVal target As Task, _
                Optional ByVal name As String, _
                Optional ByVal data As Variant) As TaskCompletion
If mTaskManager Is Nothing Then
    Set mTaskManager = New TradeBuild.TaskManager
End If
Set startTask = mTaskManager.startTask(target, name, data)
End Function

Public Sub unsyncStudyPicker()
If mStudyPickerForm Is Nothing Then Exit Sub
mStudyPickerForm.initialise Nothing, Nothing
End Sub

Public Sub updateDefaultStudyConfiguration( _
                ByVal value As StudyConfiguration)
Dim sc As StudyConfiguration

If mDefaultStudyConfigurations Is Nothing Then
    Set mDefaultStudyConfigurations = New Collection
End If
On Error Resume Next
mDefaultStudyConfigurations.remove calcDefaultStudyKey(value.name, value.serviceProviderName)
On Error GoTo 0

Set sc = value.clone
sc.underlyingStudyId = ""
mDefaultStudyConfigurations.add sc, calcDefaultStudyKey(value.name, value.serviceProviderName)
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcDefaultStudyKey( _
                ByVal studyName As String, _
                ByVal serviceProviderName As String) As String
calcDefaultStudyKey = "$$" & studyName & "$$" & serviceProviderName & "$$"
End Function

