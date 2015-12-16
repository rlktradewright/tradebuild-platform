VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
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

Private Const ModuleName                As String = "StudyValueHandlers"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyValueHandlers As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyValueHandlers = New Collection
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mStudyValueHandlers.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal chartMgr As ChartManager, _
                ByVal Region As ChartRegion, _
                ByVal pStudy As IStudy, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition, _
                ByVal UpdatePerTick As Boolean) As StudyValueHandler
Const ProcName As String = "Add"
On Error GoTo Err

Set Add = New StudyValueHandler
Add.Initialise chartMgr, _
                Region, _
                pStudy, _
                studyValueConfig, _
                studyValueDef, _
                UpdatePerTick
mStudyValueHandlers.Add Add, studyValueConfig.ValueName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Set mStudyValueHandlers = New Collection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item( _
                ByVal index As Variant) As StudyValueHandler
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mStudyValueHandlers.Item(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mStudyValueHandlers.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Update()
Const ProcName As String = "Update"
On Error GoTo Err

Dim lSvh As StudyValueHandler
For Each lSvh In mStudyValueHandlers
    lSvh.Update
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





