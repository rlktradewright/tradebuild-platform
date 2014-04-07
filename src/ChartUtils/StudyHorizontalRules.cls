VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyHorizontalRules"
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

Private Const ModuleName                As String = "StudyHorizontalRules"

Private Const ConfigSectionHorizontalRule                   As String = "HorizontalRule"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyHorizRules As Collection

Private mConfig As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyHorizRules = New Collection
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

Friend Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is mConfig Then Exit Property
Set mConfig = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mStudyHorizRules.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add() As StudyHorizontalRule
Const ProcName As String = "Add"
On Error GoTo Err

Set Add = New StudyHorizontalRule
mStudyHorizRules.Add Add
storeRuleSettings Add

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Clone() As StudyHorizontalRules
Dim hr As StudyHorizontalRule
Dim newHr As StudyHorizontalRule
Const ProcName As String = "Clone"
On Error GoTo Err

Set Clone = New StudyHorizontalRules
For Each hr In mStudyHorizRules
    newHr = Clone.Add
    newHr.Color = hr.Color
    newHr.Style = hr.Style
    newHr.Thickness = hr.Thickness
    newHr.Y = hr.Y
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Long) As StudyHorizontalRule
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mStudyHorizRules.Item(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim shrSect As ConfigurationSection
Dim shr As StudyHorizontalRule

Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

For Each shrSect In mConfig
    Set shr = New StudyHorizontalRule
    shr.LoadFromConfig shrSect
    mStudyHorizRules.Add shr
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mStudyHorizRules.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Dim shr As StudyHorizontalRule
Const ProcName As String = "storeSettings"
On Error GoTo Err

For Each shr In mStudyHorizRules
    storeRuleSettings shr
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeRuleSettings( _
                ByVal shr As StudyHorizontalRule)
Const ProcName As String = "storeRuleSettings"
On Error GoTo Err

If Not mConfig Is Nothing Then
    shr.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionHorizontalRule & "(" & GenerateGUIDString & ")")
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



