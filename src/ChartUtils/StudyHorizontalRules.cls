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
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

End Property

Public Property Get count() As Long
count = mStudyHorizRules.count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add() As StudyHorizontalRule
Set Add = New StudyHorizontalRule
mStudyHorizRules.Add Add
storeRuleSettings Add
End Function

Public Function Clone() As StudyHorizontalRules
Dim hr As StudyHorizontalRule
Dim newHr As StudyHorizontalRule
Set Clone = New StudyHorizontalRules
For Each hr In mStudyHorizRules
    newHr = Clone.Add
    newHr.Color = hr.Color
    newHr.Style = hr.Style
    newHr.Thickness = hr.Thickness
    newHr.Y = hr.Y
Next
End Function

Public Function Item( _
                ByVal index As Long) As StudyHorizontalRule
Attribute Item.VB_UserMemId = 0
Set Item = mStudyHorizRules.Item(index)
End Function

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim shrSect As ConfigurationSection
Dim shr As StudyHorizontalRule

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

For Each shrSect In mConfig
    Set shr = New StudyHorizontalRule
    shr.LoadFromConfig shrSect
    mStudyHorizRules.Add shr
Next
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyHorizRules.[_NewEnum]
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Dim shr As StudyHorizontalRule
For Each shr In mStudyHorizRules
    storeRuleSettings shr
Next
End Sub

Private Sub storeRuleSettings( _
                ByVal shr As StudyHorizontalRule)
If Not mConfig Is Nothing Then
    shr.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionHorizontalRule & "(" & GenerateGUIDString & ")")
End If
End Sub



