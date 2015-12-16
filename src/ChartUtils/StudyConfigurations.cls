VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                As String = "StudyConfigurations"


Private Const ConfigSectionStudyConfig                  As String = "StudyConfig"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyConfigurations As Collection

Private mConfig                                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyConfigurations = New Collection
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

Public Property Let ConfigurationSection( _
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

Count = mStudyConfigurations.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                ByVal studyConfig As StudyConfiguration)
On Error Resume Next
' the studyconfig may already be in the collection because it was Added when
' recreating this object from the config
mStudyConfigurations.Add studyConfig, studyConfig.ID
If Err.Number = 0 Then storeStudyConfigSettings studyConfig
End Sub

' Only called when a chart has been cleared, so no need to worry about whether regions are
' still in use etc
Friend Sub Finish()
Dim studyConfig As StudyConfiguration
Const ProcName As String = "Finish"

On Error GoTo Err

For Each studyConfig In mStudyConfigurations
    studyConfig.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item( _
                ByVal index As Variant) As StudyConfiguration
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mStudyConfigurations.Item(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim scSect As ConfigurationSection
Dim sc As StudyConfiguration

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

For Each scSect In mConfig
    Set sc = New StudyConfiguration
    sc.LoadFromConfig scSect
    mStudyConfigurations.Add sc, sc.ID
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyConfigurations.[_NewEnum]
End Function

Public Sub Remove( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "Remove"

On Error GoTo Err

mStudyConfigurations.Remove studyConfig.ID
studyConfig.RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Update()
Const ProcName As String = "Update"
On Error GoTo Err

Dim lSc As StudyConfiguration
For Each lSc In mStudyConfigurations
    lSc.Update
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Dim sc As StudyConfiguration
Const ProcName As String = "storeSettings"

On Error GoTo Err

For Each sc In mStudyConfigurations
    storeStudyConfigSettings sc
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeStudyConfigSettings( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "storeStudyConfigSettings"

On Error GoTo Err

If Not mConfig Is Nothing Then
    studyConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudyConfig & "(" & studyConfig.ID & ")")
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



