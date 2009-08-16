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
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

End Property

Public Property Get count() As Long
count = mStudyConfigurations.count
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
For Each studyConfig In mStudyConfigurations
    studyConfig.Finish
Next
End Sub

Public Function Item( _
                ByVal index As Variant) As StudyConfiguration
Attribute Item.VB_UserMemId = 0
Set Item = mStudyConfigurations.Item(index)
End Function

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim scSect As ConfigurationSection
Dim sc As StudyConfiguration

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

For Each scSect In mConfig
    Set sc = New StudyConfiguration
    sc.LoadFromConfig scSect
    mStudyConfigurations.Add sc, sc.ID
Next
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mStudyConfigurations.[_NewEnum]
End Function

Public Sub Remove( _
                ByVal studyConfig As StudyConfiguration)
mStudyConfigurations.Remove studyConfig.ID
studyConfig.RemoveFromConfig
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Dim sc As StudyConfiguration
For Each sc In mStudyConfigurations
    storeStudyConfigSettings sc
Next
End Sub

Private Sub storeStudyConfigSettings( _
                ByVal studyConfig As StudyConfiguration)
If Not mConfig Is Nothing Then
    studyConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudyConfig & "(" & studyConfig.ID & ")")
End If
End Sub



