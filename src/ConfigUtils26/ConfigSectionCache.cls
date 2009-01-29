VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigSectionCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ConfigSectionCache"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigSections                             As New Collection

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

Friend Function Add( _
                ByVal pConfigSection As ConfigurationSection) As ConfigurationSection
mConfigSections.Add pConfigSection, pConfigSection.path
Set Add = pConfigSection
End Function

Friend Function Item( _
                ByVal path As String) As ConfigurationSection
On Error Resume Next
Set Item = mConfigSections.Item(path)
End Function

Friend Sub Remove( _
                ByVal pConfigSection As ConfigurationSection)
Dim i As Long
Dim section As ConfigurationSection

For i = mConfigSections.Count - 1 To 1 Step -1
    Set section = mConfigSections(i)
    If InStr(1, section.path, pConfigSection.path) = 1 Then
        mConfigSections.Remove i
    End If
Next
End Sub

Friend Sub RemoveChildren( _
                ByVal pConfigSection As ConfigurationSection)
Dim i As Long
Dim section As ConfigurationSection

For i = mConfigSections.Count - 1 To 1 Step -1
    Set section = mConfigSections(i)
    If InStr(1, section.path, pConfigSection.path) = 1 And Not section Is pConfigSection Then
        mConfigSections.Remove i
    End If
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


