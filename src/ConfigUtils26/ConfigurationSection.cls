VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "ConfigurationSection"

'@================================================================================
' Member variables
'@================================================================================

Private mConfig                                     As ConfigItem

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

Public Function AddConfigurationSection( _
                ByVal relativeSectionPath As String, _
                Optional ByVal nameAttribute As String) As ConfigurationSection

End Function

Public Function GetConfigurationSection( _
                ByVal relativeSectionPath As String, _
                Optional ByVal nameAttribute As String) As ConfigurationSection

End Function

Public Function GetAttribute( _
                ByVal name As String, _
                Optional ByVal defaultValue As String) As String
                
End Function

Public Function GetSetting( _
                ByVal relativeSettingPath As String) As String

End Function

Friend Function Initialise( _
                ByVal config As ConfigItem)
Set mConfig = config
End Function

Public Sub RemoveAttribute( _
                ByVal name As String)
                
End Sub

Public Sub RemoveConfigurationSection( _
                ByVal relativeSectionPath As String, _
                Optional ByVal nameAttribute As String)

End Sub

Public Sub RemoveSetting( _
                ByVal relativeSettingPath As String)
                
End Sub

Public Function SetAttribute( _
                ByVal name As String, _
                ByVal value As String) As String
                
End Function

Public Function SetSetting( _
                ByVal relativeSettingPath As String, _
                ByVal value As String, _
                Optional ByVal defaultValue) As String

End Function

'@================================================================================
' Helper Functions
'@================================================================================



