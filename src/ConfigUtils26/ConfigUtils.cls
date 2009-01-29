VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ModuleName                            As String = "ConfigUtils"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function CreateConfigurationFile( _
                ByVal pConfigFile As configFile, _
                ByVal filename As String, _
                ByVal permittedServiceProviders As ServiceProviderRoles, _
                ByVal includeDefaultStudyLibrary As Boolean) As ConfigurationFile
Dim failpoint As Long
On Error GoTo Err

Set CreateConfigurationFile = New ConfigurationFile
CreateConfigurationFile.Initialise pConfigFile, _
                                    filename, _
                                    permittedServiceProviders, _
                                    includeDefaultStudyLibrary

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "CreateConfigurationFile" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function IsRelativeConfigPath( _
                ByVal configPath As String) As Boolean
Dim cp As ConfigurationPath
Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(configPath)
IsRelativeConfigPath = cp.IsRelative

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IsRelativeConfigPath" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function IsValidConfigAttributePath( _
                ByVal configPath As String) As Boolean
Dim cp As ConfigurationPath
Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(configPath)
IsValidConfigAttributePath = cp.IsValidAttributePath

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IsValidConfigAttributePath" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function IsValidConfigValuePath( _
                ByVal configPath As String) As Boolean
Dim cp As ConfigurationPath
Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(configPath)
IsValidConfigValuePath = cp.IsValidValuePath

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IsValidConfigValuePath" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function IsValidConfigSectionPath( _
                ByVal configPath As String) As Boolean
Dim cp As ConfigurationPath
Dim failpoint As Long
On Error GoTo Err

Set cp = gGetConfigPath(configPath)
IsValidConfigSectionPath = cp.IsValidSectionPath

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "IsValidConfigSectionPath" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

'@================================================================================
' Helper Functions
'@================================================================================


