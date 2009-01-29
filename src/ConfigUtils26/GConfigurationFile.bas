Attribute VB_Name = "GConfigurationFile"
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

Public Const ProjectName                            As String = "ConfigUtils26"
Private Const ModuleName                            As String = "GConfigurationFile"

Public Const AttributeNameAppConfigDefault          As String = "Default"
Public Const AttributeNameName                      As String = "Name"
Public Const AttributeNamePrivate                   As String = "__Private"
Public Const AttributeNameRenderer                  As String = "__Renderer"
Public Const AttributeNameType                      As String = "__Type"

Public Const AttributeValueFalse                    As String = "False"
Public Const AttributeValueTrue                     As String = "True"
Public Const AttributeValueTypeBoolean              As String = "Boolean"
Public Const AttributeValueTypeSelection            As String = "Selection"

Public Const ConfigNameAppConfig                    As String = "AppConfig"
Public Const ConfigNameAppConfigs                   As String = "AppConfigs"
Public Const ConfigNameSelection                    As String = "__Selection"
Public Const ConfigNameSelections                   As String = "__Selections"
Public Const ConfigNameTradeBuild                   As String = "TradeBuild"

Public Const ConfigNodeServiceProviders             As String = "Service Providers"
Public Const ConfigNodeStudyLibraries               As String = "Study Libraries"

Public Const DefaultAppConfigName                   As String = "Default config"

Public Const SectionPathSeparator                   As String = "/"
Public Const AttributeNameSeparator                 As String = "&"
Public Const ValueNameSeparator                     As String = "."
Public Const RootSectionName                        As String = "Configuration"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigPaths                                As New Collection

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

Public Function gGetConfigPath( _
                ByVal path As String) As ConfigurationPath
On Error Resume Next
Set gGetConfigPath = mConfigPaths(path)
On Error GoTo 0

If gGetConfigPath Is Nothing Then
    Set gGetConfigPath = New ConfigurationPath
    gGetConfigPath.Initialise path
    mConfigPaths.Add gGetConfigPath, path
End If
End Function

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("configutils.log")
Set gLogger = lLogger
End Property

Public Property Get gRegExp() As RegExp
Static lRegexp As RegExp
If lRegexp Is Nothing Then Set lRegexp = New RegExp
Set gRegExp = lRegexp
End Property

Public Function gXMLEncode(ByVal value As String) As String
gXMLEncode = Replace(Replace(Replace(Replace(Replace(value, "&", "&amp;"), "'", "&apos;"), """", "&quot;"), "<", "&lt;"), ">", "&gt;")
End Function

Public Function gXMLDecode(ByVal value As String) As String
gXMLDecode = Replace(Replace(Replace(Replace(value, "&amp;", "&"), "&apos;", "'"), "&lt;", "<"), "&gt;", ">")
End Function

'@================================================================================
' Helper Functions
'@================================================================================


