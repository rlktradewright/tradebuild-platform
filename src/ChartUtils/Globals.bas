Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
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

Public Const ProjectName                        As String = "ChartUtils27"
Private Const ModuleName                        As String = "Globals"

Private Const ConfigSectionDefaultStudyConfig   As String = "DefaultStudyConfig"

Public Const OneMicroSecond                     As Double = 1.15740740740741E-11

Public Const RegionNameCustom                   As String = "$custom"
Public Const RegionNameDefault                  As String = "$default"
Public Const RegionNameUnderlying               As String = "$underlying"
Public Const RegionNamePrice                    As String = "Price"
Public Const RegionNameVolume                   As String = "Volume"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultStudyConfigurations         As EnumerableCollection

Private mConfig                             As ConfigurationSection

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

Public Property Get gLogger() As FormattingLogger
Static lLogger As FormattingLogger
If lLogger Is Nothing Then Set lLogger = CreateFormattingLogger("chartutils", ProjectName)
Set gLogger = lLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gGetDefaultStudyConfiguration( _
                ByVal Name As String, _
                ByVal StudyLibManager As StudyLibraryManager, _
                ByVal studyLibName As String) As StudyConfiguration
Const ProcName As String = "gGetDefaultStudyConfiguration"
On Error GoTo Err

If mDefaultStudyConfigurations Is Nothing Then Set mDefaultStudyConfigurations = New Collection

On Error Resume Next
Dim studyConfig As StudyConfiguration
Set studyConfig = mDefaultStudyConfigurations.Item(calcDefaultStudyKey(Name, studyLibName))
On Error GoTo Err

If Not studyConfig Is Nothing Then
    Set gGetDefaultStudyConfiguration = studyConfig.Clone
    ' ensure that each instance of the default study config has its own
    ' StudyValueHandlers
    gGetDefaultStudyConfiguration.StudyValueHandlers = New StudyValueHandlers
Else
    'no default Study config currently exists so we'll create one from the Study definition
    Dim sd As StudyDefinition
    Set sd = StudyLibManager.GetStudyDefinition(Name, studyLibName)

    Set studyConfig = New StudyConfiguration
    studyConfig.Name = Name
    studyConfig.StudyLibraryName = studyLibName

    Select Case sd.DefaultRegion
        Case StudyDefaultRegions.StudyDefaultRegionNone
            studyConfig.ChartRegionName = RegionNameUnderlying
        Case StudyDefaultRegions.StudyDefaultRegionCustom
            studyConfig.ChartRegionName = RegionNameCustom
        Case StudyDefaultRegions.StudyDefaultRegionUnderlying
            studyConfig.ChartRegionName = RegionNameUnderlying
        Case Else
            studyConfig.ChartRegionName = RegionNameUnderlying
    End Select

    studyConfig.Parameters = StudyLibManager.GetStudyDefaultParameters(Name, studyLibName)
    
    Dim InputValueNames() As String
    ReDim InputValueNames(sd.StudyInputDefinitions.Count - 1) As String
    
    InputValueNames(0) = DefaultStudyValueName
    If sd.StudyInputDefinitions.Count > 1 Then
        Dim i As Long
        For i = 2 To sd.StudyInputDefinitions.Count
            InputValueNames(i - 1) = sd.StudyInputDefinitions.Item(i).Name
        Next
    End If
    studyConfig.InputValueNames = InputValueNames

    Dim studyValueDef As StudyValueDefinition
    Dim studyValueConfig As StudyValueConfiguration
    
    For Each studyValueDef In sd.StudyValueDefinitions
        Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(studyValueDef.Name)

        studyValueConfig.IncludeInChart = studyValueDef.IncludeInChart
        Select Case studyValueDef.ValueMode
            Case StudyValueModes.ValueModeNone
                studyValueConfig.DataPointStyle = studyValueDef.ValueStyle
                
            Case StudyValueModes.ValueModeLine
                studyValueConfig.LineStyle = studyValueDef.ValueStyle

            Case StudyValueModes.ValueModeBar
                studyValueConfig.BarStyle = studyValueDef.ValueStyle

            Case StudyValueModes.ValueModeText
                studyValueConfig.TextStyle = studyValueDef.ValueStyle

        End Select

        Select Case studyValueDef.DefaultRegion
            Case StudyValueDefaultRegions.StudyValueDefaultRegionNone
                studyValueConfig.ChartRegionName = RegionNameDefault
            Case StudyValueDefaultRegions.StudyValueDefaultRegionCustom
                studyValueConfig.ChartRegionName = RegionNameCustom
            Case StudyValueDefaultRegions.StudyValueDefaultRegionDefault
                studyValueConfig.ChartRegionName = RegionNameDefault
            Case StudyValueDefaultRegions.StudyValueDefaultRegionUnderlying
                studyValueConfig.ChartRegionName = RegionNameUnderlying
        End Select

    Next
    gSetDefaultStudyConfiguration studyConfig
    Set gGetDefaultStudyConfiguration = studyConfig
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub gLoadDefaultStudyConfigurationsFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "gLoadDefaultStudyConfigurationsFromConfig"
On Error GoTo Err

Set mConfig = config

Set mDefaultStudyConfigurations = New EnumerableCollection

Dim scSect As ConfigurationSection
For Each scSect In mConfig
    Dim sc As StudyConfiguration
    Set sc = New StudyConfiguration
    sc.LoadFromConfig scSect
    
    Dim lKey As String: lKey = calcDefaultStudyKey(sc.Name, sc.StudyLibraryName)
    If mDefaultStudyConfigurations.Contains(lKey) Then
        gLogger.Log "Config file contains more than one default configuration for Study " & sc.Name & "(" & sc.StudyLibraryName & ")", ProcName, ModuleName
    Else
        mDefaultStudyConfigurations.Add sc, lKey
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSetDefaultStudyConfiguration( _
                ByVal Value As StudyConfiguration)
Const ProcName As String = "gSetDefaultStudyConfiguration"
On Error GoTo Err

If mDefaultStudyConfigurations Is Nothing Then
    Set mDefaultStudyConfigurations = New EnumerableCollection
End If

Dim key As String
key = calcDefaultStudyKey(Value.Name, Value.StudyLibraryName)

If mDefaultStudyConfigurations.Contains(key) Then
    mDefaultStudyConfigurations.Item(key).RemoveFromConfig
    mDefaultStudyConfigurations.Remove key
End If

Dim sc As StudyConfiguration
Set sc = Value.Clone
mDefaultStudyConfigurations.Add sc, key
If Not mConfig Is Nothing Then sc.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionDefaultStudyConfig & "(" & sc.ID & ")")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcDefaultStudyKey(ByVal studyName As String, ByVal StudyLibraryName As String) As String
calcDefaultStudyKey = "$$" & studyName & "$$" & StudyLibraryName & "$$"
End Function


