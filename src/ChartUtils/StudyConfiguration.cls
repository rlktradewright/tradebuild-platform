VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyConfiguration"
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

Private Const ModuleName                As String = "StudyConfiguration"

Private Const ConfigSectionParameters                           As String = "Parameters"
Private Const ConfigSectionStudyConfigs                         As String = "StudyConfigs"
Private Const ConfigSectionStudyValueConfigurations             As String = "StudyValueConfigurations"
Private Const ConfigSectionStudyHorizontalRules                 As String = "StudyHorizontalRules"

Private Const ConfigSettingName                                 As String = ".Name"
Private Const ConfigSettingStudyLibraryName                     As String = ".StudyLibraryName"
Private Const ConfigSettingInstanceFullyQualifiedName           As String = ".InstanceFullyQualifiedName"
Private Const ConfigSettingInputValueNames                      As String = ".InputValueNames"
Private Const ConfigSettingChartRegionName                      As String = ".ChartRegionName"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyConfigs                       As StudyConfigurations

Private mParameters                         As Parameters
Private mStudyValueConfigurations           As StudyValueConfigurations
Private mStudyValueHandlers                 As StudyValueHandlers
Private mStudyHorizontalRules               As StudyHorizontalRules
Private mHorizontalRulesLineSeries          As LineSeries
Private mName                               As String
Private mStudyLibraryName                   As String
Private mInstanceFullyQualifiedName         As String
Private mStudy                              As IStudy
Private mUnderlyingStudy                    As IStudy
Private mInputValueNames()                  As String
Private mChartRegionName                    As String

Private mConfig                             As ConfigurationSection

Private mIdent                              As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mInputValueNames(0) As String
mInputValueNames(0) = DefaultStudyValueName
Set mStudyConfigs = New StudyConfigurations
Set mStudyValueConfigurations = New StudyValueConfigurations
Set mStudyValueHandlers = New StudyValueHandlers
Set mStudyHorizontalRules = New StudyHorizontalRules
mIdent = GenerateGUIDString
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

Public Property Let ChartRegionName(ByVal Value As String)
Const ProcName As String = "ChartRegionName"
On Error GoTo Err

mChartRegionName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartRegionName() As String
ChartRegionName = mChartRegionName
End Property

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

Public Property Let HorizontalRulesLineSeries( _
                ByVal Value As LineSeries)
Set mHorizontalRulesLineSeries = Value
End Property

Public Property Get HorizontalRulesLineSeries() As LineSeries
Set HorizontalRulesLineSeries = mHorizontalRulesLineSeries
End Property

Friend Property Get ID() As String
ID = mIdent
End Property

Public Property Let InputValueNames(ByRef Value() As String)
Const ProcName As String = "InputValueNames"
On Error GoTo Err

mInputValueNames = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InputValueNames() As String()
InputValueNames = mInputValueNames
End Property

Friend Property Let InstanceFullyQualifiedName(ByVal Value As String)
Const ProcName As String = "InstanceFullyQualifiedName"
On Error GoTo Err

mInstanceFullyQualifiedName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InstanceFullyQualifiedName() As String
InstanceFullyQualifiedName = mInstanceFullyQualifiedName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
On Error GoTo Err

mName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Parameters(ByVal Value As Parameters)
Const ProcName As String = "Parameters"
On Error GoTo Err

Set mParameters = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parameters() As Parameters
Set Parameters = mParameters
End Property

Public Property Let Study(ByVal Value As IStudy)
Set mStudy = Value
End Property

Public Property Get Study() As IStudy
Set Study = mStudy
End Property

Public Property Get StudyConfigurations() As StudyConfigurations
Set StudyConfigurations = mStudyConfigs
End Property

Public Property Get StudyHorizontalRules() As StudyHorizontalRules
Set StudyHorizontalRules = mStudyHorizontalRules
End Property

Public Property Let StudyLibraryName(ByVal Value As String)
Const ProcName As String = "StudyLibraryName"
On Error GoTo Err

mStudyLibraryName = Value
storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyLibraryName() As String
StudyLibraryName = mStudyLibraryName
End Property

Public Property Let StudyValueConfigurations(ByVal Value As StudyValueConfigurations)
Set mStudyValueConfigurations = Value
End Property

Public Property Get StudyValueConfigurations() As StudyValueConfigurations
Set StudyValueConfigurations = mStudyValueConfigurations
End Property

Friend Property Let StudyValueHandlers(ByVal Value As StudyValueHandlers)
Set mStudyValueHandlers = Value
End Property

Friend Property Get StudyValueHandlers() As StudyValueHandlers
Set StudyValueHandlers = mStudyValueHandlers
End Property

Public Property Let UnderlyingStudy(ByVal Value As IStudy)
Set mUnderlyingStudy = Value
End Property

Public Property Get UnderlyingStudy() As IStudy
Set UnderlyingStudy = mUnderlyingStudy
End Property

Public Property Get ValueSeries(ByVal Value As String) As IGraphicObjectSeries
Const ProcName As String = "ValueSeries"
On Error GoTo Err

Set ValueSeries = mStudyValueHandlers.Item(Value).ValueSeries

Exit Property

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then
    Set ValueSeries = Nothing
    Exit Function
End If
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ClearStudyConfigs()
Set mStudyConfigs = New StudyConfigurations
End Sub

Friend Sub ClearStudyValueHandlers()
Set mStudyValueHandlers = New StudyValueHandlers
End Sub

Public Function Clone() As StudyConfiguration
Const ProcName As String = "Clone"
On Error GoTo Err

Set Clone = New StudyConfiguration
Clone.ChartRegionName = mChartRegionName
Clone.InputValueNames = mInputValueNames
Clone.Name = mName
Clone.Parameters = mParameters.Clone
Clone.StudyLibraryName = mStudyLibraryName

Dim newHrs As StudyHorizontalRules
Set newHrs = Clone.StudyHorizontalRules

Dim hr As StudyHorizontalRule
For Each hr In mStudyHorizontalRules
    Dim newHr As StudyHorizontalRule
    Set newHr = newHrs.Add
    newHr.Color = hr.Color
    newHr.Style = hr.Style
    newHr.Thickness = hr.Thickness
    newHr.Y = hr.Y
Next

Clone.Study = mStudy

Dim newSvcs As StudyValueConfigurations
Set newSvcs = Clone.StudyValueConfigurations

Dim svc As StudyValueConfiguration
For Each svc In mStudyValueConfigurations
    Dim newSvc As StudyValueConfiguration
    Set newSvc = newSvcs.Add(svc.ValueName)
    newSvc.BarFormatterFactoryName = svc.BarFormatterFactoryName
    newSvc.BarFormatterLibraryName = svc.BarFormatterLibraryName
    newSvc.BarStyle = svc.BarStyle
    newSvc.ChartRegionName = svc.ChartRegionName
    newSvc.DataPointStyle = svc.DataPointStyle
    newSvc.IncludeInChart = svc.IncludeInChart
    newSvc.Layer = svc.Layer
    newSvc.LineStyle = svc.LineStyle
    newSvc.TextStyle = svc.TextStyle
Next

Clone.UnderlyingStudy = mUnderlyingStudy

Dim newScs As StudyConfigurations
Set newScs = Clone.StudyConfigurations

Dim sc As StudyConfiguration
For Each sc In mStudyConfigs
    Dim newSc As StudyConfiguration
    Set newSc = sc.Clone
    newScs.Add newSc
Next

' don't do a deep copy of the studyValueHandlers because it's
' immutable as far as non-Friend callers are concerned
Clone.StudyValueHandlers = mStudyValueHandlers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' Only called when a chart has been cleared, so no need to worry about whether regions are
' still in use etc
Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mStudyValueHandlers Is Nothing Then
    Dim svh As StudyValueHandler
    For Each svh In mStudyValueHandlers
        mStudy.RemoveStudyValueListener svh
    Next
    mStudyValueHandlers.Clear
End If

Dim sc As StudyConfiguration
For Each sc In mStudyConfigs
    sc.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config

If mConfig Is Nothing Then Exit Sub

mIdent = mConfig.InstanceQualifier
mName = mConfig.GetSetting(ConfigSettingName)
mStudyLibraryName = mConfig.GetSetting(ConfigSettingStudyLibraryName)
mInstanceFullyQualifiedName = mConfig.GetSetting(ConfigSettingInstanceFullyQualifiedName)
mInputValueNames = Split(mConfig.GetSetting(ConfigSettingInputValueNames), ",")
mChartRegionName = mConfig.GetSetting(ConfigSettingChartRegionName)

Set mParameters = LoadParametersFromConfig(mConfig.GetConfigurationSection(ConfigSectionParameters))

Set mStudyValueConfigurations = New StudyValueConfigurations
mStudyValueConfigurations.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionStudyValueConfigurations)

Set mStudyHorizontalRules = New StudyHorizontalRules
mStudyHorizontalRules.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudyHorizontalRules)

Set mStudyConfigs = New StudyConfigurations
mStudyConfigs.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudyConfigs)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Update()
Const ProcName As String = "Update"
On Error GoTo Err

If Not mStudyValueHandlers Is Nothing Then mStudyValueHandlers.Update
If Not mStudyConfigs Is Nothing Then mStudyConfigs.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStudyLibraryName, mStudyLibraryName
mConfig.SetSetting ConfigSettingInstanceFullyQualifiedName, mInstanceFullyQualifiedName
mConfig.SetSetting ConfigSettingInputValueNames, Join(mInputValueNames, ",")
mConfig.SetSetting ConfigSettingChartRegionName, mChartRegionName

mParameters.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionParameters)
mStudyValueConfigurations.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudyValueConfigurations)
mStudyHorizontalRules.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudyHorizontalRules)

mStudyConfigs.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudyConfigs)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
