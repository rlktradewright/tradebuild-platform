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
Private mStudy                              As Study
Private mUnderlyingStudy                    As Study
Private mInputValueNames()                  As String
Private mChartRegionName                    As String

Private mConfig                             As ConfigurationSection

Private mIdent                              As String

Private mBars                               As Bars

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

Public Property Let ChartRegionName(ByVal value As String)
Const ProcName As String = "ChartRegionName"
Dim failpoint As String
On Error GoTo Err

mChartRegionName = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartRegionName() As String
ChartRegionName = mChartRegionName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalRulesLineSeries( _
                ByVal value As LineSeries)
Set mHorizontalRulesLineSeries = value
End Property

Public Property Get HorizontalRulesLineSeries() As LineSeries
Set HorizontalRulesLineSeries = mHorizontalRulesLineSeries
End Property

Friend Property Get ID() As String
ID = mIdent
End Property

Public Property Let InputValueNames(ByRef value() As String)
Const ProcName As String = "InputValueNames"
Dim failpoint As String
On Error GoTo Err

mInputValueNames = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get InputValueNames() As String()
InputValueNames = mInputValueNames
End Property

Friend Property Let InstanceFullyQualifiedName(ByVal value As String)
Const ProcName As String = "InstanceFullyQualifiedName"
Dim failpoint As String
On Error GoTo Err

mInstanceFullyQualifiedName = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get InstanceFullyQualifiedName() As String
InstanceFullyQualifiedName = mInstanceFullyQualifiedName
End Property

Public Property Let Name(ByVal value As String)
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

mName = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Parameters(ByVal value As Parameters)
Const ProcName As String = "Parameters"
Dim failpoint As String
On Error GoTo Err

Set mParameters = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Parameters() As Parameters
Set Parameters = mParameters
End Property

Public Property Let StudyLibraryName(ByVal value As String)
Const ProcName As String = "StudyLibraryName"
Dim failpoint As String
On Error GoTo Err

mStudyLibraryName = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get StudyLibraryName() As String
StudyLibraryName = mStudyLibraryName
End Property

Public Property Get StudyHorizontalRules() As StudyHorizontalRules
Set StudyHorizontalRules = mStudyHorizontalRules
End Property

Public Property Let Study(ByVal value As Study)
Dim lBarStudy As BarStudy

Const ProcName As String = "Study"
Dim failpoint As String
On Error GoTo Err

Set mStudy = value
If mStudy Is Nothing Then Exit Property

If TypeOf mStudy Is BarStudy Then
    Set lBarStudy = mStudy
    Set mBars = lBarStudy.Bars
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
    
End Property

Public Property Get Study() As Study
Set Study = mStudy
End Property

Public Property Get StudyConfigurations() As StudyConfigurations
Set StudyConfigurations = mStudyConfigs
End Property

Public Property Get StudyValueConfigurations() As StudyValueConfigurations
Set StudyValueConfigurations = mStudyValueConfigurations
End Property

Friend Property Let StudyValueHandlers(ByVal value As StudyValueHandlers)
Set mStudyValueHandlers = value
End Property

Friend Property Get StudyValueHandlers() As StudyValueHandlers
Set StudyValueHandlers = mStudyValueHandlers
End Property

Public Property Let UnderlyingStudy(ByVal value As Study)
Set mUnderlyingStudy = value
End Property

Public Property Get UnderlyingStudy() As Study
Set UnderlyingStudy = mUnderlyingStudy
End Property

Public Property Get ValueSeries(ByVal value As String) As IGraphicObjectSeries
Const ProcName As String = "ValueSeries"
Dim failpoint As String
On Error GoTo Err

Set ValueSeries = mStudyValueHandlers.Item(value).ValueSeries

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Clone() As StudyConfiguration
Const ProcName As String = "Clone"
Dim failpoint As String
On Error GoTo Err

Set Clone = New StudyConfiguration
Clone.ChartRegionName = mChartRegionName
Clone.InputValueNames = mInputValueNames
Clone.Name = mName
Clone.Parameters = mParameters.Clone
Clone.StudyLibraryName = mStudyLibraryName

Dim hr As StudyHorizontalRule
Dim newHr As StudyHorizontalRule
Dim newHrs As StudyHorizontalRules
Set newHrs = Clone.StudyHorizontalRules
For Each hr In mStudyHorizontalRules
    Set newHr = newHrs.Add
    newHr.Color = hr.Color
    newHr.Style = hr.Style
    newHr.Thickness = hr.Thickness
    newHr.Y = hr.Y
Next

Clone.Study = mStudy

Dim svc As StudyValueConfiguration
Dim newSvc As StudyValueConfiguration
Dim newSvcs As StudyValueConfigurations

Set newSvcs = Clone.StudyValueConfigurations

For Each svc In mStudyValueConfigurations
    Set newSvc = newSvcs.Add(svc.ValueName)
    newSvc.BarStyle = svc.BarStyle
    newSvc.ChartRegionName = svc.ChartRegionName
    newSvc.DataPointStyle = svc.DataPointStyle
    newSvc.IncludeInChart = svc.IncludeInChart
    newSvc.Layer = svc.Layer
    newSvc.LineStyle = svc.LineStyle
    newSvc.SetBarFormatterFactory svc.BarFormatterFactory, svc.Bars
    newSvc.TextStyle = svc.TextStyle
Next

Clone.UnderlyingStudy = mUnderlyingStudy

Dim sc As StudyConfiguration
Dim newSc As StudyConfiguration
Dim newScs As StudyConfigurations

Set newScs = Clone.StudyConfigurations
For Each sc In mStudyConfigs
    Set newSc = sc.Clone
    newScs.Add newSc
Next

' don't do a deep copy of the studyValueHandlers because it's
' immutable as far as non-Friend callers are concerned
Clone.StudyValueHandlers = mStudyValueHandlers

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

' Only called when a chart has been cleared, so no need to worry about whether regions are
' still in use etc
Friend Sub Finish()
Dim svh As StudyValueHandler
Dim sc As StudyConfiguration

Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

If Not mStudyValueHandlers Is Nothing Then
    For Each svh In mStudyValueHandlers
        mStudy.RemoveStudyValueListener svh
    Next
    mStudyValueHandlers.Clear
End If

For Each sc In mStudyConfigs
    sc.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
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
mStudyValueConfigurations.LoadFromConfig mConfig.GetConfigurationSection(ConfigSectionStudyValueConfigurations), mBars

Set mStudyHorizontalRules = New StudyHorizontalRules
mStudyHorizontalRules.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudyHorizontalRules)

Set mStudyConfigs = New StudyConfigurations
mStudyConfigs.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudyConfigs)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
