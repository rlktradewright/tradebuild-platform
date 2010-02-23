VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartManager"
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

Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

Event StudyAdded(ByVal studyConfig As StudyConfiguration)
Event StudyRemoved(ByVal studyConfig As StudyConfiguration)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                                As String = "ChartManager"

Private Const ConfigSectionBaseStudyConfig              As String = "BaseStudyConfig"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mChartController                                As ChartController
Attribute mChartController.VB_VarHelpID = -1
Private mStudyManager                                   As studyManager

Private mRegionEntries                                  As Collection

' we use this collection to locate studyConfigurations from Study objects
Private mAllStudyConfigs                                As Collection

Private mInstanceFullyQualifiedNames                    As Collection

Private mUpdatePerTick                                  As Boolean

Private mPeriods                                        As Periods

' the Study on which all others are built
Private mBaseStudy                                      As Study
Private mBaseStudyConfig                                As StudyConfiguration
Private mBars                                           As Bars

Private mConfig                                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mRegionEntries = New Collection
Set mAllStudyConfigs = New Collection
Set mInstanceFullyQualifiedNames = New Collection
mUpdatePerTick = True

End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TaskCompletionEvent)

Const ProcName As String = "TaskCompletionListener_taskCompleted"
On Error GoTo Err

If ev.errorNumber <> 0 Then
    Debug.Print "Task: " & ev.taskName; "; Error: " & ev.errorNumber; "; Message: "; ev.errorMessage
End If

mChartController.EnableDrawing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BaseStudy() As Study
Const ProcName As String = "BaseStudy"
On Error GoTo Err

Set BaseStudy = mBaseStudy

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BaseStudyConfiguration() As StudyConfiguration
Const ProcName As String = "BaseStudyConfiguration"
On Error GoTo Err

If Not mBaseStudyConfig Is Nothing Then
    Set BaseStudyConfiguration = mBaseStudyConfig.Clone ' return a defensive copy
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let BaseStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "BaseStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

If Not TypeOf studyConfig.Study Is BarStudy Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Base Study must implement the BarStudy interface"
End If

Dim sc As StudyConfiguration
Set sc = studyConfig.Clone ' take a defensive copy
setBaseStudyConfiguration sc

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Chart() As ChartController
Const ProcName As String = "Chart"
On Error GoTo Err

Set Chart = mChartController

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
If Not mBaseStudyConfig Is Nothing Then
    mBaseStudyConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBaseStudyConfig)
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get RegionNames() As String()
Dim names() As String
Dim re As RegionEntry
Dim i As Long

Const ProcName As String = "RegionNames"
Dim failpoint As String
On Error GoTo Err

If mRegionEntries.count = 0 Then Exit Property

ReDim names(mRegionEntries.count - 1) As String

For i = 1 To mRegionEntries.count
    re = mRegionEntries(i)
    names(i - 1) = re.Region.Name
Next
RegionNames = names

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let UpdatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration) As Study
Dim lStudy As Study

Const ProcName As String = "AddStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

Set lStudy = AddStudy(studyConfig)

removeExistingStudyConfig lStudy

AddStudyConfig studyConfig

Set AddStudyConfiguration = lStudy
RaiseEvent StudyAdded(studyConfig)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub ClearChart()
Const ProcName As String = "ClearChart"
Dim failpoint As String
On Error GoTo Err

mChartController.ClearChart
Set mPeriods = mChartController.Periods

If Not mBaseStudyConfig Is Nothing Then
    mBaseStudyConfig.Finish
    mBaseStudyConfig.RemoveFromConfig
    Set mBaseStudyConfig = Nothing
    Set mBaseStudy = Nothing
End If

Set mAllStudyConfigs = New Collection
Set mRegionEntries = New Collection
Set mBars = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub Finish()

On Error GoTo Err

mChartController.Finish
Set mChartController = Nothing
Set mStudyManager = Nothing
Set mAllStudyConfigs = Nothing

Set mPeriods = Nothing
If Not mBaseStudyConfig Is Nothing Then
    mBaseStudyConfig.Finish
    Set mBaseStudyConfig = Nothing
    Set mBaseStudy = Nothing
End If
Set mRegionEntries = Nothing
Set mBars = Nothing

Exit Sub

Err:
'ignore any errors
End Sub

Friend Function GetPeriod(ByVal pTimestamp As Date) As Period
Static sPeriod As Period
Static sTimestamp As Date


Const ProcName As String = "GetPeriod"
Dim failpoint As String
On Error GoTo Err

If pTimestamp = sTimestamp Then
    If Not sPeriod Is Nothing Then
        Set GetPeriod = sPeriod
        Exit Function
    End If
End If

On Error Resume Next
Set GetPeriod = mPeriods(pTimestamp)
On Error GoTo Err

If GetPeriod Is Nothing Then
    Set GetPeriod = mPeriods.Add(pTimestamp)
    mChartController.ScrollX 1
End If
Set sPeriod = GetPeriod
sTimestamp = pTimestamp

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function GetSpecialValue(ByVal valueType As SpecialValues) As Variant

Const ProcName As String = "GetSpecialValue"
Dim failpoint As String
On Error GoTo Err

Select Case valueType
Case SpecialValues.SVCurrentSessionEndTime
    GetSpecialValue = mChartController.CurrentSessionEndTime
Case SpecialValues.SVCurrentSessionStartTime
    GetSpecialValue = mChartController.CurrentSessionStartTime
Case SpecialValues.SVHighPrice
    GetSpecialValue = mBars.HighValue
Case SpecialValues.SVLowPrice
    GetSpecialValue = mBars.LowValue
Case SpecialValues.SVPreviousClosePrice
    GetSpecialValue = mBars.CloseValue(-1)
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function GetStudyConfig( _
                ByVal studyId As String) As StudyConfiguration
Const ProcName As String = "GetStudyConfig"
Dim failpoint As String
On Error GoTo Err

Set GetStudyConfig = mAllStudyConfigs(studyId)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal pStudyManager As studyManager, _
                ByVal pChart As ChartController)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mStudyManager = pStudyManager
Set mChartController = pChart
Set mPeriods = mChartController.Periods

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal BaseStudy As Study)
Dim baseStudySect As ConfigurationSection
Dim studyConfig As StudyConfiguration

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = config

If Not mConfig Is Nothing Then
    Set baseStudySect = mConfig.GetConfigurationSection(ConfigSectionBaseStudyConfig)
    Set studyConfig = New StudyConfiguration
    studyConfig.Study = BaseStudy
    studyConfig.UnderlyingStudy = mStudyManager.GetUnderlyingStudy(BaseStudy)
    
    studyConfig.LoadFromConfig baseStudySect
    
    setBaseStudyConfiguration studyConfig
    
    'AddDependentStudies studyConfig
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub NotifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal timestamp As Date)
Const ProcName As String = "NotifyInput"
Dim failpoint As String
On Error GoTo Err

mStudyManager.NotifyInput inputHandle, inputValue, timestamp

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function RemoveStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)
Dim sc As StudyConfiguration

Const ProcName As String = "RemoveStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

removeStudyValueListeners studyConfig

For Each sc In studyConfig.StudyConfigurations
    RemoveStudyConfiguration sc
Next

removeStudyConfig studyConfig

RaiseEvent StudyRemoved(studyConfig)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function ReplaceStudyConfiguration( _
                ByVal oldStudyConfig As StudyConfiguration, _
                ByVal newStudyConfig As StudyConfiguration) As Study

Const ProcName As String = "ReplaceStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

Set ReplaceStudyConfiguration = AddStudy(newStudyConfig)
removeStudyValueListeners oldStudyConfig

If newStudyConfig.Study Is oldStudyConfig.Study Then
    
    moveChildStudyConfigs oldStudyConfig, newStudyConfig

    removeStudyConfig oldStudyConfig
    
    AddStudyConfig newStudyConfig
Else

    AddStudyConfig newStudyConfig
    
    RaiseEvent StudyRemoved(oldStudyConfig)
    RaiseEvent StudyAdded(newStudyConfig)
    
    reconfigureChildStudies oldStudyConfig, newStudyConfig
    
    removeStudyConfig oldStudyConfig
    
    StartStudy newStudyConfig.Study

End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Sub ScrollToTime(ByVal pTime As Date)
Dim periodNumber As Long

Const ProcName As String = "ScrollToTime"
Dim failpoint As String
On Error GoTo Err

periodNumber = mPeriods(pTime).periodNumber
mChartController.LastVisiblePeriod = periodNumber + Int((mChartController.LastVisiblePeriod - mChartController.FirstVisiblePeriod) / 2) - 1

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub SetInputRegion( _
                ByVal inputHandle As Long, _
                ByVal ChartRegionName As String)
Dim inputDescr As InputDescriptor
Dim studyConfig As StudyConfiguration
Dim studyValueConfig As StudyValueConfiguration
Dim source As Study

Const ProcName As String = "SetInputRegion"
Dim failpoint As String
On Error GoTo Err

inputDescr = mStudyManager.GetInputDescriptor(inputHandle)
Set source = inputDescr.source

On Error Resume Next
Set studyConfig = mAllStudyConfigs.Item(source.ID)
On Error GoTo Err

If studyConfig Is Nothing Then
    Set studyConfig = New StudyConfiguration
    studyConfig.Study = source
    studyConfig.Name = source.InstanceName
    mAllStudyConfigs.Add studyConfig, source.ID
End If

' need to do this here to ensure we have the definition with all inputs
'studyConfig.studyDefinition = inputDescr.source.studyDefinition

Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(inputDescr.inputName)
studyValueConfig.ChartRegionName = ChartRegionName

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function StartStudy( _
                ByVal pStudy As Study) As TaskController
Dim lTaskCompletion As TaskController

Const ProcName As String = "StartStudy"
Dim failpoint As String
On Error GoTo Err

Set lTaskCompletion = mStudyManager.StartStudy(pStudy, ReplayNumbers.ReplayAll)

If lTaskCompletion Is Nothing Then Exit Function

Set StartStudy = lTaskCompletion
lTaskCompletion.AddTaskCompletionListener Me
mChartController.DisableDrawing

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub AddDependentStudies( _
                ByVal studyConfig As StudyConfiguration)
Dim sc As StudyConfiguration
Const ProcName As String = "AddDependentStudies"
Dim failpoint As String
On Error GoTo Err

For Each sc In studyConfig.StudyConfigurations
    sc.UnderlyingStudy = studyConfig.Study
    AddStudyConfiguration sc
    StartStudy sc.Study
    AddDependentStudies sc
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function AddStudy( _
                ByVal studyConfig As StudyConfiguration) As Study
Dim lStudy As Study
Dim Region As ChartRegion
Dim studyHorizRule As StudyHorizontalRule
Dim horizRulesLineSeries As LineSeries
Dim line As ChartSkil26.line
Dim InputValueNames() As String

Const ProcName As String = "AddStudy"
Dim failpoint As String
On Error GoTo Err

If mBaseStudy Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Base Study Configuration has not yet been set"
End If

If studyConfig.UnderlyingStudy Is Nothing Then
    studyConfig.UnderlyingStudy = mBaseStudy
End If

InputValueNames = studyConfig.InputValueNames
Set lStudy = mStudyManager.AddStudy(studyConfig.Name, _
                            studyConfig.UnderlyingStudy, _
                            InputValueNames, _
                            studyConfig.Parameters, _
                            studyConfig.StudyLibraryName)
' AddStudy will have replaced any $DEFAULT input value names with
' the correct default Study value Name
studyConfig.InputValueNames = InputValueNames

If lStudy Is mBaseStudy Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Use the BaseStudyConfiguration property to set the base Study config"
End If

studyConfig.Study = lStudy

determineDefaultRegionName studyConfig

Set Region = findRegion(studyConfig.ChartRegionName, lStudy.InstanceName, False)
    
setupStudyValueListeners studyConfig

If studyConfig.StudyHorizontalRules.count > 0 Then
    Set horizRulesLineSeries = Region.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid + 1)
    horizRulesLineSeries.Extended = True
    horizRulesLineSeries.ExtendAfter = True
    horizRulesLineSeries.ExtendBefore = True
    For Each studyHorizRule In studyConfig.StudyHorizontalRules
        Set line = horizRulesLineSeries.Add
        line.Color = studyHorizRule.Color
        line.LineStyle = studyHorizRule.Style
        line.Thickness = studyHorizRule.Thickness
        line.Point1 = Region.NewPoint(0, studyHorizRule.Y, CoordsRelative, CoordsLogical)
        line.Point2 = Region.NewPoint(100, studyHorizRule.Y, CoordsRelative, CoordsLogical)
    Next
    studyConfig.HorizontalRulesLineSeries = horizRulesLineSeries
End If

Set AddStudy = lStudy

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub AddStudyConfig( _
                ByVal studyConfig As StudyConfiguration)
Dim parentStudyConfig As StudyConfiguration

Const ProcName As String = "AddStudyConfig"
Dim failpoint As String
On Error GoTo Err

setInstanceFullyQualifiedName studyConfig

If Not studyConfig.Study Is mBaseStudy Then
    On Error Resume Next
    Set parentStudyConfig = mAllStudyConfigs(studyConfig.UnderlyingStudy.ID)
    On Error GoTo Err
    If Not parentStudyConfig Is Nothing Then parentStudyConfig.StudyConfigurations.Add studyConfig
End If

mAllStudyConfigs.Add studyConfig, studyConfig.Study.ID

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub determineDefaultRegionName( _
                ByVal studyConfig As StudyConfiguration)
Dim underlyingStudyConfig As StudyConfiguration

Const ProcName As String = "determineDefaultRegionName"
Dim failpoint As String
On Error GoTo Err

If studyConfig.ChartRegionName = RegionNameDefault Or _
    studyConfig.ChartRegionName = "" _
Then
    ' we will use the Region for the underlying Study
    Set underlyingStudyConfig = mAllStudyConfigs(studyConfig.UnderlyingStudy.ID)
    studyConfig.ChartRegionName = underlyingStudyConfig.ChartRegionName
ElseIf studyConfig.ChartRegionName = RegionNameCustom Then
    studyConfig.ChartRegionName = studyConfig.Study.InstancePath
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function findRegion( _
                ByVal regionName As String, _
                ByVal title As String, _
                ByVal incrementUsageCount As Boolean) As ChartRegion
Dim re As RegionEntry

Const ProcName As String = "findRegion"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
re = mRegionEntries(regionName)
On Error GoTo Err

Set findRegion = re.Region

If findRegion Is Nothing Then
    If mChartController.Regions.Contains(regionName) Then
        Set findRegion = mChartController.Regions.Item(regionName)
    Else
        Set findRegion = mChartController.Regions.Add(20, , , , regionName)
        findRegion.GridlineSpacingY = 0.8
        findRegion.title.Text = title
        findRegion.title.Color = vbBlue
    End If
    Set re.Region = findRegion
    If incrementUsageCount Then re.usageCount = 1
    mRegionEntries.Add re, regionName
ElseIf incrementUsageCount Then
    re.usageCount = re.usageCount + 1
    mRegionEntries.Remove regionName
    mRegionEntries.Add re, regionName
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function includeStudyValueInChart( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition) As TaskController

Dim lStudyValueHandler As StudyValueHandler
Dim Study As Study
Dim Region As ChartRegion
Dim regionName As String
Dim lTaskCompletion As TaskController

Const ProcName As String = "includeStudyValueInChart"
Dim failpoint As String
On Error GoTo Err

Set Study = studyConfig.Study
If studyValueConfig.ChartRegionName = RegionNameDefault Or studyValueConfig.ChartRegionName = "" Then
    ' then use the Study's default Region
    regionName = studyConfig.ChartRegionName
ElseIf studyValueConfig.ChartRegionName = RegionNameCustom Then
    regionName = studyValueDef.Name & "." & Study.InstancePath
    'studyValueConfig.ChartRegionName = regionName
Else
    regionName = studyValueConfig.ChartRegionName
End If

Set Region = findRegion(regionName, Study.InstanceName, True)

Set lStudyValueHandler = studyConfig.StudyValueHandlers.Add(Me, _
                                        Region, _
                                        studyValueConfig, _
                                        studyValueDef, _
                                        mUpdatePerTick)


Set lTaskCompletion = Study.AddStudyValueListener( _
                            lStudyValueHandler, _
                            studyValueConfig.ValueName, _
                            ReplayNumbers.ReplayAll, _
                            "AddStudyValueListener for value " & studyValueConfig.ValueName & " to Study " & Study.InstanceName)

If Not lTaskCompletion Is Nothing Then
    lTaskCompletion.AddTaskCompletionListener Me
    mChartController.DisableDrawing
End If
Set includeStudyValueInChart = lTaskCompletion

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub moveChildStudyConfigs( _
                ByVal oldStudyConfig As StudyConfiguration, _
                ByVal newStudyConfig As StudyConfiguration)
Dim childStudyConfig As StudyConfiguration
Const ProcName As String = "moveChildStudyConfigs"
Dim failpoint As String
On Error GoTo Err

For Each childStudyConfig In oldStudyConfig.StudyConfigurations
    oldStudyConfig.StudyConfigurations.Remove childStudyConfig
    newStudyConfig.StudyConfigurations.Add childStudyConfig
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub reconfigureChildStudies( _
                ByVal oldStudyConfig As StudyConfiguration, _
                ByVal newStudyConfig As StudyConfiguration)
Dim oldChildStudyConfig As StudyConfiguration
Dim newChildStudyConfig As StudyConfiguration

Const ProcName As String = "reconfigureChildStudies"
Dim failpoint As String
On Error GoTo Err

For Each oldChildStudyConfig In oldStudyConfig.StudyConfigurations
    Set newChildStudyConfig = oldChildStudyConfig.Clone
    newChildStudyConfig.UnderlyingStudy = newStudyConfig.Study
    If oldChildStudyConfig.ChartRegionName = oldStudyConfig.ChartRegionName Then
        newChildStudyConfig.ChartRegionName = newStudyConfig.ChartRegionName
    End If
    
    
    removeStudyValueListeners oldChildStudyConfig
    
    AddStudy newChildStudyConfig
    
    AddStudyConfig newChildStudyConfig
    
    reconfigureChildStudies oldChildStudyConfig, newChildStudyConfig
    
    removeStudyConfig oldChildStudyConfig
    
    StartStudy newChildStudyConfig.Study
    
    RaiseEvent StudyRemoved(oldChildStudyConfig)
    RaiseEvent StudyAdded(newChildStudyConfig)
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub removeExistingStudyConfig( _
                ByVal pStudy As Study)
Dim prevStudyConfig As StudyConfiguration

Const ProcName As String = "removeExistingStudyConfig"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
Set prevStudyConfig = mAllStudyConfigs(pStudy.ID)
On Error GoTo Err

If Not prevStudyConfig Is Nothing Then
    RemoveStudyConfiguration prevStudyConfig
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub removeStudyConfig( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "removeStudyConfig"
Dim failpoint As String
On Error GoTo Err

If studyConfig.Study Is mBaseStudy Then
    mBaseStudyConfig.RemoveFromConfig
Else
    Dim parentStudyConfig As StudyConfiguration
    Set parentStudyConfig = mAllStudyConfigs(studyConfig.UnderlyingStudy.ID)
    
    parentStudyConfig.StudyConfigurations.Remove studyConfig
End If

mAllStudyConfigs.Remove studyConfig.Study.ID
mInstanceFullyQualifiedNames.Remove studyConfig.InstanceFullyQualifiedName

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub removeStudyValueListeners( _
                ByVal studyConfig As StudyConfiguration)
Dim svh As StudyValueHandler
Dim horizRulesLineSeries As LineSeries
Dim regionName As String
Dim re As RegionEntry
Dim Region As ChartRegion

Const ProcName As String = "removeStudyValueListeners"
Dim failpoint As String
On Error GoTo Err

mChartController.DisableDrawing
For Each svh In studyConfig.StudyValueHandlers
    studyConfig.Study.RemoveStudyValueListener svh
    regionName = svh.Region.Name
    re = mRegionEntries(regionName)
    If re.usageCount = 1 Then
        mChartController.Regions.Remove re.Region
        mRegionEntries.Remove regionName
    Else
        re.Region.RemoveGraphicObjectSeries svh.ValueSeries
        mRegionEntries.Remove regionName
        re.usageCount = re.usageCount - 1
        mRegionEntries.Add re, regionName
    End If
Next

studyConfig.StudyValueHandlers.Clear

Set horizRulesLineSeries = studyConfig.HorizontalRulesLineSeries
If Not horizRulesLineSeries Is Nothing Then
    Set re.Region = Nothing
    On Error Resume Next
    re = mRegionEntries(studyConfig.ChartRegionName)
    On Error GoTo Err
    If Not re.Region Is Nothing Then
        Set Region = re.Region
        Region.RemoveGraphicObjectSeries horizRulesLineSeries
    End If
End If

mChartController.EnableDrawing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setBaseStudyConfig( _
                ByVal studyConfig As StudyConfiguration)
Dim bs As BarStudy
Const ProcName As String = "setBaseStudyConfig"
Dim failpoint As String
On Error GoTo Err

Set mBaseStudyConfig = studyConfig
Set mBaseStudy = studyConfig.Study
Set bs = mBaseStudy
Set mBars = bs.Bars
If Not mConfig Is Nothing Then
    mBaseStudyConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBaseStudyConfig)
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setBaseStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)

Const ProcName As String = "setBaseStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

If mBaseStudy Is Nothing Then
    setBaseStudyConfig studyConfig
    
    AddStudyConfig studyConfig
    
    setupStudyValueListeners mBaseStudyConfig
    RaiseEvent StudyAdded(mBaseStudyConfig)
    AddDependentStudies studyConfig

ElseIf Not studyConfig.Study Is mBaseStudy Then
    AddStudyConfig studyConfig
    
    setupStudyValueListeners studyConfig
    removeStudyValueListeners mBaseStudyConfig
    
    RaiseEvent StudyRemoved(mBaseStudyConfig)
    RaiseEvent StudyAdded(studyConfig)
    
    reconfigureChildStudies mBaseStudyConfig, studyConfig
    
    removeStudyConfig mBaseStudyConfig

    setBaseStudyConfig studyConfig
Else
    
    setupStudyValueListeners studyConfig
    removeStudyValueListeners mBaseStudyConfig
    
    moveChildStudyConfigs mBaseStudyConfig, studyConfig

    removeStudyConfig mBaseStudyConfig
    
    setBaseStudyConfig studyConfig
    AddStudyConfig mBaseStudyConfig
    
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setInstanceFullyQualifiedName( _
                ByVal studyConfig As StudyConfiguration)
Dim key As String
Dim keyInc As String
Dim i As Long
Dim s As String

Const ProcName As String = "setInstanceFullyQualifiedName"
Dim failpoint As String
On Error GoTo Err

key = studyConfig.Study.InstancePath

On Error Resume Next
s = mInstanceFullyQualifiedNames(key)
On Error GoTo Err

If s = "" Then
    studyConfig.InstanceFullyQualifiedName = key
Else
    Do
        i = i + 1
        keyInc = key & "." & i
        
        On Error Resume Next
        s = ""
        s = mInstanceFullyQualifiedNames(keyInc)
        On Error GoTo Err
    
    Loop Until s = ""
    studyConfig.InstanceFullyQualifiedName = keyInc
End If

mInstanceFullyQualifiedNames.Add studyConfig.InstanceFullyQualifiedName, studyConfig.InstanceFullyQualifiedName

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function setupStudyValueListeners( _
                ByVal studyConfig As StudyConfiguration) As TaskController()
Dim studyValueConfig As StudyValueConfiguration
Dim studyValueDefs As StudyValueDefinitions
Dim taskCompletions() As TaskController
Dim taskCompletionsIndex As Long
Dim tc As TaskController

Const ProcName As String = "setupStudyValueListeners"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

Set studyValueDefs = studyConfig.Study.StudyDefinition.StudyValueDefinitions

failpoint = 300

For Each studyValueConfig In studyConfig.StudyValueConfigurations
    If studyValueConfig.IncludeInChart Then

        failpoint = 400

        Set tc = includeStudyValueInChart(studyConfig, _
                                studyValueConfig, _
                                studyValueDefs.Item(studyValueConfig.ValueName))

        failpoint = 500

        If Not tc Is Nothing Then
            ReDim Preserve taskCompletions(taskCompletionsIndex) As TaskController
            Set taskCompletions(taskCompletionsIndex) = tc
            taskCompletionsIndex = taskCompletionsIndex + 1
        End If
    End If
Next

failpoint = 600

setupStudyValueListeners = taskCompletions

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function


