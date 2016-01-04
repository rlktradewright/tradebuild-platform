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

Implements ITaskCompletionListener

'@================================================================================
' Events
'@================================================================================

Event BaseStudyConfigurationChanged(ByVal studyConfig As StudyConfiguration)
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
Private mStudyLibManager                                As StudyLibraryManager
Private mStudyManager                                   As StudyManager
Private mBarFormatterLibManager                         As BarFormatterLibManager

Private mIncludeBarsOutsideSession                      As Boolean

Private mRegionEntries                                  As EnumerableCollection

' we use this collection to locate studyConfigurations from Study objects
Private mAllStudyConfigs                                As Collection

Private mInstanceFullyQualifiedNames                    As Collection

Private mUpdatePerTick                                  As Boolean

Private mPeriods                                        As Periods

' the Study on which all others are built
Private mBaseStudy                                      As IStudy
Private mBaseStudyConfig                                As StudyConfiguration
Private mBars                                           As Bars

Private mConfig                                         As ConfigurationSection

Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mRegionEntries = New EnumerableCollection
Set mAllStudyConfigs = New Collection
Set mInstanceFullyQualifiedNames = New Collection
mUpdatePerTick = True
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' ITaskCompletionListener Interface Members
'@================================================================================

Private Sub ITaskCompletionListener_taskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_taskCompleted"
On Error GoTo Err

If ev.ErrorNumber <> 0 Then
    Debug.Print "Task: " & ev.TaskName; "; Error: " & ev.ErrorNumber; "; Message: "; ev.ErrorMessage
End If

mChartController.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then Set mBars = ev.Future.Value

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BarFormatterLibManager() As BarFormatterLibManager
Const ProcName As String = "BarFormatterLibManager"
On Error GoTo Err

Set BarFormatterLibManager = mBarFormatterLibManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BaseStudy() As IStudy
Const ProcName As String = "BaseStudy"
On Error GoTo Err

Set BaseStudy = mBaseStudy

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BaseStudyConfiguration() As StudyConfiguration
Const ProcName As String = "BaseStudyConfiguration"
On Error GoTo Err

If Not mBaseStudyConfig Is Nothing Then
    Set BaseStudyConfiguration = mBaseStudyConfig.Clone ' return a defensive copy
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Chart() As ChartController
Const ProcName As String = "Chart"
On Error GoTo Err

Set Chart = mChartController

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is mConfig Then Exit Property
Set mConfig = Value
If Not mBaseStudyConfig Is Nothing Then
    mBaseStudyConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBaseStudyConfig)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RegionNames() As String()
Const ProcName As String = "RegionNames"
On Error GoTo Err

If mRegionEntries.Count = 0 Then Exit Property

ReDim names(mRegionEntries.Count - 1) As String

Dim i As Long
For i = 1 To mRegionEntries.Count
    Dim re As RegionEntry
    re = mRegionEntries(i)
    names(i - 1) = re.Region.Name
Next
RegionNames = names

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyLibraryManager() As StudyLibraryManager
Set StudyLibraryManager = mStudyLibManager
End Property

Public Property Let UpdatePerTick(ByVal Value As Boolean)
mUpdatePerTick = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration, _
                Optional ByVal pNumberOfValuesToReplay As Long = ReplayNumbers.ReplayAll) As IStudy
Const ProcName As String = "AddStudyConfiguration"
On Error GoTo Err

Dim lStudy As IStudy
Set lStudy = AddStudy(studyConfig)

removeExistingStudyConfig lStudy

ApplyStudyConfiguration studyConfig, pNumberOfValuesToReplay

Set AddStudyConfiguration = lStudy
RaiseEvent StudyAdded(studyConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ApplyStudyConfiguration( _
                ByVal pStudyConfig As StudyConfiguration, _
                Optional ByVal pNumberOfValuesToReplay As Long = ReplayNumbers.ReplayAll)
Const ProcName As String = "ApplyStudyConfiguration"
On Error GoTo Err

setInstanceFullyQualifiedName pStudyConfig
addStudyConfig pStudyConfig
setupStudyValueListeners pStudyConfig, pNumberOfValuesToReplay
setupHorizontalRules pStudyConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearChart()
Const ProcName As String = "ClearChart"
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
On Error GoTo Err

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

Public Function GetDefaultStudyConfiguration( _
                ByVal Name As String, _
                ByVal studyLibName As String) As StudyConfiguration
Const ProcName As String = "GetDefaultStudyConfiguration"
On Error GoTo Err

Set GetDefaultStudyConfiguration = gGetDefaultStudyConfiguration(Name, mStudyLibManager, studyLibName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetPeriod(ByVal pTimestamp As Date) As Period
Const ProcName As String = "GetPeriod"
On Error GoTo Err

Static sPeriod As Period
Static sTimestamp As Date

If pTimestamp = sTimestamp Then
    If Not sPeriod Is Nothing Then
        Set GetPeriod = sPeriod
        Exit Function
    End If
End If

On Error Resume Next
Set GetPeriod = mPeriods(pTimestamp)
On Error GoTo Err

If GetPeriod Is Nothing Then Set GetPeriod = mPeriods.Add(pTimestamp)

Set sPeriod = GetPeriod
sTimestamp = pTimestamp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetSpecialValue(ByVal valueType As SpecialValues) As Variant
Const ProcName As String = "GetSpecialValue"
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetStudyConfiguration( _
                ByVal studyId As String) As StudyConfiguration
Const ProcName As String = "GetStudyConfiguration"
On Error GoTo Err

Set GetStudyConfiguration = mAllStudyConfigs(studyId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pChart As ChartController, _
                ByVal pStudyManager As StudyManager, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pIncludeBarsOutsideSession As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStudyManager = pStudyManager
Set mStudyLibManager = mStudyManager.StudyLibraryManager
Set mBarFormatterLibManager = pBarFormatterLibManager
mIncludeBarsOutsideSession = pIncludeBarsOutsideSession
Set mChartController = pChart
Set mPeriods = mChartController.Periods

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal BaseStudy As IStudy, _
                ByVal pIncludeBarsOutsideSession As Boolean)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
mIncludeBarsOutsideSession = pIncludeBarsOutsideSession

If Not mConfig Is Nothing Then
    Dim baseStudySect As ConfigurationSection
    Set baseStudySect = mConfig.GetConfigurationSection(ConfigSectionBaseStudyConfig)
    
    Dim studyConfig As StudyConfiguration
    Set studyConfig = New StudyConfiguration
    studyConfig.Study = BaseStudy
    studyConfig.UnderlyingStudy = mStudyManager.GetUnderlyingStudy(BaseStudy)
    
    studyConfig.LoadFromConfig baseStudySect
    
    setTheBaseStudyConfiguration studyConfig
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal timestamp As Date)
Const ProcName As String = "NotifyInput"
On Error GoTo Err

mStudyManager.NotifyInput inputHandle, inputValue, timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function RemoveStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "RemoveStudyConfiguration"
On Error GoTo Err

UnApplyStudyConfiguration studyConfig

Dim sc As StudyConfiguration
For Each sc In studyConfig.StudyConfigurations
    RemoveStudyConfiguration sc
Next

RaiseEvent StudyRemoved(studyConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ReplaceStudyConfiguration( _
                ByVal oldStudyConfig As StudyConfiguration, _
                ByVal newStudyConfig As StudyConfiguration) As IStudy
Const ProcName As String = "ReplaceStudyConfiguration"
On Error GoTo Err

If oldStudyConfig.Study Is BaseStudy Then
    newStudyConfig.Study = oldStudyConfig.Study
    SetBaseStudyConfiguration newStudyConfig
    Set ReplaceStudyConfiguration = newStudyConfig.Study
Else
    UnApplyStudyConfiguration oldStudyConfig
    
    Set ReplaceStudyConfiguration = AddStudy(newStudyConfig)
    ApplyStudyConfiguration newStudyConfig
    
    If newStudyConfig.Study Is oldStudyConfig.Study Then
        moveChildStudyConfigs oldStudyConfig, newStudyConfig
    Else
        reconfigureChildStudies oldStudyConfig, newStudyConfig
        StartStudy newStudyConfig.Study
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Sub ScrollToTime(ByVal pTime As Date)
Const ProcName As String = "ScrollToTime"
On Error GoTo Err

Dim periodNumber As Long
periodNumber = mPeriods(pTime).periodNumber
mChartController.LastVisiblePeriod = periodNumber + Int((mChartController.LastVisiblePeriod - mChartController.FirstVisiblePeriod) / 2) - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetBaseStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "SetBaseStudyConfiguration"
On Error GoTo Err

AssertArgument TypeOf studyConfig.Study Is IBarStudy, "Base Study must implement the BarStudy interface"

Dim sc As StudyConfiguration
Set sc = studyConfig.Clone ' take a defensive copy
sc.ClearStudyValueHandlers
sc.ClearStudyConfigs
setTheBaseStudyConfiguration sc

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetInputRegion( _
                ByVal inputHandle As Long, _
                ByVal ChartRegionName As String)
Const ProcName As String = "SetInputRegion"
On Error GoTo Err

Dim inputDescr As InputDescriptor
inputDescr = mStudyManager.GetInputDescriptor(inputHandle)

Dim lInputHandler As IStudy
Set lInputHandler = inputDescr.StudyInputHandler

On Error Resume Next
Dim studyConfig As StudyConfiguration
Set studyConfig = mAllStudyConfigs.Item(lInputHandler.ID)
On Error GoTo Err

If studyConfig Is Nothing Then
    Set studyConfig = New StudyConfiguration
    studyConfig.Study = lInputHandler
    studyConfig.Name = lInputHandler.InstanceName
    mAllStudyConfigs.Add studyConfig, lInputHandler.ID
End If

' need to do this here to ensure we have the definition with all inputs
'studyConfig.studyDefinition = inputDescr.lInputHandler.studyDefinition

Dim studyValueConfig As StudyValueConfiguration
Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(inputDescr.InputName)
studyValueConfig.ChartRegionName = ChartRegionName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartStudy( _
                ByVal pStudy As IStudy) As TaskController
Const ProcName As String = "StartStudy"
On Error GoTo Err

Dim lTaskCompletion As TaskController
Set lTaskCompletion = mStudyManager.StartStudy(pStudy, ReplayNumbers.ReplayAll)

If lTaskCompletion Is Nothing Then Exit Function

Set StartStudy = lTaskCompletion
lTaskCompletion.AddTaskCompletionListener Me
mChartController.DisableDrawing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub UnApplyStudyConfiguration( _
                ByVal pStudyConfig As StudyConfiguration)
Const ProcName As String = "UnApplyStudyConfiguration"
On Error GoTo Err

mChartController.DisableDrawing

removeStudyConfig pStudyConfig
removeStudyValueListeners pStudyConfig
removeHorizontalRules pStudyConfig

mChartController.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UpdateLastBar()
Const ProcName As String = "UpdateLastBar"
On Error GoTo Err

If Not mBaseStudyConfig Is Nothing Then mBaseStudyConfig.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub AddDependentStudies( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "AddDependentStudies"
On Error GoTo Err

Dim sc As StudyConfiguration
For Each sc In studyConfig.StudyConfigurations
    sc.UnderlyingStudy = studyConfig.Study
    AddStudyConfiguration sc
    StartStudy sc.Study
    AddDependentStudies sc
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function AddStudy( _
                ByVal studyConfig As StudyConfiguration) As IStudy
Const ProcName As String = "AddStudy"
On Error GoTo Err

Assert Not mBaseStudy Is Nothing, "Base Study has not yet been set"
AssertArgument Not studyConfig Is Nothing, "studyConfig is nothing"

If studyConfig.UnderlyingStudy Is Nothing Then
    studyConfig.UnderlyingStudy = mBaseStudy
End If

'Dim InputValueNames() As String
'InputValueNames = studyConfig.InputValueNames

Dim lStudy As IStudy
Set lStudy = mStudyManager.AddStudy(studyConfig.Name, _
                            studyConfig.UnderlyingStudy, _
                            studyConfig.InputValueNames, _
                            mIncludeBarsOutsideSession, _
                            studyConfig.Parameters, _
                            studyConfig.StudyLibraryName)
'' AddStudy will have replaced any $DEFAULT input Value names with
'' the correct default Study Value Name
'studyConfig.InputValueNames = InputValueNames

AssertArgument Not lStudy Is mBaseStudy, "Use the BaseStudyConfiguration property to set the base Study config"

studyConfig.Study = lStudy

Set AddStudy = lStudy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub addStudyConfig( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "addStudyConfig"
On Error GoTo Err

If Not studyConfig.Study Is mBaseStudy Then
    On Error Resume Next
    Dim parentStudyConfig As StudyConfiguration
    Set parentStudyConfig = mAllStudyConfigs(studyConfig.UnderlyingStudy.ID)
    On Error GoTo Err
    If Not parentStudyConfig Is Nothing Then parentStudyConfig.StudyConfigurations.Add studyConfig
End If

mAllStudyConfigs.Add studyConfig, studyConfig.Study.ID

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub determineDefaultRegionName( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "determineDefaultRegionName"
On Error GoTo Err

If studyConfig.ChartRegionName = RegionNameUnderlying Or _
    studyConfig.ChartRegionName = RegionNameDefault Or _
    studyConfig.ChartRegionName = "" _
Then
    studyConfig.ChartRegionName = getUnderlyingStudyRegionName(studyConfig)
ElseIf studyConfig.ChartRegionName = RegionNameCustom Then
    studyConfig.ChartRegionName = studyConfig.Study.InstancePath
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getRegion( _
                ByVal regionName As String, _
                ByVal title As String, _
                ByVal incrementUsageCount As Boolean) As ChartRegion
Const ProcName As String = "getRegion"
On Error GoTo Err

Dim re As RegionEntry

If mRegionEntries.Contains(regionName) Then
    re = mRegionEntries(regionName)
    re.usageCount = re.usageCount + 1
    mRegionEntries.Remove regionName
    mRegionEntries.Add re, regionName
    Set getRegion = re.Region
Else
    If mChartController.Regions.Contains(regionName) Then
        Set getRegion = mChartController.Regions.Item(regionName)
    Else
        Set getRegion = mChartController.Regions.Add(20, , , , regionName)
        getRegion.YGridlineSpacing = 0.8
        getRegion.title.Text = title
    End If
    Set re.Region = getRegion
    If incrementUsageCount Then re.usageCount = 1
    mRegionEntries.Add re, regionName
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getUnderlyingStudyRegionName(ByVal pStudyConfig As StudyConfiguration) As String
Const ProcName As String = "getUnderlyingStudyRegionName"
On Error GoTo Err

Dim underlyingStudyConfig As StudyConfiguration

Set underlyingStudyConfig = mAllStudyConfigs(pStudyConfig.UnderlyingStudy.ID)
getUnderlyingStudyRegionName = underlyingStudyConfig.ChartRegionName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub includeStudyValueInChart( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition, _
                ByVal numberOfValuesToReplay As Long)
Const ProcName As String = "includeStudyValueInChart"
On Error GoTo Err

Dim Study As IStudy
Set Study = studyConfig.Study

Dim regionName As String
If studyValueConfig.ChartRegionName = RegionNameDefault Or studyValueConfig.ChartRegionName = "" Then
    ' then use the Study's default Region
    regionName = studyConfig.ChartRegionName
ElseIf studyValueConfig.ChartRegionName = RegionNameCustom Then
    regionName = studyValueDef.Name & "." & Study.InstancePath
ElseIf studyValueConfig.ChartRegionName = RegionNameUnderlying Then
    regionName = getUnderlyingStudyRegionName(studyConfig)
Else
    regionName = studyValueConfig.ChartRegionName
End If

Dim Region As ChartRegion
Set Region = getRegion(regionName, Study.InstanceName, True)

Dim lStudyValueHandler As StudyValueHandler
Set lStudyValueHandler = studyConfig.StudyValueHandlers.Add(Me, _
                                        Region, _
                                        Study, _
                                        studyValueConfig, _
                                        studyValueDef, _
                                        mUpdatePerTick)


Dim lTaskCompletion As TaskController
Set lTaskCompletion = Study.AddStudyValueListener( _
                            lStudyValueHandler, _
                            studyValueConfig.ValueName, _
                            numberOfValuesToReplay, _
                            "AddStudyValueListener for Value " & studyValueConfig.ValueName & " to Study " & Study.InstanceName)

If Not lTaskCompletion Is Nothing Then
    lTaskCompletion.AddTaskCompletionListener Me
    mChartController.DisableDrawing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub moveChildStudyConfigs( _
                ByVal oldStudyConfig As StudyConfiguration, _
                ByVal newStudyConfig As StudyConfiguration)
Const ProcName As String = "moveChildStudyConfigs"
On Error GoTo Err

Dim childStudyConfig As StudyConfiguration
For Each childStudyConfig In oldStudyConfig.StudyConfigurations
    oldStudyConfig.StudyConfigurations.Remove childStudyConfig
    newStudyConfig.StudyConfigurations.Add childStudyConfig
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reconfigureChildStudies( _
                ByVal oldStudyConfig As StudyConfiguration, _
                ByVal newStudyConfig As StudyConfiguration)
Const ProcName As String = "reconfigureChildStudies"
On Error GoTo Err

Dim oldChildStudyConfig As StudyConfiguration
For Each oldChildStudyConfig In oldStudyConfig.StudyConfigurations
    Dim newChildStudyConfig As StudyConfiguration
    Set newChildStudyConfig = oldChildStudyConfig.Clone
    
    newChildStudyConfig.UnderlyingStudy = newStudyConfig.Study
    If oldChildStudyConfig.ChartRegionName = oldStudyConfig.ChartRegionName Then
        newChildStudyConfig.ChartRegionName = newStudyConfig.ChartRegionName
    End If
    
    RaiseEvent StudyAdded(newChildStudyConfig)
    
    UnApplyStudyConfiguration oldChildStudyConfig
    ApplyStudyConfiguration newChildStudyConfig
    
    reconfigureChildStudies oldChildStudyConfig, newChildStudyConfig
    
    StartStudy newChildStudyConfig.Study
    
    RaiseEvent StudyRemoved(oldChildStudyConfig)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub removeExistingStudyConfig( _
                ByVal pStudy As IStudy)
Const ProcName As String = "removeExistingStudyConfig"
On Error GoTo Err

On Error Resume Next
Dim prevStudyConfig As StudyConfiguration
Set prevStudyConfig = mAllStudyConfigs(pStudy.ID)
On Error GoTo Err

If Not prevStudyConfig Is Nothing Then
    RemoveStudyConfiguration prevStudyConfig
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub removeStudyConfig( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "removeStudyConfig"
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeHorizontalRules(ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "removeHorizontalRules"
On Error GoTo Err

Dim horizRulesLineSeries As LineSeries
Set horizRulesLineSeries = studyConfig.HorizontalRulesLineSeries

If horizRulesLineSeries Is Nothing Then Exit Sub
    
Dim re As RegionEntry
Set re.Region = Nothing
On Error Resume Next
re = mRegionEntries(studyConfig.ChartRegionName)
On Error GoTo Err
If Not re.Region Is Nothing Then
    Dim Region As ChartRegion
    Set Region = re.Region
    Region.RemoveGraphicObjectSeries horizRulesLineSeries
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeStudyValueListeners( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "removeStudyValueListeners"
On Error GoTo Err

Dim svh As StudyValueHandler
For Each svh In studyConfig.StudyValueHandlers
    studyConfig.Study.RemoveStudyValueListener svh
    
    Dim regionName As String
    regionName = svh.Region.Name
    
    Dim re As RegionEntry
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBaseStudyConfig( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "setBaseStudyConfig"
On Error GoTo Err

Set mBaseStudyConfig = studyConfig
Set mBaseStudy = studyConfig.Study

Dim bs As IBarStudy
Set bs = mBaseStudy
mFutureWaiter.Add bs.BarsFuture
If Not mConfig Is Nothing Then
    mBaseStudyConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBaseStudyConfig)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTheBaseStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "setTheBaseStudyConfiguration"
On Error GoTo Err

If mBaseStudy Is Nothing Then
    setBaseStudyConfig studyConfig
    
    ApplyStudyConfiguration studyConfig
    
    RaiseEvent StudyAdded(mBaseStudyConfig)
    AddDependentStudies studyConfig
ElseIf Not studyConfig.Study Is mBaseStudy Then
    UnApplyStudyConfiguration mBaseStudyConfig
    ApplyStudyConfiguration studyConfig
    
    reconfigureChildStudies mBaseStudyConfig, studyConfig
    
    setBaseStudyConfig studyConfig
Else
    UnApplyStudyConfiguration mBaseStudyConfig
    ApplyStudyConfiguration studyConfig
    
    moveChildStudyConfigs mBaseStudyConfig, studyConfig

    setBaseStudyConfig studyConfig
End If

RaiseEvent BaseStudyConfigurationChanged(studyConfig)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInstanceFullyQualifiedName( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "setInstanceFullyQualifiedName"
On Error GoTo Err

Dim key As String
key = studyConfig.Study.InstancePath

On Error Resume Next
Dim s As String
s = mInstanceFullyQualifiedNames(key)
On Error GoTo Err

If s = "" Then
    studyConfig.InstanceFullyQualifiedName = key
Else
    Dim keyInc As String
    Dim i As Long
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupHorizontalRules(ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "setupHorizontalRules"
On Error GoTo Err

If studyConfig.StudyHorizontalRules.Count = 0 Then Exit Sub
If Not mRegionEntries.Contains(studyConfig.ChartRegionName) Then Exit Sub

Dim Region As ChartRegion
Set Region = getRegion(studyConfig.ChartRegionName, studyConfig.Study.InstanceName, False)

Dim horizRulesLineSeries As LineSeries
Set horizRulesLineSeries = Region.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid + 1)
horizRulesLineSeries.Extended = True
horizRulesLineSeries.ExtendAfter = True
horizRulesLineSeries.ExtendBefore = True

Dim studyHorizRule As StudyHorizontalRule
For Each studyHorizRule In studyConfig.StudyHorizontalRules
    Dim line As ChartSkil27.line
    Set line = horizRulesLineSeries.Add
    line.Color = studyHorizRule.Color
    line.LineStyle = studyHorizRule.Style
    line.Thickness = studyHorizRule.Thickness
    line.Point1 = NewPoint(0, studyHorizRule.Y, CoordsRelative, CoordsLogical)
    line.Point2 = NewPoint(100, studyHorizRule.Y, CoordsRelative, CoordsLogical)
Next
studyConfig.HorizontalRulesLineSeries = horizRulesLineSeries

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupStudyValueListeners( _
                ByVal pStudyConfig As StudyConfiguration, _
                ByVal pNumberOfValuesToReplay As Long)
Const ProcName As String = "setupStudyValueListeners"
On Error GoTo Err

Dim studyValueDefs As StudyValueDefinitions
Set studyValueDefs = pStudyConfig.Study.StudyDefinition.StudyValueDefinitions

determineDefaultRegionName pStudyConfig

Dim studyValueConfig As StudyValueConfiguration
For Each studyValueConfig In pStudyConfig.StudyValueConfigurations
    If studyValueConfig.IncludeInChart Then
        includeStudyValueInChart pStudyConfig, _
                                studyValueConfig, _
                                studyValueDefs.Item(studyValueConfig.ValueName), _
                                pNumberOfValuesToReplay

    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



