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

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements Tasks.TaskCompletionListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'Private Const TaskTypeStartStudy As Long = 1
'Private Const TaskTypeReplayBars As Long = 2
'Private Const TaskTypeAddValueListener As Long = 3

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Public Type RegionEntry
    region              As ChartRegion
    usageCount          As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mChartController As chartController
Attribute mChartController.VB_VarHelpID = -1
Private mStudyManager As studyManager

Private mSourceStudies As Collection

Private mRegions As Collection

Private mStudyConfigurations As StudyConfigurations

Private mUpdatePerTick As Boolean

Private mPeriods As Periods

Private mHighPrice As Double
Private mLowPrice As Double
Private mPrevClosePrice As Double

' the study on which all others are built
Private mBaseStudy As study

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mSourceStudies = New Collection
Set mRegions = New Collection
Set mStudyConfigurations = New StudyConfigurations
mUpdatePerTick = True

End Sub

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As Tasks.TaskCompletionEvent)

If ev.resultCode <> 0 Then
    Debug.Print "Task: " & ev.TaskName; "; Error: " & ev.resultCode; "; Message: " & ev.message
End If

mChartController.suppressDrawing = False

End Sub

'================================================================================
' mChartController Event Handlers
'================================================================================

Private Sub mChartController_ChartCleared()
Dim studyConfig As StudyConfiguration
Set mPeriods = mChartController.Periods
For Each studyConfig In mStudyConfigurations
    studyConfig.finish
Next
Set mSourceStudies = New Collection
Set mRegions = New Collection
Set mStudyConfigurations = New StudyConfigurations
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get chartController() As chartController
Set chartController = mChartController
End Property

Public Property Get regionNames() As String()
Dim names() As String
Dim re As RegionEntry
Dim i As Long

ReDim names(mRegions.count - 1) As String

For i = 1 To mRegions.count
    re = mRegions(i)
    names(i - 1) = re.region.name
Next
regionNames = names
End Property

Public Property Get StudyConfigurations() As StudyConfigurations
Set StudyConfigurations = mStudyConfigurations
End Property

Public Property Let suppressDrawing( _
                ByVal value As Boolean)
mChartController.suppressDrawing = value
End Property

Public Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addChartRegion( _
                    ByVal name As String, _
                    ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double) As ChartRegion
Dim re As RegionEntry
Set addChartRegion = mChartController.addChartRegion(percentheight, minimumPercentHeight, name)
Set re.region = addChartRegion
mRegions.add re, name
End Function

Public Function addStudy( _
                ByVal studyConfig As StudyConfiguration, _
                Optional noStudyValueListeners As Boolean) As study
Dim study As study
Dim region As ChartRegion
Dim studyHorizRule As StudyHorizontalRule
Dim horizRulesLineSeries As LineSeries
Dim line As ChartSkil25.line

If studyConfig.underlyingStudy Is Nothing Then
    studyConfig.underlyingStudy = mBaseStudy
    'studyConfig.inputValueNames = "Close"
End If

Set study = mStudyManager.addStudy(studyConfig.name, _
                            studyConfig.underlyingStudy, _
                            studyConfig.inputValueNames, _
                            studyConfig.parameters, _
                            studyConfig.StudyLibraryName)

If mBaseStudy Is Nothing Then Set mBaseStudy = study

'studyConfig.instanceName = study.instanceName
'studyConfig.instanceFullyQualifiedName = study.instancePath
studyConfig.study = study

determineDefaultRegionName studyConfig

Set region = findRegion(studyConfig.chartRegionName, study.instanceName, False)
    
If Not noStudyValueListeners Then
    setupStudyValueListeners studyConfig
End If

If studyConfig.studyHorizontalRules.count > 0 Then
    Set horizRulesLineSeries = region.addLineSeries(LayerNumbers.LayerGrid + 1)
    horizRulesLineSeries.extended = True
    horizRulesLineSeries.extendAfter = True
    horizRulesLineSeries.extendBefore = True
    For Each studyHorizRule In studyConfig.studyHorizontalRules
        horizRulesLineSeries.color = studyHorizRule.color
        horizRulesLineSeries.style = studyHorizRule.style
        horizRulesLineSeries.thickness = studyHorizRule.thickness
        Set line = horizRulesLineSeries.add
        line.point1 = region.newPoint(0, studyHorizRule.y, CoordsRelative, CoordsLogical)
        line.point2 = region.newPoint(100, studyHorizRule.y, CoordsRelative, CoordsLogical)
    Next
    studyConfig.horizontalRulesLineSeries = horizRulesLineSeries
End If

Set addStudy = study
End Function

Public Sub clearChart()
Set mRegions = New Collection
initialiseChart
End Sub

Public Sub finish()
Dim studyConfig As StudyConfiguration

On Error GoTo err

mChartController.clearChart

For Each studyConfig In mStudyConfigurations
    studyConfig.finish
Next

Set mStudyConfigurations = Nothing

Exit Sub

err:
'ignore any errors
End Sub

Friend Function getPeriod(ByVal pTimestamp As Date) As Period
Static sPeriod As Period
Static sTimestamp As Date


If pTimestamp = sTimestamp Then
    If Not sPeriod Is Nothing Then
        Set getPeriod = sPeriod
        Exit Function
    End If
End If

On Error Resume Next
Set getPeriod = mPeriods(pTimestamp)
On Error GoTo 0

If getPeriod Is Nothing Then
    Set getPeriod = mPeriods.add(pTimestamp)
    mChartController.scrollX 1
End If
Set sPeriod = getPeriod
sTimestamp = pTimestamp
End Function

Friend Function getSpecialValue(ByVal valueType As SpecialValues) As Variant

Select Case valueType
Case SpecialValues.SVCurrentSessionEndTime
    getSpecialValue = mChartController.currentSessionEndTime
Case SpecialValues.SVCurrentSessionStartTime
    getSpecialValue = mChartController.currentSessionStartTime
Case SpecialValues.SVHighPrice
    getSpecialValue = mHighPrice
Case SpecialValues.SVLowPrice
    getSpecialValue = mLowPrice
Case SpecialValues.SVPreviousClosePrice
    getSpecialValue = mPrevClosePrice
End Select
End Function

Friend Sub initialise( _
                ByVal pStudyManager As studyManager, _
                ByVal pChartController As chartController)
Set mStudyManager = pStudyManager
Set mChartController = pChartController
initialiseChart
End Sub

Public Sub notifyInput( _
                ByVal inputHandle As Long, _
                ByVal inputValue As Variant, _
                ByVal timestamp As Date)
mStudyManager.notifyInput inputHandle, inputValue, timestamp
End Sub

Public Function removeStudy( _
                ByVal studyConfig As StudyConfiguration)
Dim svh As StudyValueHandler
Dim horizRulesLineSeries As LineSeries
Dim regionName As String
Dim re As RegionEntry
Dim region As ChartRegion

mChartController.suppressDrawing = True
For Each svh In studyConfig.studyValueHandlers
    studyConfig.study.removeStudyValueListener svh
    regionName = svh.region.name
    re = mRegions(regionName)
    If re.usageCount = 1 Then
        svh.finish False
        mChartController.removeChartRegion re.region
        mRegions.remove regionName
    Else
        svh.finish True
        mRegions.remove regionName
        re.usageCount = re.usageCount - 1
        mRegions.add re, regionName
    End If
Next

Set horizRulesLineSeries = studyConfig.horizontalRulesLineSeries
If Not horizRulesLineSeries Is Nothing Then
    Set re.region = Nothing
    On Error Resume Next
    re = mRegions(studyConfig.chartRegionName)
    On Error GoTo 0
    If Not re.region Is Nothing Then
        Set region = re.region
        region.removeLineSeries horizRulesLineSeries
    End If
End If
mChartController.suppressDrawing = False
mStudyConfigurations.remove studyConfig
End Function

Public Sub scrollToTime(ByVal pTime As Date)
Dim periodNumber As Long
periodNumber = mPeriods(pTime).periodNumber
mChartController.lastVisiblePeriod = periodNumber + Int((mChartController.lastVisiblePeriod - mChartController.firstVisiblePeriod) / 2) - 1
End Sub

Public Sub setInputRegion( _
                ByVal inputHandle As Long, _
                ByVal chartRegionName As String)
Dim inputDescr As InputDescriptor
Dim studyConfig As StudyConfiguration
Dim studyValueConfig As StudyValueConfiguration
Dim source As InputStudy

On Error GoTo err

inputDescr = mStudyManager.getInputDescriptor(inputHandle)
Set source = inputDescr.source

On Error Resume Next
Set studyConfig = mStudyConfigurations.item(source.name)
On Error GoTo 0

If studyConfig Is Nothing Then
    Set studyConfig = New StudyConfiguration
'    studyConfig.instanceName = source.name
'    studyConfig.instanceFullyQualifiedName = source.name
    studyConfig.study = source
    studyConfig.name = source.name
    mStudyConfigurations.add studyConfig
End If

' need to do this here to ensure we have the definition with all inputs
'studyConfig.studyDefinition = inputDescr.source.studyDefinition

Set studyValueConfig = studyConfig.studyValueConfigurations.add(inputDescr.inputName)
studyValueConfig.chartRegionName = chartRegionName
Exit Sub

err:
If err.Number = VBErrorCodes.VbErrSubscriptOutOfRange Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSupport" & "." & "ChartManager" & ":" & "setInputRegion", _
            "inputHandle argument invalid"
End If
err.Raise err.Number
End Sub

Public Function setupStudyValueListeners( _
                ByVal studyConfig As StudyConfiguration) As TaskCompletion()
Dim studyValueConfig As StudyValueConfiguration
Dim studyValueDefs As StudyValueDefinitions
Dim taskCompletions() As TaskCompletion
Dim taskCompletionsIndex As Long
Dim tc As TaskCompletion

Set studyValueDefs = studyConfig.study.StudyDefinition.StudyValueDefinitions

For Each studyValueConfig In studyConfig.studyValueConfigurations
    If studyValueConfig.includeInChart Then
        Set tc = includeStudyValueInChart(studyConfig, _
                                studyValueConfig, _
                                studyValueDefs.item(studyValueConfig.valueName))
        If Not tc Is Nothing Then
            ReDim Preserve taskCompletions(taskCompletionsIndex) As TaskCompletion
            Set taskCompletions(taskCompletionsIndex) = tc
            taskCompletionsIndex = taskCompletionsIndex + 1
        End If
    End If
Next

mStudyConfigurations.add studyConfig

If taskCompletionsIndex > 0 Then setupStudyValueListeners = taskCompletions
End Function

Public Function startStudy( _
                ByVal pStudy As study) As TaskCompletion
Dim lTaskCompletion As Tasks.TaskCompletion
Set lTaskCompletion = mStudyManager.startStudy(pStudy, ReplayNumbers.ReplayAll)

If lTaskCompletion Is Nothing Then Exit Function

Set startStudy = lTaskCompletion
lTaskCompletion.addTaskCompletionListener Me
mChartController.suppressDrawing = True
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub determineDefaultRegionName( _
                ByVal studyConfig As StudyConfiguration)
Dim underlyingStudyConfig As StudyConfiguration

If studyConfig.chartRegionName = RegionNameDefault Then
    ' we will use the region for the underlying study
    For Each underlyingStudyConfig In mStudyConfigurations
        If underlyingStudyConfig.study Is studyConfig.underlyingStudy Then Exit For
    Next
    studyConfig.chartRegionName = underlyingStudyConfig.chartRegionName
ElseIf studyConfig.chartRegionName = RegionNameCustom Then
    studyConfig.chartRegionName = studyConfig.study.instancePath
End If
End Sub

Private Function findRegion( _
                ByVal regionName As String, _
                ByVal title As String, _
                ByVal incrementUsageCount As Boolean) As ChartRegion
Dim re As RegionEntry

On Error Resume Next
re = mRegions(regionName)
On Error GoTo 0

Set findRegion = re.region

If findRegion Is Nothing Then
    Set findRegion = mChartController.getChartRegion(regionName)
    If findRegion Is Nothing Then
        Set findRegion = mChartController.addChartRegion(20, , regionName)
    End If
    Set re.region = findRegion
    If incrementUsageCount Then re.usageCount = 1
    mRegions.add re, regionName
    findRegion.gridlineSpacingY = 0.8
    findRegion.showGrid = True
    findRegion.setTitle title, vbBlue, Nothing
ElseIf incrementUsageCount Then
    re.usageCount = re.usageCount + 1
    mRegions.remove regionName
    mRegions.add re, regionName
End If
End Function

Private Function includeStudyValueInChart( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudyValueDefinition) As TaskCompletion

Dim lStudyValueHandler As StudyValueHandler
Dim study As study
Dim region As ChartRegion
Dim regionName As String
Dim lTaskCompletion As Tasks.TaskCompletion

Set study = studyConfig.study
If studyValueConfig.chartRegionName = RegionNameDefault Then
    ' then use the study's default region
    regionName = studyConfig.chartRegionName
ElseIf studyValueConfig.chartRegionName = RegionNameCustom Then
    regionName = study.instancePath
    'studyValueConfig.chartRegionName = regionName
Else
    regionName = studyValueConfig.chartRegionName
End If

Set region = findRegion(regionName, study.instanceName, True)

Set lStudyValueHandler = studyConfig.studyValueHandlers.add(Me, _
                                        region, _
                                        studyValueConfig, _
                                        studyValueDef, _
                                        mUpdatePerTick)


Set lTaskCompletion = study.addStudyValueListener( _
                            lStudyValueHandler, _
                            studyValueConfig.valueName, _
                            ReplayNumbers.ReplayAll, _
                            "AddStudyValueListener for value " & studyValueConfig.valueName & " to study " & study.instanceName)

If Not lTaskCompletion Is Nothing Then
    'mOutstandingTasks = mOutstandingTasks + 1
    lTaskCompletion.addTaskCompletionListener Me
    mChartController.suppressDrawing = True
End If
Set includeStudyValueInChart = lTaskCompletion
End Function

Private Sub initialiseChart()

mChartController.suppressDrawing = True

mChartController.clearChart
mChartController.chartBackColor = vbWhite
mChartController.autoscale = True
mChartController.pointerStyle = PointerCrosshairs
mChartController.twipsPerBar = 100
mChartController.showHorizontalScrollBar = True

Set mPeriods = mChartController.Periods

mChartController.suppressDrawing = False

End Sub




