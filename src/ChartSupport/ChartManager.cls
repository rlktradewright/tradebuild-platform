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

Private Const TaskTypeStartStudy As Long = 1
Private Const TaskTypeReplayBars As Long = 2
Private Const TaskTypeAddValueListener As Long = 3

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Public Type RegionEntry
    region              As ChartSkil.ChartRegion
    usageCount          As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mChartController As ChartSkil.chartController
Attribute mChartController.VB_VarHelpID = -1
Private mStudyManager As StudiesSupp.studyManager

Private mSourceStudies As Collection

Private mRegions As Collection

Private mStudyConfigurations As StudyConfigurations

Private mBackfilling As Boolean
Private mFirstBackfill As Boolean   ' indicates if this is the first lot of
                                    ' historic bars that has been played through
                                    ' this chart
                                            
Private mUpdatePerTick As Boolean

Private mInitialNumberOfBars As Long

Private mPeriods As ChartSkil.Periods

Private mInitialised As Boolean

Private mHorizontalLineKeys As Collection

Private mHighPrice As Double
Private mLowPrice As Double
Private mPrevClosePrice As Double

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mSourceStudies = New Collection
Set mRegions = New Collection
Set mStudyConfigurations = New StudyConfigurations
Set mHorizontalLineKeys = New Collection
mUpdatePerTick = True
mFirstBackfill = True

End Sub

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As Tasks.TaskCompletionEvent)
'mOutstandingTasks = mOutstandingTasks - 1
'If mOutstandingTasks = 0 Then mChartController.suppressDrawing = False
mChartController.suppressDrawing = False

If ev.Data = TaskTypeReplayBars Then
    mBackfilling = False
    mFirstBackfill = False
    'updatePreviousBar
End If
End Sub

'================================================================================
' mChartController Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get chartController() As ChartSkil.chartController
Set chartController = mChartController
End Property

Public Property Get regionNames() As String()
Dim names() As String
Dim region As ChartSkil.ChartRegion
Dim re As RegionEntry
Dim i As Long

ReDim names(mRegions.count) As String

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

'Public Property Get timeframeCaption() As String
'Dim units As String
'Select Case mPeriodUnits
'Case TimePeriodUnits.TimePeriodMinute
'    timeframeCaption = IIf(mPeriodLength = 1, "1 Min", mPeriodLength & " Mins")
'Case TimePeriodUnits.TimePeriodHour
'    timeframeCaption = IIf(mPeriodLength = 1, "1 Hour", mPeriodLength & " Hrs")
'Case TimePeriodUnits.TimePeriodDay
'    timeframeCaption = IIf(mPeriodLength = 1, "Daily", mPeriodLength & " Days")
'Case TimePeriodUnits.TimePeriodWeek
'    timeframeCaption = IIf(mPeriodLength = 1, "Weekly", mPeriodLength & " Wks")
'Case TimePeriodUnits.TimePeriodMonth
'    timeframeCaption = IIf(mPeriodLength = 1, "Monthly", mPeriodLength & " Mths")
'Case TimePeriodUnits.TimePeriodLunarMonth
'
'Case TimePeriodUnits.TimePeriodYear
'    timeframeCaption = IIf(mPeriodLength = 1, "Yearly", mPeriodLength & " Yrs")
'End Select
'End Property

Public Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addChartRegion( _
                    ByVal name As String, _
                    ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double) As ChartSkil.ChartRegion
Dim re As RegionEntry
Set addChartRegion = mChartController.addChartRegion(percentheight, minimumPercentHeight, name)
Set re.region = addChartRegion
mRegions.add re, name
End Function

Public Function addInput( _
                ByVal sourceName As String, _
                ByVal inputName As String, _
                ByVal description As String, _
                ByVal valueType As StudiesSupp.StudyInputTypes, _
                ByVal isDefault As Boolean, _
                ByVal chartRegionName As String) As Long
Dim sourceStudy As TradeBuildSP.IStudy
Dim studyConfig As StudyConfiguration
Dim studyValueConfig As StudyValueConfiguration

On Error Resume Next
Set sourceStudy = mSourceStudies(sourceName)
On Error GoTo 0

If sourceStudy Is Nothing Then
    Set sourceStudy = mStudyManager.addSource(sourceName)
    mSourceStudies.add sourceStudy, sourceName
    
    Set studyConfig = New StudyConfiguration
    studyConfig.instanceName = sourceName
    studyConfig.instanceFullyQualifiedName = sourceName
    studyConfig.studyId = sourceName
    studyConfig.name = sourceName
    studyConfig.studyDefinition = sourceStudy.studyDefinition
    mStudyConfigurations.add studyConfig
Else
    For Each studyConfig In mStudyConfigurations
        If studyConfig.studyId = sourceStudy.id Then Exit For
    Next
End If

addInput = mStudyManager.addInput(sourceStudy, inputName, description, valueType, isDefault)

Set studyValueConfig = studyConfig.studyValueConfigurations.add(inputName)
studyValueConfig.chartRegionName = chartRegionName

End Function

Public Function addStudy( _
                ByVal studyConfig As StudyConfiguration) As TradeBuildSP.IStudy
Dim study As TradeBuildSP.IStudy
Dim studyId As String
Dim studyValueConfig As StudyValueConfiguration
Dim underlyingStudyConfig As StudyConfiguration
Dim regionName As String
Dim region As ChartSkil.ChartRegion
Dim studyHorizRule As StudyHorizontalRule
Dim horizRulesLineSeries As ChartSkil.LineSeries
Dim line As ChartSkil.line
Dim studyValueDefs As StudiesSupp.StudyValueDefinitions
Dim studyValueDef As StudiesSupp.StudyValueDefinition
Dim i As Long

If studyConfig.underlyingStudyId = "" Then
    ' this is a default study configuration
    'studyConfig.underlyingStudyId = mBars.Id !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'studyConfig.inputValueNames = "Close"
End If

Set study = mStudyManager.addStudy(studyConfig.name, _
                            studyId, _
                            studyConfig.underlyingStudyId, _
                            studyConfig.inputValueNames, _
                            studyConfig.parameters, _
                            studyConfig.serviceProviderName)

studyConfig.instanceName = study.instanceName
studyConfig.instanceFullyQualifiedName = study.instancePath
studyConfig.studyId = study.id

If studyConfig.chartRegionName = RegionNameDefault Then
    ' we will use the region for the underlying study
    For Each underlyingStudyConfig In mStudyConfigurations
        If underlyingStudyConfig.studyId = studyConfig.underlyingStudyId Then Exit For
    Next
    regionName = underlyingStudyConfig.chartRegionName
    studyConfig.chartRegionName = regionName
ElseIf studyConfig.chartRegionName = RegionNameCustom Then
    regionName = study.instancePath
    studyConfig.chartRegionName = regionName
Else
    regionName = studyConfig.chartRegionName
End If
Set region = findRegion(regionName, study.instanceName, False)
    
Set studyValueDefs = studyConfig.studyDefinition.StudyValueDefinitions

For Each studyValueConfig In studyConfig.studyValueConfigurations
    If studyValueConfig.includeInChart Then
        includeStudyValueInChart studyConfig, _
                                study, _
                                regionName, _
                                studyValueConfig, _
                                studyValueDefs.item(studyValueConfig.valueName)
    End If
Next

If studyConfig.studyHorizontalRules.count > 0 Then
    Set horizRulesLineSeries = region.addLineSeries(LayerNumbers.LayerGrid + 1)
    horizRulesLineSeries.extended = True
    horizRulesLineSeries.extendAfter = True
    horizRulesLineSeries.extendBefore = True
    For Each studyHorizRule In studyConfig.studyHorizontalRules
        horizRulesLineSeries.Color = studyHorizRule.Color
        horizRulesLineSeries.style = studyHorizRule.style
        horizRulesLineSeries.thickness = studyHorizRule.thickness
        Set line = horizRulesLineSeries.addLine
        line.point1 = region.newPoint(0, studyHorizRule.y, CoordsRelative, CoordsLogical)
        line.point2 = region.newPoint(100, studyHorizRule.y, CoordsRelative, CoordsLogical)
    Next
    studyConfig.horizontalRulesLineSeries = horizRulesLineSeries
End If

mStudyConfigurations.add studyConfig
startStudy studyId

Set addStudy = study
End Function

Public Sub clearChart()
mChartController.clearChart
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

Friend Function getPeriod(ByVal pTimestamp As Date) As ChartSkil.Period
Static sPeriod As ChartSkil.Period
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
    Set getPeriod = mChartController.addperiod(pTimestamp)
    mChartController.scrollX 1
End If
Set sPeriod = getPeriod
sTimestamp = pTimestamp
End Function

Friend Function getSpecialValue(ByVal valueType As SpecialValues) As Variant

Select Case valueType
Case SpecialValues.SVCurrentSessionEndTime
    getSpecialValue = mChartController.CurrentSessionEndTime
Case SpecialValues.SVCurrentSessionStartTime
    getSpecialValue = mChartController.CurrentSessionStartTime
Case SpecialValues.SVHighPrice
    getSpecialValue = mHighPrice
Case SpecialValues.SVLowPrice
    getSpecialValue = mLowPrice
Case SpecialValues.SVPreviousClosePrice
    getSpecialValue = mPrevClosePrice
End Select
End Function

Friend Sub initialise( _
                ByVal pStudyManager As StudiesSupp.studyManager, _
                ByVal pChartController As ChartSkil.chartController)
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

'Public Sub notifyHigh(ByVal value As Double)
'mHighPrice = value
'End Sub
'
'Public Sub notifyLow(ByVal value As Double)
'mLowPrice = value
'End Sub
'
'Public Sub notifyPreviousClose(ByVal value As Double)
'mPrevClosePrice = value
'End Sub
'
'Public Sub notifyTrade(ByVal value As Double)
''If mUpdatePerTick Then mChartBar.tick value
'End Sub
'
'Public Sub notifyVolume(ByVal value As Long)
'If mUpdatePerTick Then setVolume value
'End Sub

Public Function removeStudy( _
                ByVal studyConfig As StudyConfiguration)
Dim studyValueConfig  As StudyValueConfiguration
Dim svh As StudyValueHandler
Dim horizRulesLineSeries As ChartSkil.LineSeries
Dim regionName As String
Dim re As RegionEntry
Dim region As ChartSkil.ChartRegion

mChartController.suppressDrawing = True
For Each svh In studyConfig.studyValueHandlers
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

'Public Sub showStudyPickerForm()
'If mTicker.State = TickerStateRunning Then showStudyPicker mTicker, Me
'End Sub
'
'Public Sub syncStudyPickerForm()
'If mTicker.State = TickerStateRunning Then syncStudyPicker mTicker, Me
'End Sub
'
'Public Sub unsyncStudyPickerForm()
'unsyncStudyPicker
'End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function addStudyDataPoint( _
                ByVal dataSeries As ChartSkil.DataPointSeries, _
                ByVal timestamp As Date) As ChartSkil.DataPoint
Dim Period As ChartSkil.Period

Set Period = getPeriod(timestamp)

On Error Resume Next
Set addStudyDataPoint = dataSeries.item(Period.periodNumber)
On Error GoTo 0

If addStudyDataPoint Is Nothing Then
    Set addStudyDataPoint = dataSeries.addDataPoint(Period.periodNumber)
End If

End Function

Private Function findRegion( _
                ByVal regionName As String, _
                ByVal title As String, _
                ByVal incrementUsageCount As Boolean) As ChartSkil.ChartRegion
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

'Private Function GenerateTimeframeKey() As String
'GenerateTimeframeKey = mPeriodLengthMinutes & "min"
'End Function

Private Sub includeStudyValueInChart( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal study As TradeBuildSP.IStudy, _
                ByVal defaultRegionName As String, _
                ByVal studyValueConfig As StudyValueConfiguration, _
                ByVal studyValueDef As StudiesSupp.StudyValueDefinition)

Dim lStudyValueHandler As StudyValueHandler
Dim region As ChartSkil.ChartRegion
Dim regionName As String
Dim lTaskCompletion As Tasks.TaskCompletion

If studyValueConfig.chartRegionName = RegionNameDefault Then
    ' then use the study's default region
    regionName = defaultRegionName
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
                            0, _
                            , _
                            TaskTypeAddValueListener)

If Not lTaskCompletion Is Nothing Then
    'mOutstandingTasks = mOutstandingTasks + 1
    lTaskCompletion.addTaskCompletionListener Me
    mChartController.suppressDrawing = True
End If
End Sub

Private Sub initialiseChart()
Dim re As RegionEntry

If mInitialised Then Exit Sub

mChartController.suppressDrawing = True

mChartController.clearChart
mChartController.chartBackColor = vbWhite
mChartController.autoscale = True
mChartController.showCrosshairs = True
mChartController.twipsPerBar = 100
mChartController.showHorizontalScrollBar = True

Set mPeriods = mChartController.Periods

mChartController.suppressDrawing = False

mInitialised = True

End Sub

Private Sub setVolume(ByVal size As Long)
'mVolumePoint.dataValue = size
'If mVolumePoint.dataValue >= mPrevBarVolume Then
'    mVolumePoint.lineColor = vbGreen
'Else
'    mVolumePoint.lineColor = vbRed
'End If
End Sub

Private Function startStudy( _
                ByVal studyId As String) As Boolean
Dim lTaskCompletion As Tasks.TaskCompletion
Set lTaskCompletion = mStudyManager.startStudy(studyId, 0, , TaskTypeStartStudy)

If lTaskCompletion Is Nothing Then Exit Function

startStudy = True
'mOutstandingTasks = mOutstandingTasks + 1
lTaskCompletion.addTaskCompletionListener Me
mChartController.suppressDrawing = True
End Function




