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

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'Enum DisplayModes
'    DisplayNone
'    displayAsPoints
'    DisplayAsLines
'    DisplayAsSteppedLines
'    displayAsHistogram
'End Enum
'
'Enum LineStyles
'    LineSolid = vbSolid
'    LineDash = vbDash
'    LineDot = vbDot
'    LineDashDot = vbDashDot
'    LineDashDotDot = vbDashDotDot
'    LineInvisible = vbInvisible
'    LineInsideSolid = vbInsideSolid
'End Enum

Enum ConditionalOperators
    OpTrue
    OpLessThan
    OpEqual
    OpGreaterThan
    OpBetween
    OpLessThanPrevious
    OpEqualPrevious
    OpGreaterThanPrevious
    OpStartsWith
    OpContains
    OpEndsWith
    OpAnd
    OpOr
    OpXor
End Enum

Enum SpecialValues
    SVCurrentSessionStartTime
    SVCurrentSessionEndTime
    SVHighPrice
    SVLowPrice
    SVPreviousClosePrice
End Enum

Enum StudyRenderingActions
    SRActSetPeriodBackgroundColor
    SRActSetValueColor
End Enum

'================================================================================
' Types
'================================================================================

Public Type ConditionalAction
    operator        As ConditionalOperators
    not             As Boolean
    value1          As Variant
    isSpecial1      As Boolean
    isSpecial2      As Boolean
    value2          As Variant
    action          As StudyRenderingActions
    actionValue     As Variant
    stopIfTrue      As Boolean      ' if set, then if this condition is true,
                                    ' then no further conditional actions are
                                    ' performed
End Type
'================================================================================
' Member variables
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Sub addChartRegion( _
                ByVal chartGUID As String, _
                ByVal name As String, _
                ByVal title As String, _
                Optional ByVal autoscale As Boolean = True, _
                Optional ByVal gridlineSpacingY As Double = 0.8, _
                Optional ByVal gridTextColor As Long = vbBlack, _
                Optional ByVal initialPercentHeight As Double = 20#, _
                Optional ByVal minimumPercentHeight As Double = 0#, _
                Optional ByVal integerYScale As Boolean = False, _
                Optional ByVal minYScale As Single = 0#, _
                Optional ByVal maxYScale As Single = 1#, _
                Optional ByVal showGrid As Boolean = True, _
                Optional ByVal showGridText As Boolean = False)
                
End Sub

Public Sub addConditionalAction( _
                ByVal studyValueGuid As String, _
                pConditionalAction As ConditionalAction)
                
End Sub

Public Sub addHorizontalLine( _
                ByVal chartGUID As String, _
                ByVal chartRegionName As String, _
                ByVal y As Single, _
                Optional ByVal lineStyle As LineStyles = LineSolid, _
                Optional ByVal lineThickness As Long = 1, _
                Optional ByVal lineColor As Long = vbBlack, _
                Optional ByVal layer As Long = 6)
                
End Sub

Public Function addTimeframe( _
                ByVal timeframeLength As Long, _
                ByVal timeframeUnits As String) As String
                
End Function

Public Function includeStudyValueInChart( _
                ByVal chartGUID As String, _
                ByVal study As TradeBuild.study, _
                ByVal studyValueName As String, _
                ByVal chartRegionName As String, _
                ByVal includeInAutoscale As Boolean, _
                ByVal layer As Long, _
                ByVal displayAs As DisplayModes, _
                ByVal lineStyle As LineStyles, _
                ByVal lineThickness As Long, _
                ByVal lineColor As Long, _
                ByVal histogramBarWidth As Single) As String
                
End Function

Public Function showEmbeddedChart( _
                ByVal timeframeGUID As String, _
                ByVal isDefaultChart As Boolean) As String
                
End Function

Public Function showFloatingChart( _
                ByVal timeframeGUID As String) As String
                
End Function

'================================================================================
' Helper Functions
'================================================================================



