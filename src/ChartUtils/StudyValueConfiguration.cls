VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueConfiguration"
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyValueConfiguration"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mIncludeInChart As Boolean
Private mValueName As String

Private mBarStyle As barStyle
Private mDataPointStyle As dataPointStyle
Private mLineStyle As lineStyle
Private mTextStyle As textStyle

Private mChartRegionName As String
Private mLayer As Long
Private mOffset As Dimension

Private mBarFormatterFactory As BarFormatterFactory
Private mBars As Bars

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

mLayer = LayerNumbers.LayerHighestUser
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

Public Property Get BarFormatterFactory() As BarFormatterFactory
Set BarFormatterFactory = mBarFormatterFactory
End Property

Public Property Get Bars() As Bars
Set Bars = mBars
End Property
    
Public Property Let barStyle(ByVal value As barStyle)
Set mBarStyle = value
End Property

Public Property Get barStyle() As barStyle
Set barStyle = mBarStyle
End Property

Public Property Let chartRegionName(ByVal value As String)
mChartRegionName = value
End Property

Public Property Get chartRegionName() As String
chartRegionName = mChartRegionName
End Property

Public Property Let dataPointStyle(ByVal value As dataPointStyle)
Set mDataPointStyle = value
End Property

Public Property Get dataPointStyle() As dataPointStyle
Set dataPointStyle = mDataPointStyle
End Property

Public Property Let includeInChart(ByVal value As Boolean)
mIncludeInChart = value
End Property

Public Property Get includeInChart() As Boolean
includeInChart = mIncludeInChart
End Property

Public Property Let layer(ByVal value As Long)
mLayer = value
End Property

Public Property Get layer() As Long
layer = mLayer
End Property

Public Property Let lineStyle(ByVal value As lineStyle)
Set mLineStyle = value
End Property

Public Property Get lineStyle() As lineStyle
Set lineStyle = mLineStyle
End Property

Public Property Let offset(ByVal value As Dimension)
mOffset = value
End Property

Public Property Get offset() As Dimension
offset = mOffset
End Property

Public Property Let textStyle(ByVal value As textStyle)
Set mTextStyle = value
End Property

Public Property Get textStyle() As textStyle
Set textStyle = mTextStyle
End Property

Friend Property Let valueName(ByVal value As String)
mValueName = value
End Property

Public Property Get valueName() As String
valueName = mValueName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub SetBarFormatterFactory(ByVal formatterFactory As BarFormatterFactory, ByVal Bars As Bars)
If Not formatterFactory Is Nothing And Bars Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetBarFormatterFactory", _
            "A Bars object must be supplied"
End If
Set mBarFormatterFactory = formatterFactory
Set mBars = Bars
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


