VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timeframes"
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

Event TimeFrameAdded(ByVal pTimeframe As Timeframe)

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type TimeframeTableEntry
    timeperiod              As Long
    timeperiodUnit          As TradeBuild.TimePeriodUnits
    key                     As String
    theTimeFrame            As Timeframe
    numberOfBarsToFetch     As Long
    historicalDataLoaded    As Boolean
End Type

'================================================================================
' Member variables
'================================================================================

Private mTimeframeTable() As TimeframeTableEntry
Private mNextTimeframeTableIndex As Long

Private mTicker As TradeBuild.Ticker

Private mContract As TradeBuild.Contract

Private mChangeListeners As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mTimeframeTable(10) As TimeframeTableEntry
Set mChangeListeners = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "Timeframes terminated"
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Contract(ByVal value As Contract)
Dim tte As TimeframeTableEntry
Dim i As Long

Set mContract = value
For i = 0 To mNextTimeframeTableIndex - 1
    tte = mTimeframeTable(i)
    tte.theTimeFrame.Contract = mContract
Next
End Property

Public Property Get Count() As Long
Count = mNextTimeframeTableIndex
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Set mTicker = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                            ByVal timeperiod As Long, _
                            ByVal timeperiodUnit As TradeBuild.TimePeriodUnits, _
                            ByRef key As String, _
                            ByVal numberOfBarsToFetch As Long, _
                            Optional ByVal SaveIntervalSeconds As Long = -1&, _
                            Optional ByVal IncludeBidAndAsk As Boolean, _
                            Optional ByVal excludeCurrentBar As Boolean = False) As Timeframe

Dim tte As TimeframeTableEntry
Dim timeframeIndex As Long

timeframeIndex = getTimeframeIndex(timeperiod, timeperiodUnit, key)
If timeframeIndex >= 0 Then
    Set add = mTimeframeTable(timeframeIndex).theTimeFrame
    If key = "" Then key = mTimeframeTable(timeframeIndex).key
Else
    Set add = New Timeframe
    Set tte.theTimeFrame = add
    tte.timeperiod = timeperiod
    tte.timeperiodUnit = timeperiodUnit
    tte.numberOfBarsToFetch = numberOfBarsToFetch
    If key <> "" Then
        tte.key = key
    Else
        tte.key = generateKey(timeperiod, timeperiodUnit)
    End If
    tte.theTimeFrame.initialise tte.key, _
                                mTicker, _
                                timeperiod, _
                                timeperiodUnit, _
                                numberOfBarsToFetch, _
                                SaveIntervalSeconds, _
                                IncludeBidAndAsk, _
                                excludeCurrentBar
    
    If Not mContract Is Nothing Then tte.theTimeFrame.Contract = mContract
    If mNextTimeframeTableIndex > UBound(mTimeframeTable) Then
        ReDim Preserve mTimeframeTable(UBound(mTimeframeTable) + 10) As TimeframeTableEntry
    End If
    mTimeframeTable(mNextTimeframeTableIndex) = tte
    mNextTimeframeTableIndex = mNextTimeframeTableIndex + 1
    fireChange CollItemAdded, tte.theTimeFrame
    RaiseEvent TimeFrameAdded(add)
End If

End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Friend Sub finish()
Dim i As Long

For i = 0 To mNextTimeframeTableIndex - 1
    mTimeframeTable(i).theTimeFrame.finish
Next
Set mTicker = Nothing
End Sub

Public Function Item(key As String) As Timeframe
Dim i As Long

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).key = key Then
        Set Item = mTimeframeTable(i).theTimeFrame
        Exit For
    End If
Next
End Function

Public Sub remove(key As String)
Dim i As Long

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).key = key Then
        mTimeframeTable(i).theTimeFrame.finish
        clearTimeframeTableEntry i
        fireChange CollItemRemoved, mTimeframeTable(i).theTimeFrame
        Exit For
    End If
Next
End Sub

Public Sub RemoveAll()
Dim i As Long

For i = 0 To mNextTimeframeTableIndex - 1
    mTimeframeTable(i).theTimeFrame.finish
    clearTimeframeTableEntry i
    fireChange CollItemRemoved, mTimeframeTable(i).theTimeFrame
Next
End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearTimeframeTableEntry(ByVal index As Long)
mTimeframeTable(index).historicalDataLoaded = False
mTimeframeTable(index).key = ""
mTimeframeTable(index).numberOfBarsToFetch = 0
Set mTimeframeTable(index).theTimeFrame = Nothing
mTimeframeTable(index).timeperiod = 0
mTimeframeTable(index).timeperiodUnit = 0
End Sub

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Timeframe)
Dim listener As CollectionChangeListener
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For Each listener In mChangeListeners
    listener.Change Change
Next
End Sub

Private Function generateKey( _
                ByVal timeperiod As Long, _
                ByVal timeperiodUnit As TradeBuild.TimePeriodUnits) As String
generateKey = timeperiod
Select Case timeperiodUnit
Case TimePeriodMinute
    generateKey = generateKey & "Mins"
Case TimePeriodHour
    generateKey = generateKey & "Hours"
Case TimePeriodDay
    generateKey = generateKey & "Days"
Case TimePeriodWeek
    generateKey = generateKey & "Weeks"
Case TimePeriodMonth
    generateKey = generateKey & "Months"
Case TimePeriodLunarMonth
    generateKey = generateKey & "LunarMonths"
Case TimePeriodYear
    generateKey = generateKey & "Years"
End Select
End Function

Private Function getTimeframeIndex( _
                ByVal timeperiod As Long, _
                ByVal timeperiodUnit As TradeBuild.TimePeriodUnits, _
                ByRef key As String) As Long
Dim i As Long

getTimeframeIndex = -1
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).timeperiod = timeperiod And _
        mTimeframeTable(i).timeperiodUnit = timeperiodUnit And _
        (mTimeframeTable(i).key = key Or key = "") _
    Then
        getTimeframeIndex = i
        Exit Function
    End If
Next
End Function


