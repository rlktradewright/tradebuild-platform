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

Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName As String = "Timeframes"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TimeframeTableEntry
    barTimePeriod           As TimePeriod
    key                     As String
    theTimeFrame            As Timeframe
    numberOfBarsToFetch     As Long
    fromTime                As Date
    toTime                  As Date
    includeBarsOutsideSession   As Boolean
    excludeCurrentbar       As Boolean
    saveIntervalSeconds     As Long
    writeBidAndAskBars      As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mTimeframeTable() As TimeframeTableEntry
Private mNextTimeframeTableIndex As Long

Private mTicker As Ticker

Private mContract As Contract

Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTimeframeTable(1) As TimeframeTableEntry
Set mChangeListeners = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "Timeframes terminated"
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

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
                ByVal barTimePeriod As TimePeriod, _
                ByRef key As String, _
                ByVal numberOfBarsToFetch As Long, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal excludeCurrentbar As Boolean) As Timeframe

Select Case barTimePeriod.units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume, _
        TimePeriodTickMovement
Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "add", _
                "Time period units not supported"
End Select

Set add = addTimeframe(barTimePeriod, _
                    key, _
                    numberOfBarsToFetch, _
                    0, _
                    0, _
                    includeBarsOutsideSession, _
                    excludeCurrentbar, _
                    Nothing, _
                    0, _
                    False)

End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

''
' Adds a timeframe of a specified length and fetches historical data for
' a specified period.
'
' @return
'   A reference to the new <code>Timeframe</code> object.
' @param timePeriod
'   The bar length for this timeframe.
' @param timePeriodUnit
'   The unit in which this timeframe is measured.
' @param key
'   A unique identifier for this timeframe. This may be left blank, in which
'   case a unique key is allocated by the system.
' @param numberOfBarsToFetch
'   The maximum number of bars of historical data to be fetched.
' @param fromTime
'   The earliest time for which historical data is to be fetched. This is specified
'   in local time, NOT the timezone for the exchange.
' @param fromTime
'   The latest time for which historical data is to be fetched. This is specified
'   in local time, NOT the timezone for the exchange.
' @see
'
'@/
Public Function addHistorical( _
                ByVal barTimePeriod As TimePeriod, _
                ByRef key As String, _
                ByVal numberOfBarsToFetch As Long, _
                Optional ByVal fromTime As Date, _
                Optional ByVal toTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean) As Timeframe

Select Case barTimePeriod.units
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume, _
        TimePeriodTickMovement
Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "add", _
                "Time period units not supported"
    
End Select

If fromTime > toTime Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild26" & "." & "Timeframes" & ":" & "addHistorical", _
            "fromTime must be earlier than toTime"
End If

Set addHistorical = addTimeframe(barTimePeriod, _
                    key, _
                    numberOfBarsToFetch, _
                    fromTime, _
                    toTime, _
                    includeBarsOutsideSession, _
                    False, _
                    Nothing, _
                    0, _
                    False)

End Function

Friend Function addTimeframe( _
                ByVal barTimePeriod As TimePeriod, _
                ByRef key As String, _
                ByVal numberOfBarsToFetch As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal excludeCurrentbar As Boolean, _
                ByVal barWriter As HistoricDataWriter, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean) As Timeframe

Dim tte As TimeframeTableEntry
Dim timeframeIndex As Long

timeframeIndex = getTimeframeIndex(barTimePeriod, _
                                numberOfBarsToFetch, _
                                fromTime, _
                                toTime, _
                                includeBarsOutsideSession, _
                                saveIntervalSeconds, _
                                writeBidAndAskBars)
If timeframeIndex >= 0 Then
    Set addTimeframe = mTimeframeTable(timeframeIndex).theTimeFrame
    Exit Function
End If

Set addTimeframe = New Timeframe
Set tte.theTimeFrame = addTimeframe
Set tte.barTimePeriod = barTimePeriod
tte.numberOfBarsToFetch = numberOfBarsToFetch
tte.fromTime = fromTime
tte.toTime = toTime
tte.includeBarsOutsideSession = includeBarsOutsideSession
tte.writeBidAndAskBars = writeBidAndAskBars
tte.saveIntervalSeconds = saveIntervalSeconds

If key <> "" Then
    tte.key = key
Else
    tte.key = generateKey(barTimePeriod, _
                        numberOfBarsToFetch, _
                        fromTime, _
                        toTime, _
                        includeBarsOutsideSession, _
                        excludeCurrentbar, _
                        saveIntervalSeconds, _
                        writeBidAndAskBars)
End If
tte.theTimeFrame.initialise mContract, _
                            tte.key, _
                            mTicker, _
                            barTimePeriod, _
                            numberOfBarsToFetch, _
                            fromTime, _
                            toTime, _
                            includeBarsOutsideSession, _
                            excludeCurrentbar, _
                            barWriter, _
                            saveIntervalSeconds, _
                            writeBidAndAskBars

If mNextTimeframeTableIndex > UBound(mTimeframeTable) Then
    ReDim Preserve mTimeframeTable(2 * (UBound(mTimeframeTable) + 1) - 1) As TimeframeTableEntry
End If
mTimeframeTable(mNextTimeframeTableIndex) = tte
mNextTimeframeTableIndex = mNextTimeframeTableIndex + 1
fireChange CollItemAdded, tte.theTimeFrame

End Function

Friend Sub finish()
RemoveAll
Set mTicker = Nothing
End Sub

Public Function item(key As String) As Timeframe
Dim i As Long

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).key = key Then
        Set item = mTimeframeTable(i).theTimeFrame
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
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearTimeframeTableEntry(ByVal index As Long)
mTimeframeTable(index).key = ""
mTimeframeTable(index).numberOfBarsToFetch = 0
Set mTimeframeTable(index).theTimeFrame = Nothing
Set mTimeframeTable(index).barTimePeriod = Nothing
End Sub

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Timeframe)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub

Private Function generateKey( _
                ByVal barTimePeriod As TimePeriod, _
                ByVal numberOfBarsToFetch As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal excludeCurrentbar As Boolean, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean) As String
generateKey = barTimePeriod.toString & _
            "(" & numberOfBarsToFetch & _
            IIf(fromTime <> 0, "," & Format(fromTime, "yyyymmddhhnnss"), "") & _
            IIf(toTime <> 0, "," & Format(toTime, "yyyymmddhhnnss"), "") & _
            ")" & _
            IIf(includeBarsOutsideSession, "I", "") & _
            IIf(excludeCurrentbar, "X", "") & _
            IIf(saveIntervalSeconds > 0, "S" & saveIntervalSeconds, "") & _
            IIf(writeBidAndAskBars, "W", "")
            
End Function

Private Function getTimeframeIndex( _
                ByVal barTimePeriod As TimePeriod, _
                ByVal numberOfBarsToFetch As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByRef includeBarsOutsideSession As Boolean, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean) As Long
Dim i As Long

getTimeframeIndex = -1
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).barTimePeriod Is barTimePeriod And _
        mTimeframeTable(i).includeBarsOutsideSession = includeBarsOutsideSession And _
        mTimeframeTable(i).numberOfBarsToFetch = numberOfBarsToFetch And _
        mTimeframeTable(i).fromTime = fromTime And _
        mTimeframeTable(i).toTime = toTime And _
        mTimeframeTable(i).saveIntervalSeconds = saveIntervalSeconds And _
        mTimeframeTable(i).writeBidAndAskBars = writeBidAndAskBars _
    Then
        getTimeframeIndex = i
        Exit Function
    End If
Next
End Function

