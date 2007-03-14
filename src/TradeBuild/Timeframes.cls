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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TimeframeTableEntry
    TimePeriod              As Long
    timePeriodUnit          As TimePeriodUnits
    key                     As String
    theTimeFrame            As Timeframe
    numberOfBarsToFetch     As Long
    includeBarsOutsideSession   As Boolean
    excludeCurrentbar       As Boolean
    saveIntervalSeconds     As Long
    writeBidAndAskBars      As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildAPIRef           As WeakReference

Private mTimeframeTable() As TimeframeTableEntry
Private mNextTimeframeTableIndex As Long

Private mTicker As TradeBuild25.Ticker

Private mContract As TradeBuild25.Contract

Private mChangeListeners As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTimeframeTable(10) As TimeframeTableEntry
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

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = weakrefs.createWeakReference(value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
                ByVal TimePeriod As Long, _
                ByVal timePeriodUnit As TimePeriodUnits, _
                ByRef key As String, _
                ByVal numberOfBarsToFetch As Long, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal excludeCurrentbar As Boolean, _
                Optional ByVal saveIntervalSeconds As Long = -1&, _
                Optional ByVal writeBidAndAskBars As Boolean) As Timeframe

Dim tte As TimeframeTableEntry
Dim timeframeIndex As Long

If saveIntervalSeconds > 0 And saveIntervalSeconds < 5 Then _
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild25.Timeframes::add", _
            "Save interval cannot be less than 5 seconds"

timeframeIndex = getTimeframeIndex(TimePeriod, _
                                timePeriodUnit, _
                                numberOfBarsToFetch, _
                                includeBarsOutsideSession, _
                                saveIntervalSeconds, _
                                writeBidAndAskBars)
If timeframeIndex >= 0 Then
    Set add = mTimeframeTable(timeframeIndex).theTimeFrame
    Exit Function
End If

Set add = New Timeframe
add.tradeBuildAPI = tb
Set tte.theTimeFrame = add
tte.TimePeriod = TimePeriod
tte.timePeriodUnit = timePeriodUnit
tte.numberOfBarsToFetch = numberOfBarsToFetch
tte.includeBarsOutsideSession = includeBarsOutsideSession
tte.writeBidAndAskBars = writeBidAndAskBars
tte.saveIntervalSeconds = saveIntervalSeconds

If key <> "" Then
    tte.key = key
Else
    tte.key = generateKey(TimePeriod, _
                        timePeriodUnit, _
                        numberOfBarsToFetch, _
                        includeBarsOutsideSession, _
                        excludeCurrentbar, _
                        saveIntervalSeconds, _
                        writeBidAndAskBars)
End If
tte.theTimeFrame.Initialise mContract, _
                            tte.key, _
                            mTicker, _
                            TimePeriod, _
                            timePeriodUnit, _
                            numberOfBarsToFetch, _
                            includeBarsOutsideSession, _
                            excludeCurrentbar, _
                            saveIntervalSeconds, _
                            writeBidAndAskBars

If mNextTimeframeTableIndex > UBound(mTimeframeTable) Then
    ReDim Preserve mTimeframeTable(UBound(mTimeframeTable) + 10) As TimeframeTableEntry
End If
mTimeframeTable(mNextTimeframeTableIndex) = tte
mNextTimeframeTableIndex = mNextTimeframeTableIndex + 1
fireChange CollItemAdded, tte.theTimeFrame

End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

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
mTimeframeTable(index).TimePeriod = 0
mTimeframeTable(index).timePeriodUnit = 0
End Sub

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Timeframe)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub

Private Function generateKey( _
                ByVal TimePeriod As Long, _
                ByVal timePeriodUnit As TimePeriodUnits, _
                ByVal numberOfBarsToFetch As Long, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal excludeCurrentbar As Boolean, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean) As String
generateKey = TimePeriod & _
            TimePeriodUnitsToString(timePeriodUnit) & _
            "(" & numberOfBarsToFetch & ")" & _
            IIf(includeBarsOutsideSession, "I", "") & _
            IIf(excludeCurrentbar, "X", "") & _
            IIf(saveIntervalSeconds > 0, "S" & saveIntervalSeconds, "") & _
            IIf(writeBidAndAskBars, "W", "")
            
End Function

Private Function getTimeframeIndex( _
                ByVal TimePeriod As Long, _
                ByVal timePeriodUnit As TimePeriodUnits, _
                ByVal numberOfBarsToFetch As Long, _
                ByRef includeBarsOutsideSession As Boolean, _
                ByVal saveIntervalSeconds As Long, _
                ByVal writeBidAndAskBars As Boolean) As Long
Dim i As Long

getTimeframeIndex = -1
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).TimePeriod = TimePeriod And _
        mTimeframeTable(i).timePeriodUnit = timePeriodUnit And _
        mTimeframeTable(i).includeBarsOutsideSession = includeBarsOutsideSession And _
        mTimeframeTable(i).numberOfBarsToFetch = numberOfBarsToFetch And _
        mTimeframeTable(i).saveIntervalSeconds = saveIntervalSeconds And _
        mTimeframeTable(i).writeBidAndAskBars = writeBidAndAskBars _
    Then
        getTimeframeIndex = i
        Exit Function
    End If
Next
End Function

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function

