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
                ev As CollectionChangeEventData)

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
    BarTimePeriod           As TimePeriod
    Key                     As String
    TheTimeFrame            As Timeframe
    NumberOfBarsToFetch     As Long
    FromTime                As Date
    ToTime                  As Date
    IncludeBarsOutsideSession   As Boolean
    ExcludeCurrentbar       As Boolean
    SaveIntervalSeconds     As Long
    WriteBidAndAskBars      As Boolean
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

Friend Property Let Contract(ByVal Value As Contract)
Dim tte As TimeframeTableEntry
Dim i As Long

Const ProcName As String = "Contract"

On Error GoTo Err

Set mContract = Value
For i = 0 To mNextTimeframeTableIndex - 1
    tte = mTimeframeTable(i)
    tte.TheTimeFrame.Contract = mContract
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Count() As Long
Count = mNextTimeframeTableIndex
End Property

Friend Property Let Ticker(ByVal Value As Ticker)
Set mTicker = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal BarTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal NumberOfBarsToFetch As Long, _
                Optional ByVal IncludeBarsOutsideSession As Boolean, _
                Optional ByVal ExcludeCurrentbar As Boolean) As Timeframe

Const ProcName As String = "Add"

On Error GoTo Err

Select Case BarTimePeriod.Units
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
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Time period units not supported"
End Select

Set Add = AddTimeframe(BarTimePeriod, _
                    Key, _
                    NumberOfBarsToFetch, _
                    0, _
                    0, _
                    IncludeBarsOutsideSession, _
                    ExcludeCurrentbar, _
                    Nothing, _
                    0, _
                    False)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Sub AddCollectionChangeListener( _
                ByVal Value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add Value

'gNotifyExistingCollectionMembers mTimeframes, value, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Adds a timeframe of a specified length and fetches historical Data for
' a specified period.
'
' @return
'   A reference to the new <code>Timeframe</code> object.
' @param timePeriod
'   The bar length for this timeframe.
' @param timePeriodUnit
'   The unit in which this timeframe is measured.
' @param Key
'   A unique identifier for this timeframe. This may be left blank, in which
'   case a unique Key is allocated by the system.
' @param numberOfBarsToFetch
'   The maximum number of bars of historical Data to be fetched.
' @param fromTime
'   The earliest time for which historical Data is to be fetched. This is specified
'   in local time, NOT the timezone for the exchange.
' @param fromTime
'   The latest time for which historical Data is to be fetched. This is specified
'   in local time, NOT the timezone for the exchange.
' @see
'
'@/
Public Function AddHistorical( _
                ByVal BarTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal NumberOfBarsToFetch As Long, _
                Optional ByVal FromTime As Date, _
                Optional ByVal ToTime As Date, _
                Optional ByVal IncludeBarsOutsideSession As Boolean) As Timeframe

Const ProcName As String = "AddHistorical"

On Error GoTo Err

Select Case BarTimePeriod.Units
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
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Time period units not supported"
    
End Select

If FromTime > ToTime Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "fromTime must be earlier than toTime"
End If

Set AddHistorical = AddTimeframe(BarTimePeriod, _
                    Key, _
                    NumberOfBarsToFetch, _
                    FromTime, _
                    ToTime, _
                    IncludeBarsOutsideSession, _
                    False, _
                    Nothing, _
                    0, _
                    False)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Function AddTimeframe( _
                ByVal BarTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal NumberOfBarsToFetch As Long, _
                ByVal FromTime As Date, _
                ByVal ToTime As Date, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal ExcludeCurrentbar As Boolean, _
                ByVal barWriter As HistoricDataWriter, _
                ByVal SaveIntervalSeconds As Long, _
                ByVal WriteBidAndAskBars As Boolean) As Timeframe

Dim tte As TimeframeTableEntry
Dim timeframeIndex As Long

Const ProcName As String = "AddTimeframe"

On Error GoTo Err

timeframeIndex = getTimeframeIndex(BarTimePeriod, _
                                NumberOfBarsToFetch, _
                                FromTime, _
                                ToTime, _
                                IncludeBarsOutsideSession, _
                                SaveIntervalSeconds, _
                                WriteBidAndAskBars)
If timeframeIndex >= 0 Then
    Set AddTimeframe = mTimeframeTable(timeframeIndex).TheTimeFrame
    Exit Function
End If

Set AddTimeframe = New Timeframe
Set tte.TheTimeFrame = AddTimeframe
Set tte.BarTimePeriod = BarTimePeriod
tte.NumberOfBarsToFetch = NumberOfBarsToFetch
tte.FromTime = FromTime
tte.ToTime = ToTime
tte.IncludeBarsOutsideSession = IncludeBarsOutsideSession
tte.WriteBidAndAskBars = WriteBidAndAskBars
tte.SaveIntervalSeconds = SaveIntervalSeconds

If Key <> "" Then
    tte.Key = Key
Else
    tte.Key = generateKey(BarTimePeriod, _
                        NumberOfBarsToFetch, _
                        FromTime, _
                        ToTime, _
                        IncludeBarsOutsideSession, _
                        ExcludeCurrentbar, _
                        SaveIntervalSeconds, _
                        WriteBidAndAskBars)
End If
tte.TheTimeFrame.Initialise mContract, _
                            tte.Key, _
                            mTicker, _
                            BarTimePeriod, _
                            NumberOfBarsToFetch, _
                            FromTime, _
                            ToTime, _
                            IncludeBarsOutsideSession, _
                            ExcludeCurrentbar, _
                            barWriter, _
                            SaveIntervalSeconds, _
                            WriteBidAndAskBars

If mNextTimeframeTableIndex > UBound(mTimeframeTable) Then
    ReDim Preserve mTimeframeTable(2 * (UBound(mTimeframeTable) + 1) - 1) As TimeframeTableEntry
End If
mTimeframeTable(mNextTimeframeTableIndex) = tte
mNextTimeframeTableIndex = mNextTimeframeTableIndex + 1
fireChange CollItemAdded, tte.TheTimeFrame

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

RemoveAll
Set mTicker = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item(Key As String) As Timeframe
Dim i As Long

Const ProcName As String = "Item"

On Error GoTo Err

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).Key = Key Then
        Set Item = mTimeframeTable(i).TheTimeFrame
        Exit For
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Remove(Key As String)
Dim i As Long

Const ProcName As String = "Remove"

On Error GoTo Err

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).Key = Key Then
        mTimeframeTable(i).TheTimeFrame.Finish
        clearTimeframeTableEntry i
        fireChange CollItemRemoved, mTimeframeTable(i).TheTimeFrame
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAll()
Dim i As Long

Const ProcName As String = "RemoveAll"

On Error GoTo Err

For i = 0 To mNextTimeframeTableIndex - 1
    mTimeframeTable(i).TheTimeFrame.Finish
    clearTimeframeTableEntry i
    fireChange CollItemRemoved, mTimeframeTable(i).TheTimeFrame
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal Value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"

On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is Value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearTimeframeTableEntry(ByVal index As Long)
Const ProcName As String = "clearTimeframeTableEntry"

On Error GoTo Err

mTimeframeTable(index).Key = ""
mTimeframeTable(index).NumberOfBarsToFetch = 0
Set mTimeframeTable(index).TheTimeFrame = Nothing
Set mTimeframeTable(index).BarTimePeriod = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Timeframe)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEventData
Const ProcName As String = "fireChange"

On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateKey( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal NumberOfBarsToFetch As Long, _
                ByVal FromTime As Date, _
                ByVal ToTime As Date, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal ExcludeCurrentbar As Boolean, _
                ByVal SaveIntervalSeconds As Long, _
                ByVal WriteBidAndAskBars As Boolean) As String
Const ProcName As String = "generateKey"

On Error GoTo Err

generateKey = BarTimePeriod.ToString & _
            "(" & NumberOfBarsToFetch & _
            IIf(FromTime <> 0, "," & Format(FromTime, "yyyymmddhhnnss"), "") & _
            IIf(ToTime <> 0, "," & Format(ToTime, "yyyymmddhhnnss"), "") & _
            ")" & _
            IIf(IncludeBarsOutsideSession, "I", "") & _
            IIf(ExcludeCurrentbar, "X", "") & _
            IIf(SaveIntervalSeconds > 0, "S" & SaveIntervalSeconds, "") & _
            IIf(WriteBidAndAskBars, "W", "")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
            
End Function

Private Function getTimeframeIndex( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal NumberOfBarsToFetch As Long, _
                ByVal FromTime As Date, _
                ByVal ToTime As Date, _
                ByRef IncludeBarsOutsideSession As Boolean, _
                ByVal SaveIntervalSeconds As Long, _
                ByVal WriteBidAndAskBars As Boolean) As Long
Dim i As Long

Const ProcName As String = "getTimeframeIndex"

On Error GoTo Err

getTimeframeIndex = -1
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).BarTimePeriod Is BarTimePeriod And _
        mTimeframeTable(i).IncludeBarsOutsideSession = IncludeBarsOutsideSession And _
        mTimeframeTable(i).NumberOfBarsToFetch = NumberOfBarsToFetch And _
        mTimeframeTable(i).FromTime = FromTime And _
        mTimeframeTable(i).ToTime = ToTime And _
        mTimeframeTable(i).SaveIntervalSeconds = SaveIntervalSeconds And _
        mTimeframeTable(i).WriteBidAndAskBars = WriteBidAndAskBars _
    Then
        getTimeframeIndex = i
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

