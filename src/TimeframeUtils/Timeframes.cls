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

Private Const ModuleName                        As String = "Timeframes"

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
    CustomSessionStartTime  As Date
    CustomSessionEndTime    As Date
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mTimeframeTable()                       As TimeframeTableEntry
Private mNextTimeframeTableIndex                As Long

Private mStudyBase                              As IStudyBase
Private mStudyInputHandler                      As StudyInputHandler
Private mContractFuture                         As IFuture
Private mHistDataStore                          As IHistoricalDataStore

Private mCollectionChangeListeners              As Listeners

Private mClockFuture                            As IFuture

Private mBarType                                As BarTypes

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTimeframeTable(1) As TimeframeTableEntry
Set mCollectionChangeListeners = New Listeners
End Sub

Private Sub Class_Terminate()
Debug.Print "Timeframes terminated"
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Public Property Get Count() As Long
Count = mNextTimeframeTableIndex
End Property

Public Property Get HistDataFactory() As IHistoricalDataStore
Set HistDataFactory = mHistDataStore
End Property

Public Property Get StudyBase() As IStudyBase
Set StudyBase = mStudyBase
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal BarTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal NumberOfBarsToFetch As Long, _
                Optional ByVal FromTime As Date, _
                Optional ByVal StateChangeListener As IStateChangeListener, _
                Optional ByVal IncludeBarsOutsideSession As Boolean, _
                Optional ByVal ExcludeCurrentbar As Boolean, _
                Optional ByVal CustomSessionStartTime As Date, _
                Optional ByVal CustomSessionEndTime As Date) As Timeframe
Const ProcName As String = "Add"
On Error GoTo Err

Select Case BarTimePeriod.Units
Case TimePeriodNone, _
        TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume, _
        TimePeriodTickMovement, _
        TimePeriodTickVolume
Case Else
        AssertArgument False, "Time period units not supported"
End Select

Set Add = addTimeframe(BarTimePeriod, _
                    Key, _
                    NumberOfBarsToFetch, _
                    FromTime, _
                    MaxDate, _
                    StateChangeListener, _
                    IncludeBarsOutsideSession, _
                    ExcludeCurrentbar, _
                    CustomSessionStartTime, _
                    CustomSessionEndTime)

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Add pListener

'gNotifyExistingCollectionMembers mTimeframes, value, Me

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
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
                Optional ByVal StateChangeListener As IStateChangeListener, _
                Optional ByVal IncludeBarsOutsideSession As Boolean, _
                Optional ByVal CustomSessionStartTime As Date, _
                Optional ByVal CustomSessionEndTime As Date) As Timeframe
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
        AssertArgument False, "Time period units not supported"
    
End Select

AssertArgument FromTime <= ToTime, "fromTime must be earlier than toTime"

Set AddHistorical = addTimeframe(BarTimePeriod, _
                    Key, _
                    NumberOfBarsToFetch, _
                    FromTime, _
                    ToTime, _
                    StateChangeListener, _
                    IncludeBarsOutsideSession, _
                    False, _
                    CustomSessionStartTime, _
                    CustomSessionEndTime)

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName

End Function

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

RemoveAll

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStudyBase As IStudyBase, _
                ByVal pContractFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pClockFuture As IFuture, _
                ByVal pBarType As BarTypes)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pStudyBase Is Nothing, "pStudyBase is Nothing"
If Not pContractFuture Is Nothing Then AssertArgument Not (pContractFuture.IsCancelled Or pContractFuture.IsFaulted), "pContractFuture is cancelled or faulted"
If Not pClockFuture Is Nothing Then AssertArgument Not (pClockFuture.IsCancelled Or pClockFuture.IsFaulted), "pClockFuture is cancelled or faulted"

Set mStudyBase = pStudyBase
Set mContractFuture = pContractFuture

Set mHistDataStore = pHistDataStore
Set mClockFuture = pClockFuture

mBarType = pBarType

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item(Key As String) As Timeframe
Const ProcName As String = "Item"
On Error GoTo Err

Dim i As Long
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).Key = Key Then
        Set Item = mTimeframeTable(i).TheTimeFrame
        Exit For
    End If
Next

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Remove(Key As String)
Const ProcName As String = "Remove"
On Error GoTo Err

Dim i As Long
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).Key = Key Then
        mTimeframeTable(i).TheTimeFrame.Finish
        clearTimeframeTableEntry i
        fireCollectionChange CollItemRemoved, mTimeframeTable(i).TheTimeFrame
        Exit For
    End If
Next

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAll()
Const ProcName As String = "RemoveAll"
On Error GoTo Err

Dim i As Long
For i = 0 To mNextTimeframeTableIndex - 1
    mTimeframeTable(i).TheTimeFrame.Finish
    clearTimeframeTableEntry i
    fireCollectionChange CollItemRemoved, mTimeframeTable(i).TheTimeFrame
Next

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Remove pListener

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addTimeframe( _
                ByVal BarTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal NumberOfBarsToFetch As Long, _
                ByVal FromTime As Date, _
                ByVal ToTime As Date, _
                ByVal StateChangeListener As IStateChangeListener, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal ExcludeCurrentbar As Boolean, _
                ByVal CustomSessionStartTime As Date, _
                ByVal CustomSessionEndTime As Date) As Timeframe
Const ProcName As String = "addTimeframe"
On Error GoTo Err

Dim timeframeIndex As Long
timeframeIndex = getTimeframeIndex(BarTimePeriod, _
                                NumberOfBarsToFetch, _
                                FromTime, _
                                ToTime, _
                                IncludeBarsOutsideSession, _
                                CustomSessionStartTime, _
                                CustomSessionEndTime)
If timeframeIndex >= 0 Then
    Set addTimeframe = mTimeframeTable(timeframeIndex).TheTimeFrame
    Exit Function
End If

Dim tte As TimeframeTableEntry
Set addTimeframe = New Timeframe
Set tte.TheTimeFrame = addTimeframe
Set tte.BarTimePeriod = BarTimePeriod
tte.NumberOfBarsToFetch = NumberOfBarsToFetch
tte.FromTime = FromTime
tte.ToTime = ToTime
tte.IncludeBarsOutsideSession = IncludeBarsOutsideSession
tte.CustomSessionStartTime = CustomSessionStartTime
tte.CustomSessionEndTime = CustomSessionEndTime

If Key <> "" Then
    tte.Key = Key
Else
    tte.Key = generateKey(BarTimePeriod, _
                        NumberOfBarsToFetch, _
                        FromTime, _
                        ToTime, _
                        IncludeBarsOutsideSession, _
                        ExcludeCurrentbar, _
                        CustomSessionStartTime, _
                        CustomSessionEndTime)
End If
tte.TheTimeFrame.Initialise mStudyBase, _
                            mContractFuture, _
                            mHistDataStore, _
                            tte.Key, _
                            BarTimePeriod, _
                            NumberOfBarsToFetch, _
                            mBarType, _
                            FromTime, _
                            ToTime, _
                            StateChangeListener, _
                            IncludeBarsOutsideSession, _
                            ExcludeCurrentbar, _
                            mClockFuture, _
                            CustomSessionStartTime, _
                            CustomSessionEndTime

If mNextTimeframeTableIndex > UBound(mTimeframeTable) Then
    ReDim Preserve mTimeframeTable(2 * (UBound(mTimeframeTable) + 1) - 1) As TimeframeTableEntry
End If
mTimeframeTable(mNextTimeframeTableIndex) = tte
mNextTimeframeTableIndex = mNextTimeframeTableIndex + 1
fireCollectionChange CollItemAdded, tte.TheTimeFrame

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub clearTimeframeTableEntry(ByVal index As Long)
Const ProcName As String = "clearTimeframeTableEntry"
On Error GoTo Err

mTimeframeTable(index).Key = ""
mTimeframeTable(index).NumberOfBarsToFetch = 0
Set mTimeframeTable(index).TheTimeFrame = Nothing
Set mTimeframeTable(index).BarTimePeriod = Nothing

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireCollectionChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Timeframe)
Const ProcName As String = "fireCollectionChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mCollectionChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mCollectionChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ICollectionChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateKey( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal NumberOfBarsToFetch As Long, _
                ByVal FromTime As Date, _
                ByVal ToTime As Date, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal ExcludeCurrentbar As Boolean, _
                ByVal CustomSessionStartTime As Date, _
                ByVal CustomSessionEndTime As Date) As String
Const ProcName As String = "generateKey"
On Error GoTo Err

generateKey = BarTimePeriod.ToString & _
            "(" & NumberOfBarsToFetch & _
            IIf(FromTime <> 0, "," & Format(FromTime, "yyyymmddhhnnss"), "") & _
            IIf(ToTime <> 0, "," & Format(ToTime, "yyyymmddhhnnss"), "") & _
            ")" & _
            IIf(CustomSessionStartTime <> 0 Or CustomSessionEndTime <> 0, _
                "[" & Format(CustomSessionStartTime, "hh:mm") & _
                "-" & Format(CustomSessionEndTime, "hh:mm") & "]", _
                "") & _
            IIf(IncludeBarsOutsideSession, "I", "") & _
            IIf(ExcludeCurrentbar, "X", "")

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTimeframeIndex( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal NumberOfBarsToFetch As Long, _
                ByVal FromTime As Date, _
                ByVal ToTime As Date, _
                ByVal IncludeBarsOutsideSession As Boolean, _
                ByVal CustomSessionStartTime As Date, _
                ByVal CustomSessionEndTime As Date) As Long
Const ProcName As String = "getTimeframeIndex"
On Error GoTo Err

getTimeframeIndex = -1

Dim i As Long
For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).BarTimePeriod Is BarTimePeriod And _
        mTimeframeTable(i).IncludeBarsOutsideSession = IncludeBarsOutsideSession And _
        mTimeframeTable(i).NumberOfBarsToFetch = NumberOfBarsToFetch And _
        mTimeframeTable(i).FromTime = FromTime And _
        mTimeframeTable(i).ToTime = ToTime And _
        mTimeframeTable(i).CustomSessionStartTime = CustomSessionStartTime And _
        mTimeframeTable(i).CustomSessionEndTime = CustomSessionEndTime _
    Then
        getTimeframeIndex = i
        Exit Function
    End If
Next

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

