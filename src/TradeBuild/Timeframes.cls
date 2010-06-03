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
    Key                     As String
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

Const ProcName As String = "Contract"

On Error GoTo Err

Set mContract = value
For i = 0 To mNextTimeframeTableIndex - 1
    tte = mTimeframeTable(i)
    tte.theTimeFrame.Contract = mContract
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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

Public Function Add( _
                ByVal barTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal numberOfBarsToFetch As Long, _
                Optional ByVal includeBarsOutsideSession As Boolean, _
                Optional ByVal excludeCurrentbar As Boolean) As Timeframe

Const ProcName As String = "Add"

On Error GoTo Err

Select Case barTimePeriod.Units
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

Set Add = AddTimeframe(barTimePeriod, _
                    Key, _
                    numberOfBarsToFetch, _
                    0, _
                    0, _
                    includeBarsOutsideSession, _
                    excludeCurrentbar, _
                    Nothing, _
                    0, _
                    False)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add value

'gNotifyExistingCollectionMembers mTimeframes, value, Me

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
                ByVal barTimePeriod As TimePeriod, _
                ByRef Key As String, _
                ByVal numberOfBarsToFetch As Long, _
                Optional ByVal fromTime As Date, _
                Optional ByVal toTime As Date, _
                Optional ByVal includeBarsOutsideSession As Boolean) As Timeframe

Const ProcName As String = "AddHistorical"

On Error GoTo Err

Select Case barTimePeriod.Units
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

If fromTime > toTime Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "fromTime must be earlier than toTime"
End If

Set AddHistorical = AddTimeframe(barTimePeriod, _
                    Key, _
                    numberOfBarsToFetch, _
                    fromTime, _
                    toTime, _
                    includeBarsOutsideSession, _
                    False, _
                    Nothing, _
                    0, _
                    False)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Function AddTimeframe( _
                ByVal barTimePeriod As TimePeriod, _
                ByRef Key As String, _
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

Const ProcName As String = "AddTimeframe"

On Error GoTo Err

timeframeIndex = getTimeframeIndex(barTimePeriod, _
                                numberOfBarsToFetch, _
                                fromTime, _
                                toTime, _
                                includeBarsOutsideSession, _
                                saveIntervalSeconds, _
                                writeBidAndAskBars)
If timeframeIndex >= 0 Then
    Set AddTimeframe = mTimeframeTable(timeframeIndex).theTimeFrame
    Exit Function
End If

Set AddTimeframe = New Timeframe
Set tte.theTimeFrame = AddTimeframe
Set tte.barTimePeriod = barTimePeriod
tte.numberOfBarsToFetch = numberOfBarsToFetch
tte.fromTime = fromTime
tte.toTime = toTime
tte.includeBarsOutsideSession = includeBarsOutsideSession
tte.writeBidAndAskBars = writeBidAndAskBars
tte.saveIntervalSeconds = saveIntervalSeconds

If Key <> "" Then
    tte.Key = Key
Else
    tte.Key = generateKey(barTimePeriod, _
                        numberOfBarsToFetch, _
                        fromTime, _
                        toTime, _
                        includeBarsOutsideSession, _
                        excludeCurrentbar, _
                        saveIntervalSeconds, _
                        writeBidAndAskBars)
End If
tte.theTimeFrame.Initialise mContract, _
                            tte.Key, _
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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

RemoveAll
Set mTicker = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Function Item(Key As String) As Timeframe
Dim i As Long

Const ProcName As String = "Item"

On Error GoTo Err

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).Key = Key Then
        Set Item = mTimeframeTable(i).theTimeFrame
        Exit For
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub Remove(Key As String)
Dim i As Long

Const ProcName As String = "Remove"

On Error GoTo Err

For i = 0 To mNextTimeframeTableIndex - 1
    If mTimeframeTable(i).Key = Key Then
        mTimeframeTable(i).theTimeFrame.Finish
        clearTimeframeTableEntry i
        fireChange CollItemRemoved, mTimeframeTable(i).theTimeFrame
        Exit For
    End If
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveAll()
Dim i As Long

Const ProcName As String = "RemoveAll"

On Error GoTo Err

For i = 0 To mNextTimeframeTableIndex - 1
    mTimeframeTable(i).theTimeFrame.Finish
    clearTimeframeTableEntry i
    fireChange CollItemRemoved, mTimeframeTable(i).theTimeFrame
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"

On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearTimeframeTableEntry(ByVal index As Long)
Const ProcName As String = "clearTimeframeTableEntry"

On Error GoTo Err

mTimeframeTable(index).Key = ""
mTimeframeTable(index).numberOfBarsToFetch = 0
Set mTimeframeTable(index).theTimeFrame = Nothing
Set mTimeframeTable(index).barTimePeriod = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Timeframe)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Const ProcName As String = "fireChange"

On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Const ProcName As String = "generateKey"

On Error GoTo Err

generateKey = barTimePeriod.ToString & _
            "(" & numberOfBarsToFetch & _
            IIf(fromTime <> 0, "," & Format(fromTime, "yyyymmddhhnnss"), "") & _
            IIf(toTime <> 0, "," & Format(toTime, "yyyymmddhhnnss"), "") & _
            ")" & _
            IIf(includeBarsOutsideSession, "I", "") & _
            IIf(excludeCurrentbar, "X", "") & _
            IIf(saveIntervalSeconds > 0, "S" & saveIntervalSeconds, "") & _
            IIf(writeBidAndAskBars, "W", "")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
            
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

Const ProcName As String = "getTimeframeIndex"

On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

