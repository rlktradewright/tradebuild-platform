VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IEnumerable

'@================================================================================
' Events
'@================================================================================

''
' Raised when a Bar is Added to or removed from the
' collection, and when the collection is cleared.
'
' @remarks
'   The following members of the <code>CollectionChangeTypes</code> enum may
'   occur in the <code>ChangeType</code> field of the
'   <code>CollectionChangedEvent</code> parameter:
' <blockquote>
'    <tt>CollItemAdded</tt><br />
' </blockquote>
' @param ev
'   contains details of the change.
'
'@/
Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "Bars"

'@================================================================================
' Member variables
'@================================================================================

Private mBarsCache                  As ValueCache

Private mCurrentBar                 As Bar
Private mCurrentBarNumber           As Long

Private mBarTimePeriod              As TimePeriod

Private mChangeListeners            As Listeners

Private mUpdateSequence             As Long

Private mSession                    As Session
Private mTickSize                   As Double

Private mNormaliseDailyTimestamps           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mChangeListeners = New Listeners
End Sub

Private Sub Class_Terminate()
Debug.Print "Bars terminated"
End Sub

'@================================================================================
' IEnumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Set IEnumerable_Enumerator = CreateEnumerator(Me)
End Property

Private Function IEnumerable_nextItem( _
                item As Variant, _
                cookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_nextItem"
On Error GoTo Err

If IsEmpty(cookie) Then
    cookie = 1
End If

If cookie > mBarsCache.Count Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Set item = mBarsCache.GetValue(CLng(cookie))
cookie = cookie + 1
IEnumerable_nextItem = True

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_remove(cookie As Variant)
Const ProcName As String = "IEnumerable_remove"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Remove not sUpported"

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function IEnumerable_skip( _
                ByVal numToSkip As Long, _
                cookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_skip"
On Error GoTo Err

If IsEmpty(cookie) Then
    cookie = 1
End If

If cookie = &H7FFFFFFF Then Exit Function

cookie = cookie + numToSkip

If cookie > mBarsCache.Count Then
    cookie = &H7FFFFFFF
    Exit Function
End If

IEnumerable_skip = True

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IEnumerable_UpdateSequence() As Long
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bar(Optional ByVal ref As Long) As Bar
Const ProcName As String = "Bar"
On Error GoTo Err

Set Bar = mBarsCache.GetValue(ref)

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BarTimePeriod() As TimePeriod
Set BarTimePeriod = mBarTimePeriod
End Property

Friend Property Get BarsCache() As ValueCache
Set BarsCache = mBarsCache
End Property

Public Property Get CloseValue(Optional ByVal ref As Long) As Double
Const ProcName As String = "CloseValue"
On Error GoTo Err

CloseValue = Bar(ref).CloseValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentBarNumber() As Long
CurrentBarNumber = mCurrentBarNumber
End Property

Public Property Get Down(Optional ByVal ref As Long) As Boolean
Const ProcName As String = "Down"
On Error GoTo Err

Down = Bar(ref).IsDownBar

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Set Enumerator = CreateEnumerator(Me)
End Property

Public Property Get Gain(Optional ByVal ref As Long) As Double
Const ProcName As String = "Gain"
On Error GoTo Err

Gain = Bar(ref).Gain

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HigherHighValue( _
                ByVal rank As Long, _
                Optional ByVal ref As Long) As Double
Const ProcName As String = "HigherHighValue"
On Error GoTo Err

Dim valueToCompare As Double
valueToCompare = HighValue(ref)

Dim currentHigh As Double
Dim numHighsFound As Long

Dim i As Long
i = 1
Do
    currentHigh = HighValue(ref - i)
    i = i + 1
    If currentHigh > valueToCompare Then
        numHighsFound = numHighsFound + 1
        valueToCompare = currentHigh
    End If
Loop While numHighsFound < rank
HigherHighValue = currentHigh

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighestCloseValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Const ProcName As String = "HighestCloseValue"
On Error GoTo Err

Static sLastPeriods      As Long
Static sLastBarNumber    As Long
Static sLastValue        As Double
Static sLastRef          As Long

If mCurrentBarNumber = sLastBarNumber And _
    periods = sLastPeriods And _
    ref = sLastRef _
Then
    If CloseValue <= sLastValue Then
        HighestCloseValue = sLastValue
        Exit Property
    Else
        HighestCloseValue = CloseValue
        sLastValue = HighestCloseValue
        Exit Property
    End If
End If

Dim i As Long
For i = 0 To periods - 1
    Dim lHigh As Double
    lHigh = CloseValue(ref - i)
    If lHigh > HighestCloseValue Then
        HighestCloseValue = lHigh
    End If
Next

sLastBarNumber = mCurrentBarNumber
sLastPeriods = periods
sLastRef = ref
sLastValue = HighestCloseValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighestHighValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Const ProcName As String = "HighestHighValue"
On Error GoTo Err

Static sLastPeriods      As Long
Static sLastBarNumber    As Long
Static sLastValue        As Double
Static sLastRef          As Long

If mCurrentBarNumber = sLastBarNumber And _
    periods = sLastPeriods And _
    ref = sLastRef _
Then
    If HighValue <= sLastValue Then
        HighestHighValue = sLastValue
        Exit Property
    Else
        HighestHighValue = HighValue
        sLastValue = HighestHighValue
        Exit Property
    End If
End If

HighestHighValue = MinDouble
Dim i As Long
For i = 0 To periods - 1
    Dim lHigh As Double
    lHigh = HighValue(ref - i)
    If lHigh > HighestHighValue Then
        HighestHighValue = lHigh
    End If
Next

sLastBarNumber = mCurrentBarNumber
sLastPeriods = periods
sLastRef = ref
sLastValue = HighestHighValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighValue(Optional ByVal ref As Long) As Double
Const ProcName As String = "HighValue"
On Error GoTo Err

HighValue = Bar(ref).HighValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Inside(Optional ByVal ref As Long) As Boolean
Const ProcName As String = "Inside"
On Error GoTo Err

Inside = (HighValue(ref) <= HighValue(ref - 1)) And (LowValue(ref) >= LowValue(ref - 1))

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Loss(Optional ByVal ref As Long) As Double
Const ProcName As String = "Loss"
On Error GoTo Err

Loss = Bar(ref).Loss

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowerLowValue( _
                ByVal rank As Long, _
                Optional ByVal ref As Long) As Double
Const ProcName As String = "LowerLowValue"
On Error GoTo Err

Dim valueToCompare As Double
valueToCompare = LowValue(ref)

Dim currentLow As Double
Dim numLowsFound As Long

Dim i As Long
i = 1
Do
    currentLow = LowValue(ref - i)
    i = i + 1
    If currentLow < valueToCompare Then
        numLowsFound = numLowsFound + 1
        valueToCompare = currentLow
    End If
Loop While numLowsFound < rank
LowerLowValue = currentLow

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowestCloseValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Const ProcName As String = "LowestCloseValue"
On Error GoTo Err

Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Double
Static lastRef          As Long

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If CloseValue >= lastValue Then
        LowestCloseValue = lastValue
        Exit Property
    Else
        LowestCloseValue = CloseValue
        lastValue = LowestCloseValue
        Exit Property
    End If
End If

Dim i As Long
For i = 0 To periods - 1
    Dim lLow As Double
    lLow = CloseValue(ref - i)
    If lLow < LowestCloseValue Then
        LowestCloseValue = lLow
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = LowestCloseValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowestLowValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Const ProcName As String = "LowestLowValue"
On Error GoTo Err

Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Double
Static lastRef          As Long

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If LowValue >= lastValue Then
        LowestLowValue = lastValue
        Exit Property
    Else
        LowestLowValue = LowValue
        lastValue = LowestLowValue
        Exit Property
    End If
End If

LowestLowValue = MaxDouble
Dim i As Long
For i = 0 To periods - 1
    Dim lLow As Double
    lLow = LowValue(ref - i)
    If lLow < LowestLowValue Then
        LowestLowValue = lLow
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = LowestLowValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowValue(Optional ByVal ref As Long) As Double
Const ProcName As String = "LowValue"
On Error GoTo Err

LowValue = Bar(ref).LowValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DailyTimestampsAreNormalised() As Boolean
DailyTimestampsAreNormalised = mNormaliseDailyTimestamps
End Property

Public Property Get OpenValue(Optional ByVal ref As Long) As Double
Const ProcName As String = "OpenValue"
On Error GoTo Err

OpenValue = Bar(ref).OpenValue

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Outside(Optional ByVal ref As Long) As Boolean
Const ProcName As String = "Outside"
On Error GoTo Err

Outside = (HighValue(ref) > HighValue(ref - 1)) And (LowValue(ref) < LowValue(ref - 1))

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Session() As Session
Set Session = mSession
End Property

Public Property Get TickSize() As Double
TickSize = mTickSize
End Property

Public Property Get TickVolume(Optional ByVal ref As Long) As Long
Const ProcName As String = "TickVolume"
On Error GoTo Err

TickVolume = Bar(ref).TickVolume

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Up(Optional ByVal ref As Long) As Boolean
Const ProcName As String = "Up"
On Error GoTo Err

Up = Bar(ref).IsUpBar

Exit Property

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                ByVal pBar As Bar, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date)
Const ProcName As String = "Add"
On Error GoTo Err

change

If mNormaliseDailyTimestamps Then
    pBar.Timestamp = gNormaliseTimestamp( _
                        pBar.Timestamp, _
                        mBarTimePeriod, _
                        mSession.SessionStartTime, _
                        mSession.SessionEndTime)
End If

mBarsCache.AddValue pBar, _
                    BarNumber, _
                    Timestamp, _
                    pBar.Timestamp
Set mCurrentBar = pBar
mCurrentBarNumber = BarNumber
mCurrentBar.BarNumber = mCurrentBarNumber
fireChange CollItemAdded, pBar

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AppendBars( _
                ByVal pBars As Bars, _
                ByVal maxNumberToAppend As Long, _
                ByVal pBarsBuilder As BarsBuilder)
Const ProcName As String = "AppendBars"
On Error GoTo Err

Dim max As Long
max = IIf(maxNumberToAppend < pBars.Count, maxNumberToAppend, pBars.Count)

Dim i As Long
For i = 1 To max
    Add pBars.Bar(i), mCurrentBarNumber + 1, pBars.Bar(i).Timestamp
    pBarsBuilder.FireBarAdded pBars.Bar(i)
Next

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function AppendBarsAsync( _
                ByVal pBars As Bars, _
                ByVal maxNumberToAppend As Long, _
                ByVal pBarsBuilder As BarsBuilder) As TaskController
Const ProcName As String = "AppendBarsAsync"
On Error GoTo Err

Dim abt As New AppendBarsTask
abt.Initialise Me, pBars, IIf(maxNumberToAppend < pBars.Count, maxNumberToAppend, pBars.Count), pBarsBuilder
Set AppendBarsAsync = StartTask(abt, PriorityNormal)

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Registers an object to act as a listener for changes to the collection.
'
' @remarks
'   The listener object must implement the <code>CollectionChangeListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddCollectionChangeListener( _
                ByVal listener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Add listener

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Returns the number of Bars in the collection.
'
' @return
'   The number of Bars in the collection.
'@/
Public Function Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mBarsCache.Count

Exit Function

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pNumberOfBarsToCache As Long, _
                ByVal pSession As Session, _
                ByVal pTickSize As Double, _
                ByVal pNormaliseDailyTimestamps As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mBarTimePeriod = pBarTimePeriod
Set mBarsCache = New ValueCache
mBarsCache.Initialise pNumberOfBarsToCache, "Bar"
Set mSession = pSession
mTickSize = pTickSize
mNormaliseDailyTimestamps = pNormaliseDailyTimestamps

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = CreateEnumerator(Me)
End Function

''
' Unregisters a listener.
'
' @param listener
'   The object which is to be unregistered.
'@/
Public Sub RemoveCollectionChangeListener( _
                ByVal listener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Remove listener

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub change()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Bar)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ICollectionChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.change ev
    Next
End If

RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
BarsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub


