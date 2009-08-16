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

Implements Enumerable

'@================================================================================
' Events
'@================================================================================

''
' Raised when a Bar is Added to or removed from the
' collection, and when the collection is cleared.
'
' @remarks
'   The folLowing members of the <code>CollectionChangeTypes</code> enum may
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
                ev As CollectionChangeEvent)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "Bars"

'@================================================================================
' Member variables
'@================================================================================

Private mBarsCache                  As ValueCache

Private mCurrentBar                 As Bar
Private mCurrentBarNumber           As Long

Private mBarTimePeriod              As TimePeriod

Private mChangeListeners            As Collection

Private mUpdateSequence             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mChangeListeners = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "Bars terminated"
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As TWUtilities30.Enumerator
Set Enumerable_Enumerator = CreateEnumerator(Me)
End Property

Private Function Enumerable_nextItem( _
                item As Variant, _
                cookie As Variant) As Boolean
If IsEmpty(cookie) Then
    cookie = 1
End If

If cookie > mBarsCache.Count Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Set item = mBarsCache.GetValue(CLng(cookie))
cookie = cookie + 1
Enumerable_nextItem = True
End Function

Private Sub Enumerable_remove(cookie As Variant)
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "Enumerable_remove", _
            "Remove not sUpported"


End Sub

Private Function Enumerable_skip( _
                ByVal numToSkip As Long, _
                cookie As Variant) As Boolean
If IsEmpty(cookie) Then
    cookie = 1
End If

If cookie = &H7FFFFFFF Then Exit Function

cookie = cookie + numToSkip

If cookie > mBarsCache.Count Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Enumerable_skip = True
End Function

Private Property Get Enumerable_UpdateSequence() As Long
Enumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bar(Optional ByVal ref As Long) As Bar
Set Bar = mBarsCache.GetValue(ref)
End Property

Public Property Get BarTimePeriod() As TimePeriod
Set BarTimePeriod = mBarTimePeriod
End Property

Friend Property Get BarsCache() As ValueCache
Set BarsCache = mBarsCache
End Property

Public Property Get CloseValue(Optional ByVal ref As Long) As Variant
CloseValue = Bar(ref).CloseValue
End Property

Public Property Get CurrentBarNumber() As Long
CurrentBarNumber = mCurrentBarNumber
End Property

Public Property Get Down(Optional ByVal ref As Long) As Boolean
Down = Bar(ref).IsDownBar
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
Gain = Bar(ref).Gain
End Property

Public Property Get HigherHighValue( _
                ByVal rank As Long, _
                Optional ByVal ref As Long) As Double
Dim valueToCompare As Double
Dim currentHigh As Double
Dim i As Long
Dim numHighsFound As Long

valueToCompare = HighValue(ref)
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
End Property

Public Property Get HighestCloseValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Dim i                   As Long
Dim High                As Double
Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Long

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If CloseValue <= lastValue Then
        HighestCloseValue = lastValue
        Exit Property
    Else
        HighestCloseValue = CloseValue
        lastValue = HighestCloseValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    High = CloseValue(ref - i)
    If High > HighestCloseValue Then
        HighestCloseValue = High
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = HighestCloseValue
End Property

Public Property Get HighestHighValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Variant
Dim i                   As Long
Dim High                As Double
Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Long

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If HighValue <= lastValue Then
        HighestHighValue = lastValue
        Exit Property
    Else
        HighestHighValue = HighValue
        lastValue = HighestHighValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    High = HighValue(ref - i)
    If High > HighestHighValue Then
        HighestHighValue = High
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = HighestHighValue
End Property

Public Property Get HighValue(Optional ByVal ref As Long) As Double
HighValue = Bar(ref).HighValue
End Property

Public Property Get Inside(Optional ByVal ref As Long) As Variant
Inside = (HighValue(ref) <= HighValue(ref - 1)) And (LowValue(ref) >= LowValue(ref - 1))
End Property

Public Property Get Loss(Optional ByVal ref As Long) As Double
Loss = Bar(ref).Loss
End Property

Public Property Get LowerLowValue( _
                ByVal rank As Long, _
                Optional ByVal ref As Long) As Double
Dim valueToCompare As Double
Dim currentLow As Double
Dim i As Long
Dim numLowsFound As Long

valueToCompare = LowValue(ref)
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
End Property

Public Property Get LowestCloseValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Dim i As Long
Dim Low As Double
Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Variant
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

For i = 0 To periods - 1
    Low = CloseValue(ref - i)
    If Low < LowestCloseValue Then
        LowestCloseValue = Low
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = LowestCloseValue
End Property

Public Property Get LowestLowValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Variant
Dim i As Long
Dim Low As Double
Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Variant
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

LowestLowValue = DummyLow
For i = 0 To periods - 1
    Low = LowValue(ref - i)
    If Low < LowestLowValue Then
        LowestLowValue = Low
    End If

Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = LowestLowValue
End Property

Public Property Get LowValue(Optional ByVal ref As Long) As Double
LowValue = Bar(ref).LowValue
End Property

Public Property Get OpenValue(Optional ByVal ref As Long) As Variant
OpenValue = Bar(ref).OpenValue
End Property

Public Property Get Outside(Optional ByVal ref As Long) As Boolean
Outside = (HighValue(ref) > HighValue(ref - 1)) And (LowValue(ref) < LowValue(ref - 1))
End Property

Public Property Get TickVolume(Optional ByVal ref As Long) As Long
TickVolume = Bar(ref).TickVolume
End Property

Public Property Get Up(Optional ByVal ref As Long) As Boolean
Up = Bar(ref).IsUpBar
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                ByVal pBar As Bar, _
                ByVal BarNumber As Long, _
                ByVal Timestamp As Date)
change
mBarsCache.AddValue pBar, _
                    BarNumber, _
                    Timestamp, _
                    pBar.Timestamp
Set mCurrentBar = pBar
mCurrentBarNumber = BarNumber
fireChange CollItemAdded, pBar
End Sub

Friend Sub AppendBars( _
                ByVal pBars As Bars, _
                ByVal maxNumberToAppend As Long)
Dim i As Long
Dim max As Long

max = IIf(maxNumberToAppend < pBars.Count, maxNumberToAppend, pBars.Count)
For i = 1 To max
    Add pBars.Bar(i), mCurrentBarNumber + 1, pBars.Bar(i).Timestamp
Next
End Sub

Friend Function AppendBarsAsync( _
                ByVal pBars As Bars, _
                ByVal maxNumberToAppend As Long) As TaskController
Dim abt As New AppendBarsTask
abt.Initialise Me, pBars, IIf(maxNumberToAppend < pBars.Count, maxNumberToAppend, pBars.Count)
Set AppendBarsAsync = StartTask(abt, PriorityNormal)
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
                ByVal listener As CollectionChangeListener)
mChangeListeners.Add listener, CStr(ObjPtr(listener))
End Sub

''
' Returns the number of Bars in the collection.
'
' @return
'   The number of Bars in the collection.
'@/
Public Function Count() As Long
Count = mBarsCache.Count
End Function

Friend Sub Initialise( _
                ByVal BarTimePeriod As TimePeriod, _
                ByVal numberOfBarsToCache As Long)
Set mBarTimePeriod = BarTimePeriod
Set mBarsCache = New ValueCache
mBarsCache.Initialise numberOfBarsToCache, "Bar"
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
                ByVal listener As CollectionChangeListener)
mChangeListeners.Remove CStr(ObjPtr(listener))
End Sub

Friend Sub ResetBarChangeIndicators()
mCurrentBar.ResetChangeIndicators
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
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent

Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.change ev
Next

RaiseEvent CollectionChanged(ev)
End Sub


