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
' Raised when a bar is added to or removed from the
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

If cookie > mBarsCache.count Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Set item = mBarsCache.getValue(CLng(cookie))
cookie = cookie + 1
Enumerable_nextItem = True
End Function

Private Sub Enumerable_remove(cookie As Variant)
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & "Enumerable_remove", _
            "Remove not supported"


End Sub

Private Function Enumerable_skip( _
                ByVal numToSkip As Long, _
                cookie As Variant) As Boolean
If IsEmpty(cookie) Then
    cookie = 1
End If

If cookie = &H7FFFFFFF Then Exit Function

cookie = cookie + numToSkip

If cookie > mBarsCache.count Then
    cookie = &H7FFFFFFF
    Exit Function
End If

Enumerable_skip = True
End Function

Private Property Get Enumerable_updateSequence() As Long
Enumerable_updateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bar(Optional ByVal ref As Long) As Bar
Set Bar = mBarsCache.getValue(ref)
End Property

Public Property Get barTimePeriod() As TimePeriod
Set barTimePeriod = mBarTimePeriod
End Property

Friend Property Get barsCache() As ValueCache
Set barsCache = mBarsCache
End Property

Public Property Get closeValue(Optional ByVal ref As Long) As Variant
closeValue = Bar(ref).closeValue
End Property

Public Property Get currentBarNumber() As Long
currentBarNumber = mCurrentBarNumber
End Property

Public Property Get down(Optional ByVal ref As Long) As Boolean
down = Bar(ref).IsDownBar
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

Public Property Get gain(Optional ByVal ref As Long) As Double
gain = Bar(ref).gain
End Property

Public Property Get higherHighValue( _
                ByVal rank As Long, _
                Optional ByVal ref As Long) As Double
Dim valueToCompare As Double
Dim currentHigh As Double
Dim i As Long
Dim numHighsFound As Long

valueToCompare = highValue(ref)
i = 1
Do
    currentHigh = highValue(ref - i)
    i = i + 1
    If currentHigh > valueToCompare Then
        numHighsFound = numHighsFound + 1
        valueToCompare = currentHigh
    End If
Loop While numHighsFound < rank
higherHighValue = currentHigh
End Property

Public Property Get highestCloseValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Double
Dim i                   As Long
Dim high                As Double
Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Long

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If closeValue <= lastValue Then
        highestCloseValue = lastValue
        Exit Property
    Else
        highestCloseValue = closeValue
        lastValue = highestCloseValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    high = closeValue(ref - i)
    If high > highestCloseValue Then
        highestCloseValue = high
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = highestCloseValue
End Property

Public Property Get highestHighValue(ByVal periods As Long, _
                                    Optional ByVal ref As Long) As Variant
Dim i                   As Long
Dim high                As Double
Static lastPeriods      As Long
Static lastBarNumber    As Long
Static lastValue        As Variant
Static lastRef          As Long

If mCurrentBarNumber = lastBarNumber And _
    periods = lastPeriods And _
    ref = lastRef _
Then
    If highValue <= lastValue Then
        highestHighValue = lastValue
        Exit Property
    Else
        highestHighValue = highValue
        lastValue = highestHighValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    high = highValue(ref - i)
    If high > highestHighValue Then
        highestHighValue = high
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = highestHighValue
End Property

Public Property Get highValue(Optional ByVal ref As Long) As Double
highValue = Bar(ref).highValue
End Property

Public Property Get inside(Optional ByVal ref As Long) As Variant
inside = (highValue(ref) <= highValue(ref - 1)) And (lowValue(ref) >= lowValue(ref - 1))
End Property

Public Property Get loss(Optional ByVal ref As Long) As Double
loss = Bar(ref).loss
End Property

Public Property Get lowerLowValue( _
                ByVal rank As Long, _
                Optional ByVal ref As Long) As Double
Dim valueToCompare As Double
Dim currentLow As Double
Dim i As Long
Dim numLowsFound As Long

valueToCompare = lowValue(ref)
i = 1
Do
    currentLow = lowValue(ref - i)
    i = i + 1
    If currentLow < valueToCompare Then
        numLowsFound = numLowsFound + 1
        valueToCompare = currentLow
    End If
Loop While numLowsFound < rank
lowerLowValue = currentLow
End Property

Public Property Get lowestCloseValue(ByVal periods As Long, _
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
    If closeValue >= lastValue Then
        lowestCloseValue = lastValue
        Exit Property
    Else
        lowestCloseValue = closeValue
        lastValue = lowestCloseValue
        Exit Property
    End If
End If

For i = 0 To periods - 1
    Low = closeValue(ref - i)
    If Low < lowestCloseValue Then
        lowestCloseValue = Low
    End If
Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = lowestCloseValue
End Property

Public Property Get lowestLowValue(ByVal periods As Long, _
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
    If lowValue >= lastValue Then
        lowestLowValue = lastValue
        Exit Property
    Else
        lowestLowValue = lowValue
        lastValue = lowestLowValue
        Exit Property
    End If
End If

lowestLowValue = DummyLow
For i = 0 To periods - 1
    Low = lowValue(ref - i)
    If Low < lowestLowValue Then
        lowestLowValue = Low
    End If

Next

lastBarNumber = mCurrentBarNumber
lastPeriods = periods
lastRef = ref
lastValue = lowestLowValue
End Property

Public Property Get lowValue(Optional ByVal ref As Long) As Double
lowValue = Bar(ref).lowValue
End Property

Public Property Get openValue(Optional ByVal ref As Long) As Variant
openValue = Bar(ref).openValue
End Property

Public Property Get outside(Optional ByVal ref As Long) As Boolean
outside = (highValue(ref) > highValue(ref - 1)) And (lowValue(ref) < lowValue(ref - 1))
End Property

'Public Property Get session() As session
'Set session = mSession
'End Property
                
Public Property Get tickVolume(Optional ByVal ref As Long) As Long
tickVolume = Bar(ref).tickVolume
End Property

Public Property Get up(Optional ByVal ref As Long) As Boolean
up = Bar(ref).IsUpBar
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub add( _
                ByVal pBar As Bar, _
                ByVal barNumber As Long, _
                ByVal timestamp As Date)
change
mBarsCache.addValue pBar, _
                    barNumber, _
                    timestamp, _
                    pBar.timestamp
Set mCurrentBar = pBar
mCurrentBarNumber = barNumber
fireChange CollItemAdded, pBar
End Sub

Friend Sub appendBars( _
                ByVal pBars As Bars, _
                ByVal maxNumberToAppend As Long)
Dim i As Long
Dim max As Long

max = IIf(maxNumberToAppend < pBars.count, maxNumberToAppend, pBars.count)
For i = 1 To max
    add pBars.Bar(i), mCurrentBarNumber + 1, pBars.Bar(i).timestamp
Next
End Sub

Friend Function appendBarsAsync( _
                ByVal pBars As Bars, _
                ByVal maxNumberToAppend As Long) As TaskController
Dim abt As New AppendBarsTask
abt.initialise Me, pBars, IIf(maxNumberToAppend < pBars.count, maxNumberToAppend, pBars.count)
Set appendBarsAsync = StartTask(abt, PriorityNormal)
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
Public Sub addCollectionChangeListener( _
                ByVal listener As CollectionChangeListener)
mChangeListeners.add listener, CStr(ObjPtr(listener))
End Sub

''
' Returns the number of Bars in the collection.
'
' @return
'   The number of Bars in the collection.
'@/
Public Function count() As Long
count = mBarsCache.count
End Function

Friend Sub initialise( _
                ByVal barTimePeriod As TimePeriod, _
                ByVal numberOfBarsToCache As Long)
Set mBarTimePeriod = barTimePeriod
Set mBarsCache = New ValueCache
mBarsCache.initialise numberOfBarsToCache, "Bar"
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
Public Sub removeCollectionChangeListener( _
                ByVal listener As CollectionChangeListener)
mChangeListeners.Remove CStr(ObjPtr(listener))
End Sub

Friend Sub resetBarChangeIndicators()
mCurrentBar.resetChangeIndicators
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
For i = 1 To mChangeListeners.count
    Set listener = mChangeListeners(i)
    listener.change ev
Next

RaiseEvent CollectionChanged(ev)
End Sub


