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

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mBarsCache                  As StudyValueCache

Private mCurrentBar                 As Bar
Private mCurrentBarNumber           As Long

Private mSession                    As session

Private mBarLength                  As Long
Private mTimePeriodUnits            As TimePeriodUnits

Private mTickVolume                 As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "Bars terminated"
End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Bar(Optional ByVal ref As Long) As Bar
Set Bar = mBarsCache.getValue(ref)
End Property

Public Property Get closeValue(Optional ByVal ref As Long) As Variant
closeValue = Bar(ref).closeValue
End Property

Public Property Get count() As Long
count = mBarsCache.count
End Property

Public Property Get currentBarNumber() As Long
currentBarNumber = mCurrentBarNumber
End Property

Public Property Get down(Optional ByVal ref As Long) As Boolean
down = Bar(ref).IsDownBar
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

Public Property Let session( _
                ByVal value As session)
Set mSession = value
End Property
                
Public Property Get tickVolume(Optional ByVal ref As Long) As Long
tickVolume = Bar(ref).tickVolume
End Property

Public Property Get up(Optional ByVal ref As Long) As Boolean
up = Bar(ref).IsUpBar
End Property

'================================================================================
' Methods
'================================================================================

''
' <p>Increments the tick volume for the current bar.</p>
'
' <p>This should only be used to set the tickvolume for historical bars, where
' the ticks cannot be counted in the normal way.</p>
'
' <p>Successive calls to this procedure during a sinlge bar cause the tick volume
' to be accumulated.</p>
'
' @param tickVol The amount by which the tick volume is to be incremented.
'
'@/
Public Sub incrementTickVolume( _
                ByVal tickVol As Long)
mTickVolume = mTickVolume + tickVol
End Sub

Friend Sub initialise( _
                ByVal barLength As Long, _
                ByVal timeUnits As TimePeriodUnits, _
                ByVal pSession As session, _
                Optional ByVal numberOfBarsToCache As Long)

mBarLength = barLength
mTimePeriodUnits = timeUnits
Set mSession = pSession
Set mBarsCache = New StudyValueCache
mBarsCache.initialise numberOfBarsToCache, "Bar"

End Sub

Public Sub notify( _
                ev As StudyValueEvent)
Dim isNewBar As Boolean
Dim barTime As Date

If mBarLength = 0 Then
    ' the caller will supply bar numbers
    isNewBar = newBar(ev.barNumber)
Else
    barTime = barStartTime(ev.timestamp, _
                            mBarLength, _
                            mTimePeriodUnits, _
                            mSession.sessionStartTime)
    If mCurrentBar Is Nothing Then
        isNewBar = True
    Else
        isNewBar = (barTime <> mCurrentBar.datetime)
    End If
End If
    
If isNewBar Then
    If mTickVolume <> 0 Then
        mCurrentBar.tickVolume = mTickVolume
        mTickVolume = 0
    End If
    
    Set mCurrentBar = New Bar
    If mBarLength = 0 Then
        mCurrentBar.datetime = ev.barStartTime
    Else
        mCurrentBar.datetime = barTime
        mCurrentBarNumber = mCurrentBarNumber + 1
    End If
    mCurrentBar.barNumber = mCurrentBarNumber
End If

mCurrentBar.Tick ev.value
If isNewBar Then
    mBarsCache.addValue mCurrentBar, _
                    IIf(ev.barNumber <> 0, ev.barNumber, mCurrentBarNumber), _
                    ev.timestamp, _
                    mCurrentBar.datetime
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function newBar(barNumber As Long) As Boolean
Static previousBarNumber As Long
If barNumber <> previousBarNumber Then
    previousBarNumber = barNumber
    mCurrentBarNumber = barNumber
    newBar = True
Else
    newBar = False
End If
End Function







