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
Attribute VB_Exposed = False
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

Private mBarsCache                  As IStudyValueCache

Private mCurrentBar                 As Bar
Attribute mCurrentBar.VB_VarHelpID = -1
Private mCurrentBarNumber           As Long

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
If ref = 0 Then
    Set Bar = mCurrentBar
Else
    Set Bar = getTargetBar(ref)
End If
End Property

Public Property Get barLength() As Long
barLength = mBarLength
End Property

Friend Property Let barLength(ByVal value As Long)
mBarLength = value
End Property

Public Property Get barTime(Optional ByVal ref As Long) As Variant
If ref = 0 Then
    If Not mCurrentBar Is Nothing Then barTime = mCurrentBar.datetime
Else
    barTime = getTargetBar(ref).datetime
End If
End Property

Friend Property Let barType(ByVal value As Long)
mBarType = value
End Property

Public Property Get closeValue(Optional ByVal ref As Long) As Variant
closeValue = getTargetBar(ref).closeValue
End Property

Public Property Get Count() As Long
Count = mBars.Count
End Property

Public Property Get currentBarNumber() As Long
currentBarNumber = mCurrentBarNumber
End Property

Public Property Get down(Optional ByVal ref As Long) As Boolean
down = (closeValue(ref) < closeValue(ref - 1))
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
highValue = getTargetBar(ref).highValue
End Property

Public Property Get id() As String
id = mIdent
End Property

Public Property Get inside(Optional ByVal ref As Long) As Variant
inside = (highValue(ref) <= highValue(ref - 1)) And (lowValue(ref) >= lowValue(ref - 1))
End Property

'Public Property Get isFirstBarInSession(Optional ByVal ref As long) As Variant
'isFirstBarInSession = Abs(getTargetBar(ref).datetime - mCurrentSessionStartTime) < OneSecond
'End Property
'
'Public Property Get isLastBarInSession() As Boolean
'isLastBarInSession = (Abs(barTime(0) - mLastBarStartTime) < OneSecond)
'End Property

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
lowValue = getTargetBar(ref).lowValue
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Get openValue(Optional ByVal ref As Long) As Variant
openValue = getTargetBar(ref).openValue
End Property

Public Property Get outside(Optional ByVal ref As Long) As Boolean
outside = (highValue(ref) > highValue(ref - 1)) And (lowValue(ref) < lowValue(ref - 1))
End Property

Public Property Get studyDefinition() As studyDefinition
Set studyDefinition = GBars.studyDefinition
End Sub

Friend Property Get studies() As TradeBuildSP.iStudies
Set studies = mStudies
End Property

Public Property Get tickVolume(Optional ByVal ref As Long) As Long
tickVolume = getTargetBar(ref).tickVolume
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = New WeakReference
mTradeBuildAPIRef.initialise value
End Property

Public Property Get up(Optional ByVal ref As Long) As Boolean
up = (closeValue(ref) > closeValue(ref - 1))
End Property

Public Property Get volume(Optional ByVal ref As Long) As Long
volume = getTargetBar(ref).volume
End Property

'================================================================================
' Methods
'================================================================================

Public Function add(ByVal timestamp As Date, _
                Optional ByVal key As String) As Bar
Dim ev As TradeBuild.StudyValueEvent

If Not mCurrentBar Is Nothing Then
    If timestamp <= mCurrentBar.datetime Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "Bars::add", _
                    "Bar added is earlier than current bar"
    End If
End If

Set add = New Bar
add.barType = mBarType
add.barLength = mBarLength
add.datetime = timestamp

If Not mCurrentBar Is Nothing Then
    If mCurrentBar.Blank Then
        mBars.Remove mBars.Count
    Else
        mCurrentBarNumber = mCurrentBarNumber + 1
    End If
Else
    mCurrentBarNumber = mCurrentBarNumber + 1
End If

If key <> "" Then
    mBars.add add, key
Else
    mBars.add add, Format(timestamp, "yyyymmddhhnnss")
End If

Set mCurrentBar = add

ev.barNumber = mCurrentBarNumber
Set ev.Source = Me
ev.timestamp = timestamp
' use the bar number as the value
ev.value = mCurrentBarNumber
ev.valueName = "newbar"
mStudies.notify ev
fireChange CollItemAdded, add
RaiseEvent BarAdded(add)
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
End Sub

Public Sub addCurrentBarListener(ByVal value As CurrentBarListener)
mCurrentBarListeners.add value
End Sub

Public Function addHistoric( _
                ByVal timestamp As Date, _
                ByVal openValue As Double, _
                ByVal highValue As Double, _
                ByVal lowValue As Double, _
                ByVal closeValue As Double, _
                ByVal volume As Long, _
                ByVal tickVolume As Long, _
                ByVal openInterest As Long, _
                Optional ByVal key As String) As Bar
Dim ev As TradeBuild.StudyValueEvent

Set addHistoric = New Bar
With addHistoric
    .barType = mBarType
    .barLength = mBarLength
    .datetime = timestamp
    .openValue = openValue
    .highValue = highValue
    .lowValue = lowValue
    .closeValue = closeValue
    .volume = volume
    .tickVolume = tickVolume
    .openInterest = openInterest
End With

If Not mCurrentBar Is Nothing Then
    If mCurrentBar.Blank Then
        mBars.Remove mBars.Count
    Else
        mCurrentBarNumber = mCurrentBarNumber + 1
    End If
Else
    mCurrentBarNumber = mCurrentBarNumber + 1
End If

If key <> "" Then
    mBars.add addHistoric, key
Else
    mBars.add addHistoric, Format(timestamp, "yyyymmddhhnnss")
End If
Set mCurrentBar = addHistoric

ev.barNumber = mCurrentBarNumber
Set ev.Source = Me

ev.timestamp = mCurrentBar.datetime
ev.value = mCurrentBar.openValue
ev.valueName = GBars.BarsValueClose
mStudies.notify ev

If mCurrentBar.IsDownBar Then
    ev.timestamp = mCurrentBar.datetime + (mCurrentBar.barLength / 3) / 1440
    ev.value = mCurrentBar.highValue
    ev.valueName = GBars.BarsValueClose
    mStudies.notify ev
    
    ev.timestamp = mCurrentBar.datetime + (2 * mCurrentBar.barLength / 3) / 1440
    ev.value = mCurrentBar.lowValue
    ev.valueName = GBars.BarsValueClose
    mStudies.notify ev
Else
    ev.timestamp = mCurrentBar.datetime + (mCurrentBar.barLength / 3) / 1440
    ev.value = mCurrentBar.lowValue
    ev.valueName = GBars.BarsValueClose
    mStudies.notify ev
    
    ev.timestamp = mCurrentBar.datetime + (2 * mCurrentBar.barLength / 3) / 1440
    ev.value = mCurrentBar.highValue
    ev.valueName = GBars.BarsValueClose
    mStudies.notify ev
End If

ev.timestamp = mCurrentBar.datetime + mCurrentBar.barLength / 1440 - OneSecond
ev.value = mCurrentBar.closeValue
ev.valueName = GBars.BarsValueClose
mStudies.notify ev

ev.value = mCurrentBar.volume
ev.valueName = GBars.BarsValueVolume
mStudies.notify ev

ev.value = mCurrentBar.tickVolume
ev.valueName = GBars.BarsValueTickVolume
mStudies.notify ev

fireChange CollItemAdded, addHistoric
RaiseEvent HistoricBarAdded(addHistoric)
End Function

Friend Sub finish()
Set mStudies = Nothing
End Sub

Friend Sub initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                Optional ByVal numberOfValuesToCache As Long, _
                Optional ByVal params As IParameters)

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

Set mMaCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SMAValueSMA)

If Not params Is Nothing Then processParameters params
End Sub

Public Function Item(index As Variant) As Bar
Set Item = mBars(index)
End Function

' doesn't make sense to remove a bar ?
'Public Sub Remove(index As Variant)
'mBars.Remove index
'mCurrentBarNumber = mCurrentBarNumber - 1
'End Sub

Public Function NewEnum() As IUnknown
Set NewEnum = mBars.[_NewEnum]
End Function

Private Sub notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim iStudies As TradeBuildSP.iStudies

ev.barNumber = mCurrentBarNumber
Set ev.Source = Me

Set iStudies = mStudies

Select Case ev.valueName
Case "value"
    mCurrentBar.Tick ev.value
    ev.valueName = GBars.BarsValueClose
    iStudies.notify ev
    
    ev.valueName = GBars.BarsValueTickVolume
    ev.value = mCurrentBar.tickVolume
    iStudies.notify ev
Case "size"
    mCurrentBar.TickSize ev.value
    ev.valueName = GBars.BarsValueSize
    iStudies.notify ev
Case "volume"
    mCurrentBar.volume = ev.value
    ev.valueName = GBars.BarsValueVolume
    iStudies.notify ev
End Select

End Sub

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Public Sub removeCurrentBarListener(ByVal value As CurrentBarListener)
Dim i As Long
For i = mCurrentBarListeners.Count To 1 Step -1
    If mCurrentBarListeners.Item(i) Is value Then mCurrentBarListeners.Remove i
Next
End Sub

Friend Sub replayBar( _
                ByVal index As Long)
RaiseEvent BarReplayed(mBars(index))
End Sub

'Public Function ToString() As String
'ToString = mStudyID
'End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function getTargetBar(ByVal ref As Long) As Bar
If ref > mCurrentBarNumber Or _
    ref < -(mCurrentBarNumber - 1) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "Bars::getTargetBar", _
            "<ref> out of bounds"
End If

Set getTargetBar = mBars.Item(refToBarNumber(ref))
End Function

Private Function newBar(barNumber As Long) As Boolean
Static previousBarNumber As Long
If barNumber <> previousBarNumber Then
    previousBarNumber = barNumber
    newBar = True
Else
    newBar = False
End If
End Function

Private Function refToBarNumber(ByVal ref As Long) As Long
If ref > 0 Then
    refToBarNumber = ref
Else
    refToBarNumber = mCurrentBarNumber + ref
End If

End Function




