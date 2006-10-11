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

Implements TradeBuildSP.IStudy

'================================================================================
' Events
'================================================================================

Event BarAdded(ByVal theBar As TradeBuild.Bar)
Event BarReplayed(ByVal theBar As TradeBuild.Bar)

Event CurrentCloseChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentHighChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentLowChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentOpenChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentTickVolumeChanged(ByVal currentBar As TradeBuild.Bar)
Event CurrentVolumeChanged(ByVal currentBar As TradeBuild.Bar)

Event HistoricBarAdded(ByVal theBar As TradeBuild.Bar)

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

Private Enum BarEventType
    CloseChanged
    HighChanged
    LowChanged
    OpenChanged
    TickVolumeChanged
    VolumeChanged
End Enum
 
'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTradeBuildAPIRef           As WeakReference

Private mIdent                      As String

Private mInputValueName             As String

Private mBars                       As Collection
Private mBarType                    As Long
Private mBarLength                  As Long

Private mName                       As String

Private WithEvents mCurrentBar      As Bar
Attribute mCurrentBar.VB_VarHelpID = -1
Private mCurrentBarNumber           As Long

Private mStudies                    As studies

Private mCurrentBarListeners        As WeakCollection
Private mChangeListeners            As WeakCollection


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mBars = New Collection

Set mStudies = New studies
mStudies.defaultValueName = "value"

Set mCurrentBarListeners = New WeakCollection
Set mChangeListeners = New WeakCollection

mBarLength = 1

End Sub

Private Sub Class_Terminate()
Debug.Print "Bars terminated"
End Sub

'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByVal valueName As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Dim lTask As AddStudyTask

If numUnderlyingValuesToUse = 0 Or mBars.Count = 0 Then
    mStudies.AddStudy study, valueName
    Exit Function
End If

Set lTask = New AddStudyTask
lTask.initialise Me, study, valueName, numUnderlyingValuesToUse
Set IStudy_addStudy = tb.taskManager.startTask(lTask, IIf(taskName = "", gGenerateGUIDString, taskName), taskData)
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

End Function

Private Property Get IStudy_baseStudy() As TradeBuildSP.IStudy

End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)

End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters

End Property

Private Function IStudy_getStudyValue( _
                ByVal valueType As String, _
                ByVal ref As Long, _
                params() As Variant) As TradeBuildSP.StudyValue

End Function

Private Property Get IStudy_id() As String
IStudy_id = mIdent
End Property

Private Sub IStudy_initialise( _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByVal inputValueName As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
mIdent = id
mInputValueName = inputValueName

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mName
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim iStudies As TradeBuildSP.iStudies

ev.barNumber = mCurrentBarNumber
Set ev.source = Me

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

Private Property Get IStudy_numberOfBarsRequired() As Long

End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
IStudy_numberOfCachedValues = mBars.Count
End Function

Private Property Get IStudy_parameters() As TradeBuildSP.IParameters

End Property

Private Sub IStudy_removeStudyValueListener( _
    ByVal listener As TradeBuildSP.IStudyValueListener)

End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = studyDefinition
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' mCurrentBar Event Handlers
'================================================================================

Private Sub mCurrentBar_CloseChanged()
fireCurrentBarEvent mCurrentBar, CloseChanged
RaiseEvent CurrentCloseChanged(mCurrentBar)
End Sub

Private Sub mCurrentBar_HighChanged()
fireCurrentBarEvent mCurrentBar, HighChanged
RaiseEvent CurrentHighChanged(mCurrentBar)
End Sub

Private Sub mCurrentBar_LowChanged()
fireCurrentBarEvent mCurrentBar, LowChanged
RaiseEvent CurrentLowChanged(mCurrentBar)
End Sub

Private Sub mCurrentBar_OpenChanged()
fireCurrentBarEvent mCurrentBar, OpenChanged
RaiseEvent CurrentOpenChanged(mCurrentBar)
End Sub

Private Sub mCurrentBar_TickVolumeChanged()
fireCurrentBarEvent mCurrentBar, TickVolumeChanged
RaiseEvent CurrentTickVolumeChanged(mCurrentBar)
End Sub

Private Sub mCurrentBar_VolumeChanged()
fireCurrentBarEvent mCurrentBar, VolumeChanged
RaiseEvent CurrentVolumeChanged(mCurrentBar)
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Bar(Optional ByVal ref As Long) As TradeBuild.Bar
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
        err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild.Bars::add", _
                    "Bar added is earlier than current bar"
    End If
End If

Set add = New Bar
add.barType = mBarType
add.barLength = mBarLength
add.datetime = timestamp

If Not mCurrentBar Is Nothing Then
    If mCurrentBar.Blank Then
        mBars.remove mBars.Count
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
Set ev.source = Me
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
        mBars.remove mBars.Count
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
Set ev.source = Me

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

Public Function Item(index As Variant) As Bar
Attribute Item.VB_UserMemId = 0
Set Item = mBars(index)
End Function

' doesn't make sense to remove a bar ?
'Public Sub Remove(index As Variant)
'mBars.Remove index
'mCurrentBarNumber = mCurrentBarNumber - 1
'End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mBars.[_NewEnum]
End Function

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

Public Sub removeCurrentBarListener(ByVal value As CurrentBarListener)
Dim i As Long
For i = mCurrentBarListeners.Count To 1 Step -1
    If mCurrentBarListeners.Item(i) Is value Then mCurrentBarListeners.remove i
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

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Bar)
Dim listener As CollectionChangeListener
Dim i As Long
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub

Private Sub fireCurrentBarEvent( _
                ByVal theBar As Bar, _
                ByVal eventType As BarEventType)
Dim listener As CurrentBarListener
Dim i As Long
Dim ev As CurrentBarEvent
Set ev.source = Me
Set ev.currentBar = theBar
For i = 1 To mCurrentBarListeners.Count
    Set listener = mCurrentBarListeners(i)
    Select Case eventType
    Case CloseChanged
        listener.CloseChanged ev
    Case HighChanged
        listener.HighChanged ev
    Case LowChanged
        listener.LowChanged ev
    Case OpenChanged
        listener.OpenChanged ev
    Case TickVolumeChanged
        listener.VolumeChanged ev
    Case VolumeChanged
        listener.TickVolumeChanged ev
    End Select
Next
End Sub
    
Private Function getTargetBar(ByVal ref As Long) As TradeBuild.Bar
If ref > mCurrentBarNumber Or _
    ref < -(mCurrentBarNumber - 1) _
Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild.Bars::getTargetBar", _
            "<ref> out of bounds"
End If

Set getTargetBar = mBars.Item(refToBarNumber(ref))
End Function

Private Function refToBarNumber(ByVal ref As Long) As Long
If ref > 0 Then
    refToBarNumber = ref
Else
    refToBarNumber = mCurrentBarNumber + ref
End If

End Function

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.target
End Function


