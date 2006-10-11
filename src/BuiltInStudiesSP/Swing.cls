VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Swing"
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
' Interfaces
'================================================================================

Implements TradeBuildSP.IStudy
Implements TradeBuildSP.IStudyValueResolver
Implements TradeBuildSP.ITaskCompletionListener

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

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mIdent As String

Private mInputValueName As String
Private mUnderlyingStudy As TradeBuildSP.IStudy

Private mParams As IParameters
Private mStudies As IStudies
Private mListeners As IStudyValueListeners

Private mBarnumber As Long
Private mNewBar As Boolean

Private mHigh As Double
Private mPrevHigh As Double
Private mPrevPrevHigh As Double
Private mLow As Double
Private mPrevLow As Double
Private mPrevPrevLow As Double

Private mSwingCache As IStudyValueCache

Private mSwingHigh As SwingPoint
Private mSwingHighCache As IStudyValueCache

Private mSwingLow As SwingPoint
Private mSwingLowCache As IStudyValueCache

Private mHighBeforeLowPoint As SwingPoint
Private mHighAfterLowPoint As SwingPoint
Private mHighAfterLowPointExcl As SwingPoint     ' excluding current bar

Private mLowBeforeHighPoint As SwingPoint
Private mLowAfterHighPoint As SwingPoint
Private mLowAfterHighPointExcl As SwingPoint   ' excluding current bar

Private mMinimumSwing As Double
Private mIncludeImplicitSwingPoints As Boolean
Private mTrend As TrendDirections

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "Swing study terminated"
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByVal ValueName As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion


Select Case UCase$(ValueName)
Case UCase$(SwingValueSwingPoint)
    Set IStudy_addStudy = mSwingCache.replayValuesWithResolver( _
                                                study, _
                                                Me, _
                                                SwingValueSwingPoint, _
                                                Me, _
                                                numUnderlyingValuesToUse, _
                                                TaskDiscriminators.TaskAddStudy, _
                                                taskName, _
                                                taskData)
    IStudy_addStudy.addTaskCompletionListener Me
End Select

End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal ValueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Select Case UCase$(ValueName)
Case UCase$(SwingValueSwingPoint)
    Set IStudy_addStudyValueListener = mSwingCache.replayValues(listener, _
                                Me, _
                                numberOfValuesToReplay, _
                                TaskDiscriminators.TaskAddStudyValueListener, _
                                taskName, _
                                taskData)
    IStudy_addStudyValueListener.addTaskCompletionListener Me
End Select
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Property Let IStudy_defaultParameters(ByVal RHS As TradeBuildSP.IParameters)
GSwing.defaultParameters = RHS
End Property

Private Property Get IStudy_defaultParameters() As TradeBuildSP.IParameters
Set IStudy_defaultParameters = GSwing.defaultParameters
End Property

Private Function IStudy_getStudyValue( _
                ByVal ValueName As String, _
                ByVal ref As Long, _
                ByRef params() As Variant) As StudyValue
Select Case UCase$(ValueName)
Case UCase$(DefaultStudyValueName), UCase$(SwingValueSwingPoint)
    IStudy_getStudyValue = Swing(ref)
End Select

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
Set mUnderlyingStudy = underlyingStudy
processParameters parameters

Set mSwingCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SwingValueSwingPoint)
Set mSwingHighCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SwingValueSwingHighPoint)
Set mSwingLowCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SwingValueSwingLowPoint)

mSwingHigh.value = DummyHigh
mSwingLow.value = DummyLow

mHighBeforeLowPoint.value = DummyHigh
mHighAfterLowPoint.value = DummyHigh
mLowBeforeHighPoint.value = DummyLow
mLowAfterHighPoint.value = DummyLow

mTrend = TrendDirections.TrendUnknown

End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = SwingShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")"
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = SwingShortName & "(" & _
                mInputValueName & "," & _
                gParamsToString(mParams) & ")." & _
                mUnderlyingStudy.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
Dim isNewBar As Boolean
Static sPrevVal As Variant
Dim prevSwingHigh As SwingPoint
Dim prevSwingLow As SwingPoint
Dim evOut As StudyValueEvent

If IsEmpty(ev.value) Then Exit Sub

If ev.timestamp = "6/7/6 13:30" Then Stop

isNewBar = newBar(ev.barnumber)

If Not isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
    sPrevVal = ev.value
End If

If mSwingHighCache.Count >= 1 Then prevSwingHigh = mSwingHighCache.getValue(0)
If mSwingLowCache.Count >= 1 Then prevSwingLow = mSwingLowCache.getValue(0)

If isNewBar Then
    mPrevPrevLow = mPrevLow
    mPrevLow = mLow
    mLow = ev.value
    mPrevPrevHigh = mPrevHigh
    mPrevHigh = mHigh
    mHigh = ev.value

    mHighAfterLowPointExcl = mHighAfterLowPoint
    mLowAfterHighPointExcl = mLowAfterHighPoint

    If mSwingHigh.value <> DummyHigh Then
        If mBarnumber - mSwingHigh.barnumber >= 3 And _
            (mSwingHigh.value - mPrevHigh) >= mMinimumSwing _
        Then
            If Not IsEmpty(prevSwingHigh.value) Then
                If mSwingHigh.value >= prevSwingHigh.value Then
                    prevSwingHigh.swingType = SwingMinorHigh
                    mSwingHighCache.updateValue prevSwingHigh, prevSwingHigh.barnumber, prevSwingHigh.timestamp
                Else
                    If prevSwingHigh.swingType = SwingUnknownHigh Then
                        prevSwingHigh.swingType = SwingMajorHigh
                       mSwingHighCache.updateValue prevSwingHigh, prevSwingHigh.barnumber, prevSwingHigh.timestamp
                    End If
                    mSwingHigh.swingType = SwingMinorHigh
                End If
            End If

            If Not IsEmpty(prevSwingHigh.value) Then
                If prevSwingLow.timestamp <= prevSwingHigh.timestamp Then
                    createImpliedSwingLow prevSwingLow
                    prevSwingLow = mSwingLowCache.getValue(0)
                End If
            End If

            'mLowBeforeHighPoint.value = DummyLow
            mLowBeforeHighPoint = mLowAfterHighPoint

            mSwingHighCache.addValue mSwingHigh, mSwingHigh.barnumber, mSwingHigh.timestamp
            
            mSwingCache.addValue mSwingHigh, mSwingHigh.barnumber, mSwingHigh.timestamp
            Set evOut.Source = Me
            evOut.barnumber = mSwingHigh.barnumber
            evOut.timestamp = mSwingHigh.timestamp
            evOut.value = mSwingHigh.value
            evOut.ValueName = SwingValueSwingPoint
            mStudies.Notify evOut
            mListeners.Notify evOut
            
            prevSwingHigh = mSwingHigh
            mSwingHigh.value = DummyHigh
            mSwingHigh.barnumber = ev.barnumber
            mSwingHigh.timestamp = ev.timestamp
            mSwingHigh.swingType = SwingUnknownHigh
            mSwingHigh.implied = False
                
            ' we now need to discard any incompletely formed low that is earlier than
            ' the high just completed, since otherwise if it eventually completed it
            ' would be notified to listeners and studies with a time earlier than the
            ' high just notified
            If mSwingLow.value <> DummyLow And _
                mSwingLow.timestamp < prevSwingHigh.timestamp _
            Then
                mSwingLow.value = DummyLow
                mSwingLow.barnumber = ev.barnumber
                mSwingLow.timestamp = ev.timestamp
                mSwingLow.swingType = SwingUnknownLow
                mSwingLow.implied = False
            End If
    End If
    End If
    If mSwingLow.value <> DummyLow Then
        If mBarnumber - mSwingLow.barnumber >= 3 And _
            (mPrevLow - mSwingLow.value) >= mMinimumSwing _
        Then
            If Not IsEmpty(prevSwingLow.value) Then
                If mSwingLow.value <= prevSwingLow.value Then
                    prevSwingLow.swingType = SwingMinorLow
                    mSwingLowCache.updateValue prevSwingLow, prevSwingLow.barnumber, prevSwingLow.timestamp
                Else
                    If prevSwingLow.swingType = SwingUnknownLow Then
                        prevSwingLow.swingType = SwingMajorLow
                        mSwingLowCache.updateValue prevSwingLow, prevSwingLow.barnumber, prevSwingLow.timestamp
                    End If
                    mSwingLow.swingType = SwingMinorLow
                End If
            End If

            If Not IsEmpty(prevSwingLow.value) Then
                If prevSwingHigh.timestamp <= prevSwingLow.timestamp Then
                    createImpliedSwingHigh prevSwingHigh
                    prevSwingHigh = mSwingHighCache.getValue(0)
                End If
            End If

            'mHighBeforeLowPoint.value = DummyHigh
            mHighBeforeLowPoint = mHighAfterLowPoint

            mSwingLowCache.addValue mSwingLow, mSwingLow.barnumber, mSwingLow.timestamp
            
            mSwingCache.addValue mSwingLow, mSwingLow.barnumber, mSwingLow.timestamp
            Set evOut.Source = Me
            evOut.barnumber = mSwingLow.barnumber
            evOut.timestamp = mSwingLow.timestamp
            evOut.value = mSwingLow.value
            evOut.ValueName = SwingValueSwingPoint
            mStudies.Notify evOut
            mListeners.Notify evOut
            
            prevSwingLow = mSwingLow
            mSwingLow.value = DummyLow
            mSwingLow.barnumber = ev.barnumber
            mSwingLow.timestamp = ev.timestamp
            mSwingLow.swingType = SwingUnknownLow
            mSwingLow.implied = False
            
            ' we now need to discard any incompletely formed high that is earlier than
            ' the low just completed, since otherwise if it eventually completed it
            ' would be notified to listeners and studies with a time earlier than the
            ' low just notified
            If mSwingHigh.value <> DummyHigh And _
                mSwingHigh.timestamp < prevSwingLow.timestamp _
            Then
                mSwingHigh.value = DummyHigh
                mSwingHigh.barnumber = ev.barnumber
                mSwingHigh.timestamp = ev.timestamp
                mSwingHigh.swingType = SwingUnknownHigh
                mSwingHigh.implied = False
            End If
        End If
    End If

    calcTrendDirection

End If

If ev.value < mLow Then mLow = ev.value
If ev.value > mHigh Then mHigh = ev.value

If Not IsEmpty(prevSwingHigh.value) Then
    If ev.value > prevSwingHigh.value Then
        If prevSwingHigh.swingType = SwingUnknownHigh Then
            prevSwingHigh.swingType = SwingMinorHigh
            mSwingHighCache.updateValue prevSwingHigh, prevSwingHigh.barnumber, prevSwingHigh.timestamp
        End If
    End If
End If
If Not IsEmpty(prevSwingLow.value) Then
    If ev.value < prevSwingLow.value Then
        If prevSwingLow.swingType = SwingUnknownLow Then
            prevSwingLow.swingType = SwingMinorLow
            mSwingLowCache.updateValue prevSwingLow, prevSwingLow.barnumber, prevSwingLow.timestamp
        End If
    End If
End If

If (mSwingHigh.value = DummyHigh And _
    ev.value >= mPrevHigh And _
    ev.value >= mPrevPrevHigh) Or _
    (mSwingHigh.value <> DummyHigh And ev.value >= mSwingHigh.value) _
Then
    mSwingHigh.value = ev.value
    mSwingHigh.barnumber = mBarnumber
    mSwingHigh.timestamp = ev.timestamp
    
    If mLowBeforeHighPoint.value >= mLowAfterHighPoint.value Then
        mLowBeforeHighPoint = mLowAfterHighPoint
    End If
    mLowAfterHighPoint.value = ev.value
    mLowAfterHighPoint.timestamp = ev.timestamp
    mLowAfterHighPoint.barnumber = mBarnumber
End If

If (mSwingLow.value = DummyLow And _
    ev.value <= mPrevLow And _
    ev.value <= mPrevPrevLow) Or _
    (mSwingLow.value <> DummyLow And ev.value <= mSwingLow.value) _
Then
    mSwingLow.value = ev.value
    mSwingLow.barnumber = mBarnumber
    mSwingLow.timestamp = ev.timestamp
    
    If mHighBeforeLowPoint.value <= mHighAfterLowPoint.value Then
        mHighBeforeLowPoint = mHighAfterLowPoint
    End If
    mHighAfterLowPoint.value = ev.value
    mHighAfterLowPoint.timestamp = ev.timestamp
    mHighAfterLowPoint.barnumber = mBarnumber
End If

If ev.value >= mHighAfterLowPoint.value Then
    mHighAfterLowPoint.value = ev.value
    mHighAfterLowPoint.timestamp = ev.timestamp
    mHighAfterLowPoint.barnumber = mBarnumber
End If
If ev.value <= mLowAfterHighPoint.value Then
    mLowAfterHighPoint.value = ev.value
    mLowAfterHighPoint.timestamp = ev.timestamp
    mLowAfterHighPoint.barnumber = mBarnumber
End If
End Sub

Public Property Get IStudy_numberOfBarsRequired() As Long
IStudy_numberOfBarsRequired = 0
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal ValueName As String = "") As Long
IStudy_numberOfCachedValues = mSwingCache.Count
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mParams.Clone
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mListeners.Remove listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = GSwing.studyDefinition
End Property

'================================================================================
' IStudyValueResolver Interface Members
'================================================================================

Private Function IStudyValueResolver_resolveValue( _
                pStudyValue As TradeBuildSP.StudyValue, _
                ByVal requiredValueName As String) As Variant
Dim lSwingPoint As SwingPoint

lSwingPoint = pStudyValue.value

Select Case UCase$(requiredValueName)
Case UCase$(SwingValueSwingPoint)
    IStudyValueResolver_resolveValue = lSwingPoint.value
End Select
End Function

'================================================================================
' ITaskCompletionListener Interface Members
'================================================================================

Private Sub ITaskCompletionListener_taskCompleted( _
                ev As TradeBuildSP.TaskCompletionEvent)
Dim task As IValueReplayTask
Set task = ev.Source
Select Case task.discriminator
Case TaskDiscriminators.TaskAddStudy
    Dim targetStudy As IStudy
    Set targetStudy = task.Target
    mStudies.Add targetStudy, task.ValueName
Case TaskDiscriminators.TaskAddStudyValueListener
    Dim targetListener As IStudyValueListener
    Set targetListener = task.Target
    mListeners.Add targetListener, task.ValueName
End Select
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get HighAfterLowPoint(Optional ByVal excludingCurrentBar As Boolean) As SwingPoint
If excludingCurrentBar Then
    If mSwingLow.value = DummyLow Then
        HighAfterLowPoint = mHighAfterLowPointExcl
    Else
        If mHighBeforeLowPoint.value > mHighAfterLowPointExcl.value Then
            HighAfterLowPoint = mHighBeforeLowPoint
        Else
            HighAfterLowPoint = mHighAfterLowPointExcl
        End If
    End If
ElseIf mSwingLow.value = DummyLow Then
    HighAfterLowPoint = mHighAfterLowPoint
Else
    If mHighBeforeLowPoint.value > mHighAfterLowPoint.value Then
        HighAfterLowPoint = mHighBeforeLowPoint
    Else
        HighAfterLowPoint = mHighAfterLowPoint
    End If
End If
End Property

Public Property Get LowAfterHighPoint(Optional ByVal excludingCurrentBar As Boolean) As SwingPoint
If excludingCurrentBar Then
    If mSwingHigh.value = DummyHigh Then
        LowAfterHighPoint = mLowAfterHighPointExcl
    Else
        If mLowBeforeHighPoint.value < mLowAfterHighPoint.value Then
            LowAfterHighPoint = mLowBeforeHighPoint
        Else
            LowAfterHighPoint = mLowAfterHighPointExcl
        End If
    End If
ElseIf mSwingHigh.value = DummyHigh Then
    LowAfterHighPoint = mLowAfterHighPoint
Else
    If mLowBeforeHighPoint.value < mLowAfterHighPoint.value Then
        LowAfterHighPoint = mLowBeforeHighPoint
    Else
        LowAfterHighPoint = mLowAfterHighPoint
    End If
End If
End Property

Public Property Get swingHighPoint(Optional ByVal ref As Long) As StudyValue
swingHighPoint = mSwingHighCache.getStudyValue(ref)
End Property

Public Property Get swingHigh(Optional ByVal ref As Long) As StudyValue
Dim sh As SwingPoint
sh = mSwingHighCache.getValue(ref)
swingHigh.barnumber = sh.barnumber
swingHigh.timestamp = sh.timestamp
swingHigh.value = sh.value
End Property

Public Property Get swingLowPoint(Optional ByVal ref As Long) As StudyValue
swingLowPoint = mSwingLowCache.getStudyValue(ref)
End Property

Public Property Get swingLow(Optional ByVal ref As Long) As StudyValue
Dim sh As SwingPoint
sh = mSwingLowCache.getValue(ref)
swingLow.barnumber = sh.barnumber
swingLow.timestamp = sh.timestamp
swingLow.value = sh.value
End Property

Public Property Get Swing(Optional ByVal ref As Long) As StudyValue
Dim sh As SwingPoint
sh = mSwingCache.getValue(ref)
Swing.barnumber = sh.barnumber
Swing.timestamp = sh.timestamp
Swing.value = sh.value
End Property

Public Property Get trend() As TrendDirections
trend = mTrend
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                Optional ByVal numberOfValuesToCache As Long, _
                Optional ByVal params As IParameters)

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer
Set mStudies = mCommonServiceConsumer.NewStudies
Set mListeners = mCommonServiceConsumer.NewStudyValueListeners

Set mSwingCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SwingValueSwingPoint)
Set mSwingHighCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SwingValueSwingHighPoint)
Set mSwingLowCache = mCommonServiceConsumer.NewStudyValueCache(numberOfValuesToCache, SwingValueSwingLowPoint)

If Not params Is Nothing Then processParameters params

mSwingHigh.value = DummyHigh
mSwingLow.value = DummyLow
mHighBeforeLowPoint.value = DummyHigh
mHighAfterLowPoint.value = DummyHigh
mLowBeforeHighPoint.value = DummyLow
mLowAfterHighPoint.value = DummyLow

mTrend = TrendDirections.TrendUnknown

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub calcTrendDirection()
Dim latestSwingLow As Double
Dim prevSwingLow As Double
Dim latestSwingHigh As Double
Dim prevSwingHigh As Double

mTrend = TrendUnknown

If mSwingHighCache.Count < 2 Then Exit Sub
If mSwingLowCache.Count < 2 Then Exit Sub

latestSwingLow = swingLow(0).value
prevSwingLow = swingLow(-1).value
latestSwingHigh = swingHigh(0).value
prevSwingHigh = swingHigh(-1).value

If (mSwingHigh.value > latestSwingHigh Or latestSwingHigh > prevSwingHigh) And _
    mSwingLow.value >= latestSwingLow And latestSwingLow >= prevSwingLow Then
    mTrend = TrendDirections.TrendUp
ElseIf (mSwingLow.value < latestSwingLow Or latestSwingLow < prevSwingLow) And _
    mSwingHigh.value <= latestSwingHigh And latestSwingHigh < prevSwingHigh Then
    mTrend = TrendDirections.TrendDown
Else
    mTrend = TrendDirections.TrendSideways
End If
End Sub

Private Sub createImpliedSwingHigh(ByRef prevSwingHigh As SwingPoint)
Dim tempSwingHigh As SwingPoint
Dim evOut As StudyValueEvent

If Not mIncludeImplicitSwingPoints Then Exit Sub

tempSwingHigh.implied = True
tempSwingHigh.value = mHighBeforeLowPoint.value
tempSwingHigh.timestamp = mHighBeforeLowPoint.timestamp
tempSwingHigh.barnumber = mHighBeforeLowPoint.barnumber

mHighBeforeLowPoint = mHighAfterLowPoint

If Not IsEmpty(prevSwingHigh) Then
    If tempSwingHigh.value >= prevSwingHigh.value Then
        prevSwingHigh.swingType = SwingMinorHigh
    Else
        If prevSwingHigh.swingType = SwingUnknownHigh Then
            prevSwingHigh.swingType = SwingMajorHigh
        End If
        tempSwingHigh.swingType = SwingMinorHigh
    End If
End If
mSwingHighCache.addValue tempSwingHigh, tempSwingHigh.barnumber, tempSwingHigh.timestamp
mSwingCache.addValue tempSwingHigh, tempSwingHigh.barnumber, tempSwingHigh.timestamp

Set evOut.Source = Me
evOut.barnumber = tempSwingHigh.barnumber
evOut.timestamp = tempSwingHigh.timestamp
evOut.value = tempSwingHigh.value
evOut.ValueName = SwingValueSwingPoint
mStudies.Notify evOut
mListeners.Notify evOut
End Sub

Private Sub createImpliedSwingLow(ByRef prevSwingLow As SwingPoint)
Dim tempSwingLow As SwingPoint
Dim evOut As StudyValueEvent

If Not mIncludeImplicitSwingPoints Then Exit Sub

tempSwingLow.implied = True
tempSwingLow.value = mLowBeforeHighPoint.value
tempSwingLow.timestamp = mLowBeforeHighPoint.timestamp
tempSwingLow.barnumber = mLowBeforeHighPoint.barnumber

mLowBeforeHighPoint = mLowAfterHighPoint

If Not IsEmpty(prevSwingLow) Then
    If tempSwingLow.value <= prevSwingLow.value Then
        prevSwingLow.swingType = SwingMinorLow
    Else
        If prevSwingLow.swingType = SwingUnknownLow Then
            prevSwingLow.swingType = SwingMajorLow
        End If
        tempSwingLow.swingType = SwingMinorLow
    End If
End If
mSwingLowCache.addValue tempSwingLow, tempSwingLow.barnumber, tempSwingLow.timestamp
mSwingCache.addValue tempSwingLow, tempSwingLow.barnumber, tempSwingLow.timestamp

Set evOut.Source = Me
evOut.barnumber = tempSwingLow.barnumber
evOut.timestamp = tempSwingLow.timestamp
evOut.value = tempSwingLow.value
evOut.ValueName = SwingValueSwingPoint
mStudies.Notify evOut
mListeners.Notify evOut
End Sub

Private Function newBar(barnumber As Long) As Boolean
Static previousBarNumber As Long
If barnumber <> previousBarNumber Then
    mBarnumber = barnumber
    previousBarNumber = barnumber
    newBar = True
Else
    newBar = False
End If
End Function

Private Sub processParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set mParams = params.Clone

Set param = mParams.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SwingParamMinimumSwingTicks)
        mMinimumSwing = param.value * mStudyServiceConsumer.Contract.MinimumTick
    Case UCase$(SwingParamIncludeImplicitSwingPoints)
        mIncludeImplicitSwingPoints = IIf((UCase$(param.value) = "Y") Or (UCase$(param.value) = "YES"), True, False)
    End Select
    Set param = mParams.getNextParameter
Loop

End Sub


