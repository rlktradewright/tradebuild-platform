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

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------

Private mStudyServiceConsumer As IStudyServiceConsumer
Private mCommonServiceConsumer As ICommonServiceConsumer

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
' Custom methods that must be implemented
'================================================================================

Private Sub customInitialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)

mStudyFoundation.ReplayUsingStudyValues = True

Set mCommonServiceConsumer = commonServiceConsumer
Set mStudyServiceConsumer = studyServiceConsumer

customProcessParameters parameters

Set mSwingCache = mStudyFoundation.FindValueCache(SwingValueSwingPoint)
Set mSwingHighCache = mStudyFoundation.FindValueCache(SwingValueSwingHighPoint)
Set mSwingLowCache = mStudyFoundation.FindValueCache(SwingValueSwingLowPoint)

mSwingHigh.value = DummyHigh
mSwingLow.value = DummyLow

mHighBeforeLowPoint.value = DummyHigh
mHighAfterLowPoint.value = DummyHigh
mLowBeforeHighPoint.value = DummyLow
mLowAfterHighPoint.value = DummyLow

mTrend = TrendDirections.TrendUnknown

End Sub

Private Sub customNotify( _
                ev As TradeBuildSP.StudyValueEvent)
Static sPrevVal As Variant
Dim prevSwingHigh As SwingPoint
Dim prevSwingLow As SwingPoint
Dim evOut As StudyValueEvent

If IsEmpty(ev.value) Then Exit Sub

If Not mStudyFoundation.isNewBar Then
    If ev.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.value

If mSwingHighCache.Count >= 1 Then prevSwingHigh = mSwingHighCache.getValue(0)
If mSwingLowCache.Count >= 1 Then prevSwingLow = mSwingLowCache.getValue(0)

If mStudyFoundation.isNewBar Then
    mPrevPrevLow = mPrevLow
    mPrevLow = mLow
    mLow = ev.value
    mPrevPrevHigh = mPrevHigh
    mPrevHigh = mHigh
    mHigh = ev.value

    mHighAfterLowPointExcl = mHighAfterLowPoint
    mLowAfterHighPointExcl = mLowAfterHighPoint
    
    If mStudyFoundation.localBarNumber = 1 Then
        mSwingHigh.barnumber = ev.barnumber
        mSwingHigh.timestamp = ev.timestamp
        mSwingHigh.value = Empty
        mSwingCache.addValue mSwingHigh, 1, ev.timestamp
        mSwingHighCache.addValue mSwingHigh, 1, ev.timestamp
        mSwingHigh.value = DummyHigh
        
        mSwingLow.barnumber = ev.barnumber
        mSwingLow.timestamp = ev.timestamp
        mSwingLow.value = Empty
        mSwingLowCache.addValue mSwingLow, 1, ev.timestamp
        mSwingLow.value = DummyLow
    End If
    
    If mSwingHigh.value <> DummyHigh Then
        If mStudyFoundation.barnumber - mSwingHigh.barnumber >= 3 And _
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

            mLowBeforeHighPoint = mLowAfterHighPoint

            mSwingHighCache.addValue mSwingHigh, mSwingHigh.barnumber, mSwingHigh.timestamp
            
            mSwingCache.addValue mSwingHigh, mSwingHigh.barnumber, mSwingHigh.timestamp
            Set evOut.Source = Me
            evOut.barnumber = mSwingHigh.barnumber
            evOut.timestamp = mSwingHigh.timestamp
            evOut.value = mSwingHigh.value
            evOut.valueName = SwingValueSwingPoint
            mStudyFoundation.notifyValue evOut
            evOut.valueName = SwingValueSwingHighPoint
            mStudyFoundation.notifyValue evOut
            
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
        If mStudyFoundation.barnumber - mSwingLow.barnumber >= 3 And _
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

            mHighBeforeLowPoint = mHighAfterLowPoint

            mSwingLowCache.addValue mSwingLow, mSwingLow.barnumber, mSwingLow.timestamp
            
            mSwingCache.addValue mSwingLow, mSwingLow.barnumber, mSwingLow.timestamp
            Set evOut.Source = Me
            evOut.barnumber = mSwingLow.barnumber
            evOut.timestamp = mSwingLow.timestamp
            evOut.value = mSwingLow.value
            evOut.valueName = SwingValueSwingPoint
            mStudyFoundation.notifyValue evOut
            evOut.valueName = SwingValueSwingLowPoint
            mStudyFoundation.notifyValue evOut
            
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
    mSwingHigh.barnumber = ev.barnumber
    mSwingHigh.timestamp = ev.timestamp
    
    If mLowBeforeHighPoint.value >= mLowAfterHighPoint.value Then
        mLowBeforeHighPoint = mLowAfterHighPoint
    End If
    mLowAfterHighPoint.value = ev.value
    mLowAfterHighPoint.timestamp = ev.timestamp
    mLowAfterHighPoint.barnumber = ev.barnumber
End If

If (mSwingLow.value = DummyLow And _
    ev.value <= mPrevLow And _
    ev.value <= mPrevPrevLow) Or _
    (mSwingLow.value <> DummyLow And ev.value <= mSwingLow.value) _
Then
    mSwingLow.value = ev.value
    mSwingLow.barnumber = ev.barnumber
    mSwingLow.timestamp = ev.timestamp
    
    If mHighBeforeLowPoint.value <= mHighAfterLowPoint.value Then
        mHighBeforeLowPoint = mHighAfterLowPoint
    End If
    mHighAfterLowPoint.value = ev.value
    mHighAfterLowPoint.timestamp = ev.timestamp
    mHighAfterLowPoint.barnumber = ev.barnumber
End If

If ev.value >= mHighAfterLowPoint.value Then
    mHighAfterLowPoint.value = ev.value
    mHighAfterLowPoint.timestamp = ev.timestamp
    mHighAfterLowPoint.barnumber = ev.barnumber
End If
If ev.value <= mLowAfterHighPoint.value Then
    mLowAfterHighPoint.value = ev.value
    mLowAfterHighPoint.timestamp = ev.timestamp
    mLowAfterHighPoint.barnumber = ev.barnumber
End If
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = 0
End Function

Private Sub customProcessParameters(ByVal params As IParameters)
Dim param As TradeBuildSP.IParameter

Set param = params.getFirstParameter

Do While Not param Is Nothing
    Select Case UCase$(param.name)
    Case UCase$(SwingParamMinimumSwingTicks)
        mMinimumSwing = param.value * mStudyServiceConsumer.Contract.MinimumTick
    Case UCase$(SwingParamIncludeImplicitSwingPoints)
        mIncludeImplicitSwingPoints = IIf((UCase$(param.value) = "Y") Or (UCase$(param.value) = "YES"), True, False)
    End Select
    Set param = params.getNextParameter
Loop

End Sub

Private Function customStudyDefinition() As IStudyDefinition
Set customStudyDefinition = GSwing.studyDefinition
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get HighAfterLow( _
                Optional ByVal excludingCurrentBar As Boolean) As StudyValue
Dim thePoint As SwingPoint
thePoint = HighAfterLowPoint(excludingCurrentBar).value
HighAfterLow.barnumber = thePoint.barnumber
HighAfterLow.timestamp = thePoint.timestamp
HighAfterLow.value = thePoint.value
End Property

Public Property Get HighAfterLowPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As StudyValue
Dim thePoint As SwingPoint
If excludingCurrentBar Then
    If mSwingLow.value = DummyLow Then
        thePoint = mHighAfterLowPointExcl
    Else
        If mHighBeforeLowPoint.value > mHighAfterLowPointExcl.value Then
           thePoint = mHighBeforeLowPoint
        Else
            thePoint = mHighAfterLowPointExcl
        End If
    End If
ElseIf mSwingLow.value = DummyLow Then
    thePoint = mHighAfterLowPoint
Else
    If mHighBeforeLowPoint.value > mHighAfterLowPoint.value Then
        thePoint = mHighBeforeLowPoint
    Else
        thePoint = mHighAfterLowPoint
    End If
End If
HighAfterLowPoint.barnumber = thePoint.barnumber
HighAfterLowPoint.timestamp = thePoint.timestamp
HighAfterLowPoint.value = thePoint
End Property

Public Property Get LowAfterHigh( _
                Optional ByVal excludingCurrentBar As Boolean) As StudyValue
Dim thePoint As SwingPoint
thePoint = LowAfterHighPoint(excludingCurrentBar).value
LowAfterHigh.barnumber = thePoint.barnumber
LowAfterHigh.timestamp = thePoint.timestamp
LowAfterHigh.value = thePoint.value
End Property

Public Property Get LowAfterHighPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As StudyValue
Dim thePoint As SwingPoint
If excludingCurrentBar Then
    If mSwingHigh.value = DummyHigh Then
        thePoint = mLowAfterHighPointExcl
    Else
        If mLowBeforeHighPoint.value < mLowAfterHighPoint.value Then
            thePoint = mLowBeforeHighPoint
        Else
            thePoint = mLowAfterHighPointExcl
        End If
    End If
ElseIf mSwingHigh.value = DummyHigh Then
    thePoint = mLowAfterHighPoint
Else
    If mLowBeforeHighPoint.value < mLowAfterHighPoint.value Then
        thePoint = mLowBeforeHighPoint
    Else
        thePoint = mLowAfterHighPoint
    End If
End If
LowAfterHighPoint.barnumber = thePoint.barnumber
LowAfterHighPoint.timestamp = thePoint.timestamp
LowAfterHighPoint.value = thePoint
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

Public Property Get SwingPoint(Optional ByVal ref As Long) As StudyValue
Dim sh As SwingPoint
sh = mSwingCache.getValue(ref)
Swing.barnumber = sh.barnumber
Swing.timestamp = sh.timestamp
Swing.value = sh
End Property

Public Property Get trend() As TrendDirections
trend = mTrend
End Property

'================================================================================
' Methods
'================================================================================

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
evOut.valueName = SwingValueSwingPoint
mStudyFoundation.notifyValue evOut
evOut.valueName = SwingValueSwingHighPoint
mStudyFoundation.notifyValue evOut
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
evOut.valueName = SwingValueSwingPoint
mStudyFoundation.notifyValue evOut
evOut.valueName = SwingValueSwingLowPoint
mStudyFoundation.notifyValue evOut
End Sub

Private Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
Select Case UCase$(valueName)
Case UCase$(SwingValueSwingPoint)
    getStudyValue = Swing(ref)
Case UCase$(SwingValueSwingHighPoint)
    getStudyValue = swingHigh(ref)
Case UCase$(SwingValueSwingLowPoint)
    getStudyValue = swingLow(ref)
End Select
End Function


'================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'================================================================================


'================================================================================
' IStudy Interface Members
'================================================================================

Private Function IStudy_addStudy( _
                ByVal study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion

Set IStudy_addStudy = mStudyFoundation.addStudy( _
                            study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function IStudy_addStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TradeBuildSP.ITaskCompletion
Set IStudy_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get IStudy_baseStudy() As IStudy
Set IStudy_baseStudy = Me
End Property

Private Function IStudy_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As StudyValue
'IStudy_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
IStudy_getStudyValue = getStudyValue(valueName, ref)
End Function

Private Property Get IStudy_id() As String
IStudy_id = mStudyFoundation.id
End Property

Private Sub IStudy_initialise( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal studyServiceConsumer As TradeBuildSP.IStudyServiceConsumer, _
                ByVal id As String, _
                ByVal parameters As TradeBuildSP.IParameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As TradeBuildSP.IStudy)
Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy, _
                            commonServiceConsumer, _
                            studyServiceConsumer
                            

customInitialise commonServiceConsumer, _
                studyServiceConsumer, _
                id, _
                parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy
End Sub

Private Property Get IStudy_instanceName() As String
IStudy_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get IStudy_instancePath() As String
IStudy_instancePath = mStudyFoundation.instancePath
End Property

Private Sub IStudy_Notify( _
                ev As TradeBuildSP.StudyValueEvent)
mStudyFoundation.NotifyInput ev
customNotify ev
End Sub

Private Property Get IStudy_numberOfBarsRequired() As Long
IStudy_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function IStudy_numberOfCachedValues( _
                Optional ByVal valueName As String = "") As Long
IStudy_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get IStudy_parameters() As IParameters
Set IStudy_parameters = mStudyFoundation.parameters
End Property

Private Sub IStudy_removeStudyValueListener( _
                ByVal listener As TradeBuildSP.IStudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Property Get IStudy_studyDefinition() As TradeBuildSP.IStudyDefinition
Set IStudy_studyDefinition = customStudyDefinition
End Property



