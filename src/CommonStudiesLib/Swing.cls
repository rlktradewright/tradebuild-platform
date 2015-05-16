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

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IStudy

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Swing"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

'-------this variable is required in all studies---------
Private mStudyFoundation As StudyFoundation
'--------------------------------------------------------




Private mCurrentBarHigh As Double
Private mPrevBarHigh As Double
Private mPrevPrevBarHigh As Double
Private mCurrentBarLow As Double
Private mPrevBarLow As Double
Private mPrevPrevBarLow As Double

Private mSwingCache As ValueCache
Private mSwingLineBuilder As StudyLineBuilder
Private mSwingLineCache As ValueCache

Private mPotentialSwingHighPoint As SwingPoint
Private mSwingHighCache As ValueCache

Private mSwingHighLineBuilder As StudyLineBuilder
Private mSwingHighLineCache As ValueCache

Private mPotentialSwingLowPoint As SwingPoint
Private mSwingLowCache As ValueCache

Private mSwingLowLineBuilder As StudyLineBuilder
Private mSwingLowLineCache As ValueCache

Private mHighBeforeLowPoint As SwingPoint
Private mHighAfterLowPoint As SwingPoint
Private mHighAfterLowPointExcl As SwingPoint     ' excluding current bar

Private mLowBeforeHighPoint As SwingPoint
Private mLowAfterHighPoint As SwingPoint
Private mLowAfterHighPointExcl As SwingPoint   ' excluding current bar

Private mMinimumSwing As Double
Private mIncludeImplicitSwingPoints As Boolean
Private mTrend As TrendDirections

Private mTickSize As Double

Private mSwingSequenceNumber As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Swing study terminated"
End Sub

'@================================================================================
' Custom methods that must be implemented
'@================================================================================

Private Sub customInitialise( _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)
Const ProcName As String = "customInitialise"
On Error GoTo Err

mStudyFoundation.ReplayUsingStudyValues = True

customProcessParameters Parameters

Set mSwingLineCache = mStudyFoundation.FindValueCache(SwingValueSwingLine)
Set mSwingHighLineCache = mStudyFoundation.FindValueCache(SwingValueSwingHighLine)
Set mSwingLowLineCache = mStudyFoundation.FindValueCache(SwingValueSwingLowLine)

Set mSwingCache = mStudyFoundation.FindValueCache(SwingValueSwingPoint)
Set mSwingHighCache = mStudyFoundation.FindValueCache(SwingValueSwingHighPoint)
Set mSwingLowCache = mStudyFoundation.FindValueCache(SwingValueSwingLowPoint)

mPotentialSwingHighPoint.Value = DummyHigh
mPotentialSwingLowPoint.Value = DummyLow

mHighBeforeLowPoint.Value = DummyHigh
mHighAfterLowPoint.Value = DummyHigh
mLowBeforeHighPoint.Value = DummyLow
mLowAfterHighPoint.Value = DummyLow

mTrend = TrendDirections.TrendUnknown

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub customNotify( _
                ev As StudyValueEventData)
Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.Value) Then Exit Sub

Dim currValue As Double
currValue = ev.sVal.Value
If inputValueHasNotChanged(currValue) Then Exit Sub

Dim prevSwingHigh As SwingPoint
If mSwingHighCache.Count >= 1 Then prevSwingHigh = mSwingHighCache.GetValue(0)

Dim prevSwingLow As SwingPoint
If mSwingLowCache.Count >= 1 Then prevSwingLow = mSwingLowCache.GetValue(0)

If mStudyFoundation.IsNewBar Then
    shiftBarHighsAndLows ev
    
    If mStudyFoundation.LocalBarNumber = 1 Then
        mCurrentBarHigh = currValue
        mCurrentBarLow = currValue
        
        initialiseSwingHighPoint mPotentialSwingHighPoint, ev.sVal
        mPotentialSwingHighPoint.Value = Empty
        
        initialiseSwingLowPoint mPotentialSwingLowPoint, ev.sVal
        mPotentialSwingLowPoint.Value = Empty
        
        mSwingHighCache.AddValue mPotentialSwingHighPoint, 1, ev.sVal.Timestamp, ev.sVal.BarStartTime
        mSwingHighLineCache.AddValue Nothing, 1, ev.sVal.Timestamp, ev.sVal.BarStartTime
        
        mSwingLowCache.AddValue mPotentialSwingLowPoint, 1, ev.sVal.Timestamp, ev.sVal.BarStartTime
        mSwingLowLineCache.AddValue Nothing, 1, ev.sVal.Timestamp, ev.sVal.BarStartTime
    
        mSwingCache.AddValue mPotentialSwingHighPoint, 1, ev.sVal.Timestamp, ev.sVal.BarStartTime
        mSwingLineCache.AddValue Nothing, 1, ev.sVal.Timestamp, ev.sVal.BarStartTime
        
    End If
    
    If mPotentialSwingHighPoint.Value <> DummyHigh Then
        If mStudyFoundation.BarNumber - mPotentialSwingHighPoint.BarNumber >= 3 And _
            (mPotentialSwingHighPoint.Value - mPrevBarHigh) >= mMinimumSwing _
        Then
            updateSwingHighType prevSwingHigh, mPotentialSwingHighPoint

            If Not IsEmpty(prevSwingHigh.Value) Then
                If prevSwingLow.SequenceNumber < prevSwingHigh.SequenceNumber Then
                    createImpliedSwingLow prevSwingLow
                    prevSwingLow = mSwingLowCache.GetValue(0)
                End If
            End If

            mPotentialSwingHighPoint.SequenceNumber = mSwingSequenceNumber
            mSwingSequenceNumber = mSwingSequenceNumber + 1
            
            mLowBeforeHighPoint = mLowAfterHighPoint

            mSwingHighCache.AddValue mPotentialSwingHighPoint, mPotentialSwingHighPoint.BarNumber, mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.BarStartTime
            
            If Not mSwingHighLineBuilder Is Nothing Then
                mSwingHighLineBuilder.Point2 = CreateStudyPoint(mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.Value)
                notifySwingHighLine
            End If
            If Not mSwingLineBuilder Is Nothing Then
                mSwingLineBuilder.Point2 = CreateStudyPoint(mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.Value)
                notifySwingLine
            End If
            
            Set mSwingHighLineBuilder = createStudyLineBuilder(mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.Value)
            mSwingHighLineCache.AddValue mSwingHighLineBuilder.StudyLine, mPotentialSwingHighPoint.BarNumber, mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.BarStartTime
            
            mSwingCache.AddValue mPotentialSwingHighPoint, mPotentialSwingHighPoint.BarNumber, mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.BarStartTime
            
            Set mSwingLineBuilder = createStudyLineBuilder(mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.Value)
            mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, mPotentialSwingHighPoint.BarNumber, mPotentialSwingHighPoint.Timestamp, mPotentialSwingHighPoint.BarStartTime
            
            notifySwingPoint
            notifySwingHighPoint
            
            notifySwingLine
            notifySwingHighLine
            
            prevSwingHigh = mPotentialSwingHighPoint
            initialiseSwingHighPoint mPotentialSwingHighPoint, ev.sVal
                
            ' we now need to discard any incompletely formed low that is earlier than
            ' the high just completed, since otherwise if it eventually completed it
            ' would be notified to listeners and studies with a time earlier than the
            ' high just notified
            If mPotentialSwingLowPoint.Value <> DummyLow And _
                mPotentialSwingLowPoint.Timestamp < prevSwingHigh.Timestamp _
            Then
                initialiseSwingLowPoint mPotentialSwingLowPoint, ev.sVal
            End If
        End If
    End If
    If mPotentialSwingLowPoint.Value <> DummyLow Then
        If mStudyFoundation.BarNumber - mPotentialSwingLowPoint.BarNumber >= 3 And _
            (mPrevBarLow - mPotentialSwingLowPoint.Value) >= mMinimumSwing _
        Then
            updateSwingLowType prevSwingLow, mPotentialSwingLowPoint

            If Not IsEmpty(prevSwingLow.Value) Then
                If prevSwingHigh.SequenceNumber <= prevSwingLow.SequenceNumber Then
                    createImpliedSwingHigh prevSwingHigh
                    prevSwingHigh = mSwingHighCache.GetValue(0)
                End If
            End If

            mPotentialSwingLowPoint.SequenceNumber = mSwingSequenceNumber
            mSwingSequenceNumber = mSwingSequenceNumber + 1
            
            mHighBeforeLowPoint = mHighAfterLowPoint

            mSwingLowCache.AddValue mPotentialSwingLowPoint, mPotentialSwingLowPoint.BarNumber, mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.BarStartTime
            
            If Not mSwingLowLineBuilder Is Nothing Then
                mSwingLowLineBuilder.Point2 = CreateStudyPoint(mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.Value)
                notifySwingLowLine
            End If
            If Not mSwingLineBuilder Is Nothing Then
                mSwingLineBuilder.Point2 = CreateStudyPoint(mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.Value)
                notifySwingLine
            End If
            
            Set mSwingLowLineBuilder = createStudyLineBuilder(mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.Value)
            mSwingLowLineCache.AddValue mSwingLowLineBuilder.StudyLine, mPotentialSwingLowPoint.BarNumber, mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.BarStartTime
            
            mSwingCache.AddValue mPotentialSwingLowPoint, mPotentialSwingLowPoint.BarNumber, mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.BarStartTime
            
            Set mSwingLineBuilder = createStudyLineBuilder(mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.Value)
            mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, mPotentialSwingLowPoint.BarNumber, mPotentialSwingLowPoint.Timestamp, mPotentialSwingLowPoint.BarStartTime
            
            notifySwingPoint
            notifySwingLowPoint
            
            notifySwingLine
            notifySwingLowLine
            
            prevSwingLow = mPotentialSwingLowPoint
            initialiseSwingLowPoint mPotentialSwingLowPoint, ev.sVal
            
            ' we now need to discard any incompletely formed high that is earlier than
            ' the low just completed, since otherwise if it eventually completed it
            ' would be notified to listeners and studies with a time earlier than the
            ' low just notified
            If mPotentialSwingHighPoint.Value <> DummyHigh And _
                mPotentialSwingHighPoint.Timestamp < prevSwingLow.Timestamp _
            Then
                initialiseSwingHighPoint mPotentialSwingHighPoint, ev.sVal
            End If
        End If
    End If

    calcTrendDirection

End If

updateCurrentLineEnds ev

If currValue < mCurrentBarLow Then mCurrentBarLow = currValue
If currValue > mCurrentBarHigh Then mCurrentBarHigh = currValue

If Not IsEmpty(prevSwingHigh.Value) Then
    If currValue > prevSwingHigh.Value Then
        If prevSwingHigh.SwingType = SwingUnknown Then
            prevSwingHigh.SwingType = SwingMinorHigh
            mSwingHighCache.UpdateValue prevSwingHigh, prevSwingHigh.BarNumber, prevSwingHigh.Timestamp
        End If
    End If
End If
If Not IsEmpty(prevSwingLow.Value) Then
    If currValue < prevSwingLow.Value Then
        If prevSwingLow.SwingType = SwingUnknown Then
            prevSwingLow.SwingType = SwingMinorLow
            mSwingLowCache.UpdateValue prevSwingLow, prevSwingLow.BarNumber, prevSwingLow.Timestamp
        End If
    End If
End If

If (mPotentialSwingHighPoint.Value = DummyHigh And _
    currValue >= mPrevBarHigh And _
    currValue >= mPrevPrevBarHigh) Or _
    (mPotentialSwingHighPoint.Value <> DummyHigh And currValue >= mPotentialSwingHighPoint.Value) _
Then
    setSwingPointFromSValue mPotentialSwingHighPoint, ev.sVal
    
    If mLowBeforeHighPoint.Value >= mLowAfterHighPoint.Value Then
        mLowBeforeHighPoint = mLowAfterHighPoint
    End If
    setSwingPointFromSValue mLowAfterHighPoint, ev.sVal
End If

If (mPotentialSwingLowPoint.Value = DummyLow And _
    currValue <= mPrevBarLow And _
    currValue <= mPrevPrevBarLow) Or _
    (mPotentialSwingLowPoint.Value <> DummyLow And currValue <= mPotentialSwingLowPoint.Value) _
Then
    setSwingPointFromSValue mPotentialSwingLowPoint, ev.sVal
    
    If mHighBeforeLowPoint.Value <= mHighAfterLowPoint.Value Then
        mHighBeforeLowPoint = mHighAfterLowPoint
    End If
    setSwingPointFromSValue mHighAfterLowPoint, ev.sVal
End If

If currValue >= mHighAfterLowPoint.Value Then
    setSwingPointFromSValue mHighAfterLowPoint, ev.sVal
End If
If currValue <= mLowAfterHighPoint.Value Then
    setSwingPointFromSValue mLowAfterHighPoint, ev.sVal
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(SwingParamMinimumSwingTicks)
        mMinimumSwing = param.Value * mTickSize
    Case UCase$(SwingParamIncludeImplicitSwingPoints)
        mIncludeImplicitSwingPoints = IIf((UCase$(param.Value) = "Y") Or (UCase$(param.Value) = "YES") Or (UCase$(param.Value) = "T") Or (UCase$(param.Value) = "TRUE") Or (UCase$(param.Value) = "1"), True, False)
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GSwing.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get HighAfterLow( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "HighAfterLow"
On Error GoTo Err

thePoint = HighAfterLowPoint(excludingCurrentBar).Value
HighAfterLow.BarNumber = thePoint.BarNumber
HighAfterLow.Timestamp = thePoint.Timestamp
HighAfterLow.Value = thePoint.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighAfterLowPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "HighAfterLowPoint"
On Error GoTo Err

If excludingCurrentBar Then
    If mPotentialSwingLowPoint.Value = DummyLow Then
        thePoint = mHighAfterLowPointExcl
    Else
        If mHighBeforeLowPoint.Value > mHighAfterLowPointExcl.Value Then
           thePoint = mHighBeforeLowPoint
        Else
            thePoint = mHighAfterLowPointExcl
        End If
    End If
ElseIf mPotentialSwingLowPoint.Value = DummyLow Then
    thePoint = mHighAfterLowPoint
Else
    If mHighBeforeLowPoint.Value > mHighAfterLowPoint.Value Then
        thePoint = mHighBeforeLowPoint
    Else
        thePoint = mHighAfterLowPoint
    End If
End If
HighAfterLowPoint.BarNumber = thePoint.BarNumber
HighAfterLowPoint.Timestamp = thePoint.Timestamp
HighAfterLowPoint.Value = thePoint

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowAfterHigh( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "LowAfterHigh"
On Error GoTo Err

thePoint = LowAfterHighPoint(excludingCurrentBar).Value
LowAfterHigh.BarNumber = thePoint.BarNumber
LowAfterHigh.Timestamp = thePoint.Timestamp
LowAfterHigh.Value = thePoint.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LowAfterHighPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "LowAfterHighPoint"
On Error GoTo Err

If excludingCurrentBar Then
    If mPotentialSwingHighPoint.Value = DummyHigh Then
        thePoint = mLowAfterHighPointExcl
    Else
        If mLowBeforeHighPoint.Value < mLowAfterHighPoint.Value Then
            thePoint = mLowBeforeHighPoint
        Else
            thePoint = mLowAfterHighPointExcl
        End If
    End If
ElseIf mPotentialSwingHighPoint.Value = DummyHigh Then
    thePoint = mLowAfterHighPoint
Else
    If mLowBeforeHighPoint.Value < mLowAfterHighPoint.Value Then
        thePoint = mLowBeforeHighPoint
    Else
        thePoint = mLowAfterHighPoint
    End If
End If
LowAfterHighPoint.BarNumber = thePoint.BarNumber
LowAfterHighPoint.Timestamp = thePoint.Timestamp
LowAfterHighPoint.Value = thePoint

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PotentialSwingHighPoint() As SwingPoint
PotentialSwingHighPoint = mPotentialSwingHighPoint
End Property

Public Property Get PotentialSwingLowPoint() As SwingPoint
PotentialSwingLowPoint = mPotentialSwingLowPoint
End Property

Public Property Get SwingHigh(Optional ByVal ref As Long) As SValue
Const ProcName As String = "SwingHigh"
On Error GoTo Err

SwingHigh = mSwingHighCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SwingHighPoint(Optional ByVal ref As Long) As SwingPoint
Const ProcName As String = "SwingHighPoint"
On Error GoTo Err

SwingHighPoint = mSwingHighCache.GetValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SwingLow(Optional ByVal ref As Long) As SValue
Const ProcName As String = "SwingLow"
On Error GoTo Err

SwingLow = mSwingLowCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SwingLowPoint(Optional ByVal ref As Long) As SwingPoint
Const ProcName As String = "SwingLowPoint"
On Error GoTo Err

SwingLowPoint = mSwingLowCache.GetValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Swing(Optional ByVal ref As Long) As SValue
Const ProcName As String = "Swing"
On Error GoTo Err

Swing = mSwingCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SwingPoint(Optional ByVal ref As Long) As SwingPoint
Const ProcName As String = "SwingPoint"
On Error GoTo Err

SwingPoint = mSwingCache.GetValue(ref)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Trend() As TrendDirections
Const ProcName As String = "trend"
On Error GoTo Err

Trend = mTrend

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcTrendDirection()
Const ProcName As String = "calcTrendDirection"
On Error GoTo Err

mTrend = TrendUnknown

If mSwingHighCache.Count < 2 Then Exit Sub
If mSwingLowCache.Count < 2 Then Exit Sub

Dim latestSwingLow As SwingPoint
latestSwingLow = SwingLow(0).Value
If IsEmpty(latestSwingLow.Value) Then Exit Sub

Dim prevSwingLow As SwingPoint
prevSwingLow = SwingLow(-1).Value
If IsEmpty(prevSwingLow.Value) Then Exit Sub

Dim latestSwingHigh As SwingPoint
latestSwingHigh = SwingHigh(0).Value
If IsEmpty(latestSwingHigh.Value) Then Exit Sub

Dim prevSwingHigh As SwingPoint
prevSwingHigh = SwingHigh(-1).Value
If IsEmpty(prevSwingHigh.Value) Then Exit Sub

If (mPotentialSwingHighPoint.Value > latestSwingHigh.Value Or latestSwingHigh.Value > prevSwingHigh.Value) And _
    mPotentialSwingLowPoint.Value >= latestSwingLow.Value And latestSwingLow.Value >= prevSwingLow.Value Then
    mTrend = TrendDirections.TrendUp
ElseIf (mPotentialSwingLowPoint.Value < latestSwingLow.Value Or latestSwingLow.Value < prevSwingLow.Value) And _
    mPotentialSwingHighPoint.Value <= latestSwingHigh.Value And latestSwingHigh.Value < prevSwingHigh.Value Then
    mTrend = TrendDirections.TrendDown
Else
    mTrend = TrendDirections.TrendSideways
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createImpliedSwingHigh(ByRef prevSwingHigh As SwingPoint)
Const ProcName As String = "createImpliedSwingHigh"
On Error GoTo Err

Dim impliedSwingHigh As SwingPoint
Dim impliedSwingHighStudyPoint As StudyPoint

If Not mIncludeImplicitSwingPoints Then Exit Sub

impliedSwingHigh = mHighBeforeLowPoint
impliedSwingHigh.Implied = True
impliedSwingHigh.SequenceNumber = mSwingSequenceNumber
mSwingSequenceNumber = mSwingSequenceNumber + 1
Set impliedSwingHighStudyPoint = CreateStudyPoint(impliedSwingHigh.Timestamp, impliedSwingHigh.Value)

'mHighBeforeLowPoint = mHighAfterLowPoint
mLowBeforeHighPoint = mLowAfterHighPoint

updateSwingHighType prevSwingHigh, impliedSwingHigh

mSwingHighCache.AddValue impliedSwingHigh, impliedSwingHigh.BarNumber, impliedSwingHigh.Timestamp, impliedSwingHigh.BarStartTime
notifySwingHighPoint

If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.Point2 = impliedSwingHighStudyPoint
    notifySwingHighLine
End If

Set mSwingHighLineBuilder = createStudyLineBuilder(impliedSwingHigh.Timestamp, impliedSwingHigh.Value)
mSwingHighLineCache.AddValue mSwingHighLineBuilder.StudyLine, impliedSwingHigh.BarNumber, impliedSwingHigh.Timestamp, impliedSwingHigh.BarStartTime

mSwingCache.AddValue impliedSwingHigh, impliedSwingHigh.BarNumber, impliedSwingHigh.Timestamp, impliedSwingHigh.BarStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = impliedSwingHighStudyPoint
    notifySwingLine
End If

Set mSwingLineBuilder = createStudyLineBuilder(impliedSwingHigh.Timestamp, impliedSwingHigh.Value)
mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, impliedSwingHigh.BarNumber, impliedSwingHigh.Timestamp, impliedSwingHigh.BarStartTime

notifySwingLine
notifySwingHighLine

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createImpliedSwingLow(ByRef prevSwingLow As SwingPoint)
Const ProcName As String = "createImpliedSwingLow"
On Error GoTo Err

Dim impliedSwingLow As SwingPoint
Dim impliedSwingLowStudyPoint As StudyPoint

If Not mIncludeImplicitSwingPoints Then Exit Sub

impliedSwingLow = mLowBeforeHighPoint
impliedSwingLow.Implied = True
impliedSwingLow.SequenceNumber = mSwingSequenceNumber
mSwingSequenceNumber = mSwingSequenceNumber + 1
Set impliedSwingLowStudyPoint = CreateStudyPoint(impliedSwingLow.Timestamp, impliedSwingLow.Value)

'mLowBeforeHighPoint = mLowAfterHighPoint
mHighBeforeLowPoint = mHighAfterLowPoint

updateSwingHighType prevSwingLow, impliedSwingLow

mSwingLowCache.AddValue impliedSwingLow, impliedSwingLow.BarNumber, impliedSwingLow.Timestamp, impliedSwingLow.BarStartTime
notifySwingLowPoint

If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.Point2 = impliedSwingLowStudyPoint
    notifySwingLowLine
End If

Set mSwingLowLineBuilder = createStudyLineBuilder(impliedSwingLow.Timestamp, impliedSwingLow.Value)
mSwingLowLineCache.AddValue mSwingLowLineBuilder.StudyLine, impliedSwingLow.BarNumber, impliedSwingLow.Timestamp, impliedSwingLow.BarStartTime

mSwingCache.AddValue impliedSwingLow, impliedSwingLow.BarNumber, impliedSwingLow.Timestamp, impliedSwingLow.BarStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = impliedSwingLowStudyPoint
    notifySwingLine
End If

Set mSwingLineBuilder = createStudyLineBuilder(impliedSwingLow.Timestamp, impliedSwingLow.Value)
mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, impliedSwingLow.BarNumber, impliedSwingLow.Timestamp, impliedSwingLow.BarStartTime

notifySwingLine
notifySwingLowLine

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createStudyLineBuilder(ByVal pTimestamp As Date, ByVal pValue As Variant) As StudyLineBuilder
Dim lPoint As StudyPoint
Const ProcName As String = "createStudyLineBuilder"
On Error GoTo Err

Set lPoint = CreateStudyPoint(pTimestamp, pValue)
Set createStudyLineBuilder = New StudyLineBuilder
createStudyLineBuilder.Point1 = lPoint
createStudyLineBuilder.Point2 = lPoint

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Const ProcName As String = "getStudyValue"
On Error GoTo Err

Select Case UCase$(valueName)
Case UCase$(SwingValueSwingPoint)
    getStudyValue = Swing(ref)
Case UCase$(SwingValueSwingHighPoint)
    getStudyValue = SwingHigh(ref)
Case UCase$(SwingValueSwingLowPoint)
    getStudyValue = SwingLow(ref)
Case UCase$(SwingValueSwingLine)
    getStudyValue = mSwingLineCache.GetSValue(ref)
Case UCase$(SwingValueSwingHighLine)
    getStudyValue = mSwingHighLineCache.GetSValue(ref)
Case UCase$(SwingValueSwingLowLine)
    getStudyValue = mSwingLowLineCache.GetSValue(ref)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialiseSwingHighPoint( _
                ByRef pSwingPoint As SwingPoint, _
                ByRef pSValue As SValue)
Const ProcName As String = "initialiseSwingHighPoint"
On Error GoTo Err

setSwingPointFromSValue pSwingPoint, pSValue
pSwingPoint.Value = DummyHigh
pSwingPoint.SwingType = SwingUnknown
pSwingPoint.Implied = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseSwingLowPoint( _
                ByRef pSwingPoint As SwingPoint, _
                ByRef pSValue As SValue)
Const ProcName As String = "initialiseSwingLowPoint"
On Error GoTo Err

setSwingPointFromSValue pSwingPoint, pSValue
pSwingPoint.Value = DummyLow
pSwingPoint.SwingType = SwingUnknown
pSwingPoint.Implied = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function inputValueHasNotChanged(ByVal pValue As Double) As Boolean
Const ProcName As String = "inputValueHasNotChanged"
On Error GoTo Err

Static sGotfirst As Boolean
Static sPrevVal As Double

If mStudyFoundation.IsNewBar Then
ElseIf Not sGotfirst Then
Else
    inputValueHasNotChanged = (pValue = sPrevVal)
End If
sGotfirst = True
sPrevVal = pValue

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub notifySwingHighLine()
Const ProcName As String = "notifySwingHighLine"
On Error GoTo Err

Dim sv As SValue
sv = mSwingHighLineCache.GetSValue(0)

Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal.BarNumber = sv.BarNumber
evOut.sVal.BarStartTime = sv.BarStartTime
evOut.sVal.Timestamp = sv.Timestamp
Set evOut.sVal.Value = mSwingHighLineBuilder.StudyLine
evOut.valueName = SwingValueSwingHighLine
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifySwingHighPoint()
Const ProcName As String = "notifySwingHighPoint"
On Error GoTo Err

Dim sp As SwingPoint
sp = mSwingHighCache.GetSValue(0).Value

Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal.BarNumber = sp.BarNumber
evOut.sVal.BarStartTime = sp.BarStartTime
evOut.sVal.Timestamp = sp.Timestamp
evOut.sVal.Value = sp.Value
evOut.valueName = SwingValueSwingHighPoint
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifySwingLine()
Const ProcName As String = "notifySwingLine"
On Error GoTo Err

Dim sv As SValue
sv = mSwingLineCache.GetSValue(0)

Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal.BarNumber = sv.BarNumber
evOut.sVal.BarStartTime = sv.BarStartTime
evOut.sVal.Timestamp = sv.Timestamp
Set evOut.sVal.Value = mSwingLineBuilder.StudyLine
evOut.valueName = SwingValueSwingLine
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifySwingLowLine()
Const ProcName As String = "notifySwingLowLine"
On Error GoTo Err

Dim sv As SValue
sv = mSwingLowLineCache.GetSValue(0)

Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal.BarNumber = sv.BarNumber
evOut.sVal.BarStartTime = sv.BarStartTime
evOut.sVal.Timestamp = sv.Timestamp
Set evOut.sVal.Value = mSwingLowLineBuilder.StudyLine
evOut.valueName = SwingValueSwingLowLine
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifySwingLowPoint()
Const ProcName As String = "notifySwingLowPoint"
On Error GoTo Err

Dim sp As SwingPoint
sp = mSwingLowCache.GetSValue(0).Value

Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal.BarNumber = sp.BarNumber
evOut.sVal.BarStartTime = sp.BarStartTime
evOut.sVal.Timestamp = sp.Timestamp
evOut.sVal.Value = sp.Value
evOut.valueName = SwingValueSwingLowPoint
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifySwingPoint()
Const ProcName As String = "notifySwingPoint"
On Error GoTo Err

Dim sp As SwingPoint
sp = mSwingCache.GetSValue(0).Value

Dim evOut As StudyValueEventData
Set evOut.Source = Me
evOut.sVal.BarNumber = sp.BarNumber
evOut.sVal.BarStartTime = sp.BarStartTime
evOut.sVal.Timestamp = sp.Timestamp
evOut.sVal.Value = sp.Value
evOut.valueName = SwingValueSwingPoint
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSwingPointFromSValue( _
                ByRef pPoint As SwingPoint, _
                ByRef pSValue As SValue)
Const ProcName As String = "setSwingPointFromSValue"
On Error GoTo Err

pPoint.BarNumber = pSValue.BarNumber
pPoint.BarStartTime = pSValue.BarStartTime
pPoint.Timestamp = pSValue.Timestamp
pPoint.Value = pSValue.Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub shiftBarHighsAndLows(ev As StudyValueEventData)
Const ProcName As String = "shiftBarHighsAndLows"
On Error GoTo Err

mPrevPrevBarLow = mPrevBarLow
mPrevBarLow = mCurrentBarLow
mCurrentBarLow = ev.sVal.Value
mPrevPrevBarHigh = mPrevBarHigh
mPrevBarHigh = mCurrentBarHigh
mCurrentBarHigh = ev.sVal.Value

mHighAfterLowPointExcl = mHighAfterLowPoint
mLowAfterHighPointExcl = mLowAfterHighPoint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateCurrentLineEnds(ev As StudyValueEventData)
Const ProcName As String = "updateCurrentLineEnds"
On Error GoTo Err

Dim lStudyPoint As StudyPoint
Set lStudyPoint = CreateStudyPoint(ev.sVal.Timestamp, ev.sVal.Value)
If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = lStudyPoint
    notifySwingLine
End If
If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.Point2 = lStudyPoint
    notifySwingHighLine
End If
If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.Point2 = lStudyPoint
    notifySwingLowLine
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateSwingHighType( _
                ByRef pPrevSwingHigh As SwingPoint, _
                ByRef pNewSwingHigh As SwingPoint)
Const ProcName As String = "updateSwingHighType"
On Error GoTo Err

If Not IsEmpty(pPrevSwingHigh.Value) Then
    If pNewSwingHigh.Value >= pPrevSwingHigh.Value Then
        pPrevSwingHigh.SwingType = SwingMinorHigh
        mSwingHighCache.UpdateValue pPrevSwingHigh, pPrevSwingHigh.BarNumber, pPrevSwingHigh.Timestamp
    Else
        If pPrevSwingHigh.SwingType = SwingUnknown Then
            pPrevSwingHigh.SwingType = SwingMajorHigh
            mSwingHighCache.UpdateValue pPrevSwingHigh, pPrevSwingHigh.BarNumber, pPrevSwingHigh.Timestamp
        End If
        pNewSwingHigh.SwingType = SwingMinorHigh
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateSwingLowType( _
                ByRef pPrevSwingLow As SwingPoint, _
                ByRef pNewSwingLow As SwingPoint)
Const ProcName As String = "updateSwingLowType"
On Error GoTo Err

If Not IsEmpty(pPrevSwingLow.Value) Then
    If pNewSwingLow.Value <= pPrevSwingLow.Value Then
        pPrevSwingLow.SwingType = SwingMinorLow
        mSwingLowCache.UpdateValue pPrevSwingLow, pPrevSwingLow.BarNumber, pPrevSwingLow.Timestamp
    Else
        If pPrevSwingLow.SwingType = SwingUnknown Then
            pPrevSwingLow.SwingType = SwingMajorLow
            mSwingLowCache.UpdateValue pPrevSwingLow, pPrevSwingLow.BarNumber, pPrevSwingLow.Timestamp
        End If
        pNewSwingLow.SwingType = SwingMinorLow
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' IStudy Interface Members
'@================================================================================

Private Function IStudy_AddStudy( _
                ByVal Study As IStudy, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudy"
On Error GoTo Err

Set IStudy_AddStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_AddStudyValueListener( _
                ByVal listener As IStudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "IStudy_AddStudyValueListener"
On Error GoTo Err

Set IStudy_AddStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_BaseStudy() As IStudy
Const ProcName As String = "IStudy_BaseStudy"
On Error GoTo Err

Set IStudy_BaseStudy = Me

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_GetStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
'IStudy_GetStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
Const ProcName As String = "IStudy_GetStudyValue"
On Error GoTo Err

IStudy_GetStudyValue = getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IStudy_GetValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "IStudy_GetValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_id() As String
Const ProcName As String = "IStudy_id"
On Error GoTo Err

IStudy_id = mStudyFoundation.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Initialise( _
                ByVal StudyManager As StudyManager, _
                ByVal Id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As IStudy, _
                ByVal pSessionFuture As IFuture)
Const ProcName As String = "IStudy_Initialise"
On Error GoTo Err

Set mStudyFoundation = New StudyFoundation
mStudyFoundation.Initialise StudyManager, _
                            Me, _
                            Id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise Id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy, _
                pSessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_InstanceName() As String
Const ProcName As String = "IStudy_InstanceName"
On Error GoTo Err

IStudy_InstanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudy_InstancePath() As String
Const ProcName As String = "IStudy_InstancePath"
On Error GoTo Err

IStudy_InstancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_Notify( _
                ev As StudyValueEventData)
Const ProcName As String = "IStudy_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_NumberOfBarsRequired() As Long
Const ProcName As String = "IStudy_NumberOfBarsRequired"
On Error GoTo Err

IStudy_NumberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IStudy_NumberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "IStudy_NumberOfCachedValues"
On Error GoTo Err

IStudy_NumberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IStudy_Parameters() As Parameters
Const ProcName As String = "IStudy_Parameters"
On Error GoTo Err

Set IStudy_Parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IStudy_RemoveStudyValueListener( _
                ByVal listener As IStudyValueListener)
Const ProcName As String = "IStudy_RemoveStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IStudy_SetInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
Const ProcName As String = "IStudy_SetInputTicksize"
On Error GoTo Err

If UCase$(inputName) = UCase$(SwingInputValue) Then mTickSize = tickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IStudy_StudyDefinition() As StudyDefinition
Const ProcName As String = "IStudy_StudyDefinition"
On Error GoTo Err

Set IStudy_StudyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



