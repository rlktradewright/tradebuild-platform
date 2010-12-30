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

Implements Study

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

Private mSwingHigh As SwingPoint
Private mSwingHighCache As ValueCache
Private mSwingHighLineBuilder As StudyLineBuilder
Private mSwingHighLineCache As ValueCache

Private mSwingLow As SwingPoint
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
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)

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

mSwingHigh.value = DummyHigh
mSwingLow.value = DummyLow

mHighBeforeLowPoint.value = DummyHigh
mHighAfterLowPoint.value = DummyHigh
mLowBeforeHighPoint.value = DummyLow
mLowAfterHighPoint.value = DummyLow

mTrend = TrendDirections.TrendUnknown

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub customNotify( _
                ev As StudyValueEventData)
Const ProcName As String = "customNotify"
On Error GoTo Err

Dim currValue As Double
Dim prevSwingHigh As SwingPoint
Dim prevSwingLow As SwingPoint

If IsEmpty(ev.sVal.value) Then Exit Sub

currValue = ev.sVal.value
If inputValueHasNotChanged(currValue) Then Exit Sub

updateCurrentLineEnds ev

If mSwingHighCache.Count >= 1 Then prevSwingHigh = mSwingHighCache.GetValue(0)
If mSwingLowCache.Count >= 1 Then prevSwingLow = mSwingLowCache.GetValue(0)

If mStudyFoundation.IsNewBar Then
    shiftBarHighsAndLows ev
    
    If mStudyFoundation.LocalBarNumber = 1 Then
        mCurrentBarHigh = currValue
        mCurrentBarLow = currValue
        
        initialiseSwingHighPoint mSwingHigh, ev.sVal
        mSwingHigh.value = Empty
        
        initialiseSwingLowPoint mSwingLow, ev.sVal
        mSwingLow.value = Empty
        
        mSwingHighCache.AddValue mSwingHigh, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHighLineCache.AddValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        
        mSwingLowCache.AddValue mSwingLow, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLowLineCache.AddValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
    
        mSwingCache.AddValue mSwingHigh, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLineCache.AddValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        
    End If
    
    If mSwingHigh.value <> DummyHigh Then
        If mStudyFoundation.barNumber - mSwingHigh.barNumber >= 3 And _
            (mSwingHigh.value - mPrevBarHigh) >= mMinimumSwing _
        Then
            updateSwingHighType prevSwingHigh, mSwingHigh

            If Not IsEmpty(prevSwingHigh.value) Then
                If prevSwingLow.SequenceNumber < prevSwingHigh.SequenceNumber Then
                    createImpliedSwingLow prevSwingLow
                    prevSwingLow = mSwingLowCache.GetValue(0)
                End If
            End If

            mSwingHigh.SequenceNumber = mSwingSequenceNumber
            mSwingSequenceNumber = mSwingSequenceNumber + 1
            
            mLowBeforeHighPoint = mLowAfterHighPoint

            mSwingHighCache.AddValue mSwingHigh, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            If Not mSwingHighLineBuilder Is Nothing Then
                mSwingHighLineBuilder.Point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
                notifySwingHighLine
            End If
            If Not mSwingLineBuilder Is Nothing Then
                mSwingLineBuilder.Point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
                notifySwingLine
            End If
            
            Set mSwingHighLineBuilder = createStudyLineBuilder(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingHighLineCache.AddValue mSwingHighLineBuilder.StudyLine, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            mSwingCache.AddValue mSwingHigh, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            Set mSwingLineBuilder = createStudyLineBuilder(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            notifySwingPoint
            notifySwingHighPoint
            
            notifySwingLine
            notifySwingHighLine
            
            prevSwingHigh = mSwingHigh
            initialiseSwingHighPoint mSwingHigh, ev.sVal
                
            ' we now need to discard any incompletely formed low that is earlier than
            ' the high just completed, since otherwise if it eventually completed it
            ' would be notified to listeners and studies with a time earlier than the
            ' high just notified
            If mSwingLow.value <> DummyLow And _
                mSwingLow.timestamp < prevSwingHigh.timestamp _
            Then
                initialiseSwingLowPoint mSwingLow, ev.sVal
            End If
        End If
    End If
    If mSwingLow.value <> DummyLow Then
        If mStudyFoundation.barNumber - mSwingLow.barNumber >= 3 And _
            (mPrevBarLow - mSwingLow.value) >= mMinimumSwing _
        Then
            updateSwingLowType prevSwingLow, mSwingLow

            If Not IsEmpty(prevSwingLow.value) Then
                If prevSwingHigh.SequenceNumber <= prevSwingLow.SequenceNumber Then
                    createImpliedSwingHigh prevSwingHigh
                    prevSwingHigh = mSwingHighCache.GetValue(0)
                End If
            End If

            mSwingLow.SequenceNumber = mSwingSequenceNumber
            mSwingSequenceNumber = mSwingSequenceNumber + 1
            
            mHighBeforeLowPoint = mHighAfterLowPoint

            mSwingLowCache.AddValue mSwingLow, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            If Not mSwingLowLineBuilder Is Nothing Then
                mSwingLowLineBuilder.Point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
                notifySwingLowLine
            End If
            If Not mSwingLineBuilder Is Nothing Then
                mSwingLineBuilder.Point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
                notifySwingLine
            End If
            
            Set mSwingLowLineBuilder = createStudyLineBuilder(mSwingLow.timestamp, mSwingLow.value)
            mSwingLowLineCache.AddValue mSwingLowLineBuilder.StudyLine, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            mSwingCache.AddValue mSwingLow, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            Set mSwingLineBuilder = createStudyLineBuilder(mSwingLow.timestamp, mSwingLow.value)
            mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            notifySwingPoint
            notifySwingLowPoint
            
            notifySwingLine
            notifySwingLowLine
            
            prevSwingLow = mSwingLow
            initialiseSwingLowPoint mSwingLow, ev.sVal
            
            ' we now need to discard any incompletely formed high that is earlier than
            ' the low just completed, since otherwise if it eventually completed it
            ' would be notified to listeners and studies with a time earlier than the
            ' low just notified
            If mSwingHigh.value <> DummyHigh And _
                mSwingHigh.timestamp < prevSwingLow.timestamp _
            Then
                initialiseSwingHighPoint mSwingHigh, ev.sVal
            End If
        End If
    End If

    calcTrendDirection

End If

If currValue < mCurrentBarLow Then mCurrentBarLow = currValue
If currValue > mCurrentBarHigh Then mCurrentBarHigh = currValue

If Not IsEmpty(prevSwingHigh.value) Then
    If currValue > prevSwingHigh.value Then
        If prevSwingHigh.SwingType = SwingUnknown Then
            prevSwingHigh.SwingType = SwingMinorHigh
            mSwingHighCache.UpdateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
        End If
    End If
End If
If Not IsEmpty(prevSwingLow.value) Then
    If currValue < prevSwingLow.value Then
        If prevSwingLow.SwingType = SwingUnknown Then
            prevSwingLow.SwingType = SwingMinorLow
            mSwingLowCache.UpdateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
        End If
    End If
End If

If (mSwingHigh.value = DummyHigh And _
    currValue >= mPrevBarHigh And _
    currValue >= mPrevPrevBarHigh) Or _
    (mSwingHigh.value <> DummyHigh And currValue >= mSwingHigh.value) _
Then
    setSwingPointFromSValue mSwingHigh, ev.sVal
    
    If mLowBeforeHighPoint.value >= mLowAfterHighPoint.value Then
        mLowBeforeHighPoint = mLowAfterHighPoint
    End If
    setSwingPointFromSValue mLowAfterHighPoint, ev.sVal
End If

If (mSwingLow.value = DummyLow And _
    currValue <= mPrevBarLow And _
    currValue <= mPrevPrevBarLow) Or _
    (mSwingLow.value <> DummyLow And currValue <= mSwingLow.value) _
Then
    setSwingPointFromSValue mSwingLow, ev.sVal
    
    If mHighBeforeLowPoint.value <= mHighAfterLowPoint.value Then
        mHighBeforeLowPoint = mHighAfterLowPoint
    End If
    setSwingPointFromSValue mHighAfterLowPoint, ev.sVal
End If

If currValue >= mHighAfterLowPoint.value Then
    setSwingPointFromSValue mHighAfterLowPoint, ev.sVal
End If
If currValue <= mLowAfterHighPoint.value Then
    setSwingPointFromSValue mLowAfterHighPoint, ev.sVal
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 0

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Const ProcName As String = "customProcessParameters"
On Error GoTo Err

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(SwingParamMinimumSwingTicks)
        mMinimumSwing = param.value * mTickSize
    Case UCase$(SwingParamIncludeImplicitSwingPoints)
        mIncludeImplicitSwingPoints = IIf((UCase$(param.value) = "Y") Or (UCase$(param.value) = "YES") Or (UCase$(param.value) = "T") Or (UCase$(param.value) = "TRUE") Or (UCase$(param.value) = "1"), True, False)
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GSwing.StudyDefinition

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

thePoint = HighAfterLowPoint(excludingCurrentBar).value
HighAfterLow.barNumber = thePoint.barNumber
HighAfterLow.timestamp = thePoint.timestamp
HighAfterLow.value = thePoint.value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HighAfterLowPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "HighAfterLowPoint"
On Error GoTo Err

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
HighAfterLowPoint.barNumber = thePoint.barNumber
HighAfterLowPoint.timestamp = thePoint.timestamp
HighAfterLowPoint.value = thePoint

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LowAfterHigh( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "LowAfterHigh"
On Error GoTo Err

thePoint = LowAfterHighPoint(excludingCurrentBar).value
LowAfterHigh.barNumber = thePoint.barNumber
LowAfterHigh.timestamp = thePoint.timestamp
LowAfterHigh.value = thePoint.value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LowAfterHighPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
Const ProcName As String = "LowAfterHighPoint"
On Error GoTo Err

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
LowAfterHighPoint.barNumber = thePoint.barNumber
LowAfterHighPoint.timestamp = thePoint.timestamp
LowAfterHighPoint.value = thePoint

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get swingHighPoint(Optional ByVal ref As Long) As SValue
Const ProcName As String = "swingHighPoint"
On Error GoTo Err

swingHighPoint = mSwingHighCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get swingHigh(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
Const ProcName As String = "swingHigh"
On Error GoTo Err

sh = mSwingHighCache.GetValue(ref)
swingHigh.barNumber = sh.barNumber
swingHigh.timestamp = sh.timestamp
swingHigh.value = sh.value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get swingLowPoint(Optional ByVal ref As Long) As SValue
Const ProcName As String = "swingLowPoint"
On Error GoTo Err

swingLowPoint = mSwingLowCache.GetSValue(ref)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get swingLow(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
Const ProcName As String = "swingLow"
On Error GoTo Err

sh = mSwingLowCache.GetValue(ref)
swingLow.barNumber = sh.barNumber
swingLow.timestamp = sh.timestamp
swingLow.value = sh.value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Swing(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
Const ProcName As String = "Swing"
On Error GoTo Err

sh = mSwingCache.GetValue(ref)
Swing.barNumber = sh.barNumber
Swing.timestamp = sh.timestamp
Swing.value = sh.value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SwingPoint(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
Const ProcName As String = "SwingPoint"
On Error GoTo Err

sh = mSwingCache.GetValue(ref)
Swing.barNumber = sh.barNumber
Swing.timestamp = sh.timestamp
Swing.value = sh

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get trend() As TrendDirections
Const ProcName As String = "trend"
On Error GoTo Err

trend = mTrend

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcTrendDirection()
Dim latestSwingLow As Double
Dim prevSwingLow As Double
Dim latestSwingHigh As Double
Dim prevSwingHigh As Double

Const ProcName As String = "calcTrendDirection"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Set impliedSwingHighStudyPoint = CreateStudyPoint(impliedSwingHigh.timestamp, impliedSwingHigh.value)

'mHighBeforeLowPoint = mHighAfterLowPoint
mLowBeforeHighPoint = mLowAfterHighPoint

updateSwingHighType prevSwingHigh, impliedSwingHigh

mSwingHighCache.AddValue impliedSwingHigh, impliedSwingHigh.barNumber, impliedSwingHigh.timestamp, impliedSwingHigh.barStartTime
notifySwingHighPoint

If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.Point2 = impliedSwingHighStudyPoint
    notifySwingHighLine
End If

Set mSwingHighLineBuilder = createStudyLineBuilder(impliedSwingHigh.timestamp, impliedSwingHigh.value)
mSwingHighLineCache.AddValue mSwingHighLineBuilder.StudyLine, impliedSwingHigh.barNumber, impliedSwingHigh.timestamp, impliedSwingHigh.barStartTime

mSwingCache.AddValue impliedSwingHigh, impliedSwingHigh.barNumber, impliedSwingHigh.timestamp, impliedSwingHigh.barStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = impliedSwingHighStudyPoint
    notifySwingLine
End If

Set mSwingLineBuilder = createStudyLineBuilder(impliedSwingHigh.timestamp, impliedSwingHigh.value)
mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, impliedSwingHigh.barNumber, impliedSwingHigh.timestamp, impliedSwingHigh.barStartTime

notifySwingLine
notifySwingHighLine

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Set impliedSwingLowStudyPoint = CreateStudyPoint(impliedSwingLow.timestamp, impliedSwingLow.value)

'mLowBeforeHighPoint = mLowAfterHighPoint
mHighBeforeLowPoint = mHighAfterLowPoint

updateSwingHighType prevSwingLow, impliedSwingLow

mSwingLowCache.AddValue impliedSwingLow, impliedSwingLow.barNumber, impliedSwingLow.timestamp, impliedSwingLow.barStartTime
notifySwingLowPoint

If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.Point2 = impliedSwingLowStudyPoint
    notifySwingLowLine
End If

Set mSwingLowLineBuilder = createStudyLineBuilder(impliedSwingLow.timestamp, impliedSwingLow.value)
mSwingLowLineCache.AddValue mSwingLowLineBuilder.StudyLine, impliedSwingLow.barNumber, impliedSwingLow.timestamp, impliedSwingLow.barStartTime

mSwingCache.AddValue impliedSwingLow, impliedSwingLow.barNumber, impliedSwingLow.timestamp, impliedSwingLow.barStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = impliedSwingLowStudyPoint
    notifySwingLine
End If

Set mSwingLineBuilder = createStudyLineBuilder(impliedSwingLow.timestamp, impliedSwingLow.value)
mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, impliedSwingLow.barNumber, impliedSwingLow.timestamp, impliedSwingLow.barStartTime

notifySwingLine
notifySwingLowLine

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
    getStudyValue = swingHigh(ref)
Case UCase$(SwingValueSwingLowPoint)
    getStudyValue = swingLow(ref)
Case UCase$(SwingValueSwingLine)
    getStudyValue = mSwingLineCache.GetSValue(ref)
Case UCase$(SwingValueSwingHighLine)
    getStudyValue = mSwingHighLineCache.GetSValue(ref)
Case UCase$(SwingValueSwingLowLine)
    getStudyValue = mSwingLowLineCache.GetSValue(ref)
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub initialiseSwingHighPoint( _
                ByRef pSwingPoint As SwingPoint, _
                ByRef pSValue As SValue)
setSwingPointFromSValue pSwingPoint, pSValue
pSwingPoint.value = DummyHigh
pSwingPoint.SwingType = SwingUnknown
pSwingPoint.Implied = False
End Sub

Private Sub initialiseSwingLowPoint( _
                ByRef pSwingPoint As SwingPoint, _
                ByRef pSValue As SValue)
setSwingPointFromSValue pSwingPoint, pSValue
pSwingPoint.value = DummyLow
pSwingPoint.SwingType = SwingUnknown
pSwingPoint.Implied = False
End Sub

Private Function inputValueHasNotChanged(ByVal pValue As Double) As Boolean
Static sGotfirst As Boolean
Static sPrevVal As Double

If mStudyFoundation.IsNewBar Then
ElseIf Not sGotfirst Then
Else
    inputValueHasNotChanged = (pValue = sPrevVal)
End If
sGotfirst = True
sPrevVal = pValue
End Function

Private Sub notifySwingHighLine()
Dim sv As SValue
Dim evOut As StudyValueEventData
Const ProcName As String = "notifySwingHighLine"
On Error GoTo Err

sv = mSwingHighLineCache.GetSValue(0)
Set evOut.Source = Me
evOut.sVal.barNumber = sv.barNumber
evOut.sVal.barStartTime = sv.barStartTime
evOut.sVal.timestamp = sv.timestamp
Set evOut.sVal.value = mSwingHighLineBuilder.StudyLine
evOut.valueName = SwingValueSwingHighLine
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifySwingHighPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEventData
Const ProcName As String = "notifySwingHighPoint"
On Error GoTo Err

sp = mSwingHighCache.GetSValue(0).value
Set evOut.Source = Me
evOut.sVal.barNumber = sp.barNumber
evOut.sVal.barStartTime = sp.barStartTime
evOut.sVal.timestamp = sp.timestamp
evOut.sVal.value = sp.value
evOut.valueName = SwingValueSwingHighPoint
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifySwingLine()
Dim sv As SValue
Dim evOut As StudyValueEventData
Const ProcName As String = "notifySwingLine"
On Error GoTo Err

sv = mSwingLineCache.GetSValue(0)
Set evOut.Source = Me
evOut.sVal.barNumber = sv.barNumber
evOut.sVal.barStartTime = sv.barStartTime
evOut.sVal.timestamp = sv.timestamp
Set evOut.sVal.value = mSwingLineBuilder.StudyLine
evOut.valueName = SwingValueSwingLine
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifySwingLowLine()
Dim sv As SValue
Dim evOut As StudyValueEventData
Const ProcName As String = "notifySwingLowLine"
On Error GoTo Err

sv = mSwingLowLineCache.GetSValue(0)
Set evOut.Source = Me
evOut.sVal.barNumber = sv.barNumber
evOut.sVal.barStartTime = sv.barStartTime
evOut.sVal.timestamp = sv.timestamp
Set evOut.sVal.value = mSwingLowLineBuilder.StudyLine
evOut.valueName = SwingValueSwingLowLine
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifySwingLowPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEventData
Const ProcName As String = "notifySwingLowPoint"
On Error GoTo Err

sp = mSwingLowCache.GetSValue(0).value
Set evOut.Source = Me
evOut.sVal.barNumber = sp.barNumber
evOut.sVal.barStartTime = sp.barStartTime
evOut.sVal.timestamp = sp.timestamp
evOut.sVal.value = sp.value
evOut.valueName = SwingValueSwingLowPoint
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub notifySwingPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEventData
Const ProcName As String = "notifySwingPoint"
On Error GoTo Err

sp = mSwingCache.GetSValue(0).value
Set evOut.Source = Me
evOut.sVal.barNumber = sp.barNumber
evOut.sVal.barStartTime = sp.barStartTime
evOut.sVal.timestamp = sp.timestamp
evOut.sVal.value = sp.value
evOut.valueName = SwingValueSwingPoint
mStudyFoundation.notifyValue evOut

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSwingPointFromSValue( _
                ByRef pPoint As SwingPoint, _
                ByRef pSValue As SValue)
pPoint.barNumber = pSValue.barNumber
pPoint.barStartTime = pSValue.barStartTime
pPoint.timestamp = pSValue.timestamp
pPoint.value = pSValue.value
End Sub

Private Sub shiftBarHighsAndLows(ev As StudyValueEventData)
mPrevPrevBarLow = mPrevBarLow
mPrevBarLow = mCurrentBarLow
mCurrentBarLow = ev.sVal.value
mPrevPrevBarHigh = mPrevBarHigh
mPrevBarHigh = mCurrentBarHigh
mCurrentBarHigh = ev.sVal.value

mHighAfterLowPointExcl = mHighAfterLowPoint
mLowAfterHighPointExcl = mLowAfterHighPoint
End Sub

Private Sub updateCurrentLineEnds(ev As StudyValueEventData)
Dim lStudyPoint As StudyPoint
Set lStudyPoint = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
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
End Sub

Private Sub updateSwingHighType( _
                ByRef pPrevSwingHigh As SwingPoint, _
                ByRef pNewSwingHigh As SwingPoint)
If Not IsEmpty(pPrevSwingHigh.value) Then
    If pNewSwingHigh.value >= pPrevSwingHigh.value Then
        pPrevSwingHigh.SwingType = SwingMinorHigh
        mSwingHighCache.UpdateValue pPrevSwingHigh, pPrevSwingHigh.barNumber, pPrevSwingHigh.timestamp
    Else
        If pPrevSwingHigh.SwingType = SwingUnknown Then
            pPrevSwingHigh.SwingType = SwingMajorHigh
            mSwingHighCache.UpdateValue pPrevSwingHigh, pPrevSwingHigh.barNumber, pPrevSwingHigh.timestamp
        End If
        pNewSwingHigh.SwingType = SwingMinorHigh
    End If
End If
End Sub

Private Sub updateSwingLowType( _
                ByRef pPrevSwingLow As SwingPoint, _
                ByRef pNewSwingLow As SwingPoint)
If Not IsEmpty(pPrevSwingLow.value) Then
    If pNewSwingLow.value <= pPrevSwingLow.value Then
        pPrevSwingLow.SwingType = SwingMinorLow
        mSwingLowCache.UpdateValue pPrevSwingLow, pPrevSwingLow.barNumber, pPrevSwingLow.timestamp
    Else
        If pPrevSwingLow.SwingType = SwingUnknown Then
            pPrevSwingLow.SwingType = SwingMajorLow
            mSwingLowCache.UpdateValue pPrevSwingLow, pPrevSwingLow.barNumber, pPrevSwingLow.timestamp
        End If
        pNewSwingLow.SwingType = SwingMinorLow
    End If
End If
End Sub

'@================================================================================
' From this point on the code is identical for all studies - do not change!!!!
'@================================================================================


'@================================================================================
' Study Interface Members
'@================================================================================

Private Function Study_addStudy( _
                ByVal Study As Study, _
                ByRef valueNames() As String, _
                ByVal numUnderlyingValuesToUse As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController

Const ProcName As String = "Study_addStudy"
On Error GoTo Err

Set Study_addStudy = mStudyFoundation.AddStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Const ProcName As String = "Study_addStudyValueListener"
On Error GoTo Err

Set Study_addStudyValueListener = mStudyFoundation.AddStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
'Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
Const ProcName As String = "Study_getStudyValue"
On Error GoTo Err

Study_getStudyValue = getStudyValue(valueName, ref)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As Session)
Const ProcName As String = "Study_initialise"
On Error GoTo Err

Set mStudyFoundation = New StudyFoundation
mStudyFoundation.initialise Me, _
                            id, _
                            Parameters, _
                            numberOfValuesToCache, _
                            inputValueNames, _
                            underlyingStudy
                            

customInitialise id, _
                Parameters, _
                numberOfValuesToCache, _
                inputValueNames, _
                underlyingStudy, _
                pSession

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEventData)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err

If UCase$(inputName) = UCase$(SwingInputValue) Then mTickSize = tickSize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property



