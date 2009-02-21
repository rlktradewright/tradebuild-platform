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




Private mHigh As Double
Private mPrevHigh As Double
Private mPrevPrevHigh As Double
Private mLow As Double
Private mPrevLow As Double
Private mPrevPrevLow As Double

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
                ByVal pSession As session)

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

End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim prevSwingHigh As SwingPoint
Dim prevSwingLow As SwingPoint

Dim failpoint As Long
On Error GoTo Err

If IsEmpty(ev.sVal.value) Then Exit Sub

If Not mStudyFoundation.isNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.point2 = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
    notifySwingLine
End If
If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.point2 = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
    notifySwingHighLine
End If
If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.point2 = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
    notifySwingLowLine
End If

If mSwingHighCache.Count >= 1 Then prevSwingHigh = mSwingHighCache.getValue(0)
If mSwingLowCache.Count >= 1 Then prevSwingLow = mSwingLowCache.getValue(0)

If mStudyFoundation.isNewBar Then
    mPrevPrevLow = mPrevLow
    mPrevLow = mLow
    mLow = ev.sVal.value
    mPrevPrevHigh = mPrevHigh
    mPrevHigh = mHigh
    mHigh = ev.sVal.value

    mHighAfterLowPointExcl = mHighAfterLowPoint
    mLowAfterHighPointExcl = mLowAfterHighPoint
    
    If mStudyFoundation.localBarNumber = 1 Then
        mSwingHigh.barNumber = ev.sVal.barNumber
        mSwingHigh.barStartTime = ev.sVal.barStartTime
        mSwingHigh.timestamp = ev.sVal.timestamp
        mSwingHigh.value = Empty
        mSwingCache.addValue mSwingHigh, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLineCache.addValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHighCache.addValue mSwingHigh, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHighLineCache.addValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHigh.value = DummyHigh
        
        mSwingLow.barNumber = ev.sVal.barNumber
        mSwingLow.barStartTime = ev.sVal.barStartTime
        mSwingLow.timestamp = ev.sVal.timestamp
        mSwingLow.value = Empty
        mSwingLowCache.addValue mSwingLow, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLowLineCache.addValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLow.value = DummyLow
    End If
    
    If mSwingHigh.value <> DummyHigh Then
        If mStudyFoundation.barNumber - mSwingHigh.barNumber >= 3 And _
            (mSwingHigh.value - mPrevHigh) >= mMinimumSwing _
        Then
            If Not IsEmpty(prevSwingHigh.value) Then
                If mSwingHigh.value >= prevSwingHigh.value Then
                    prevSwingHigh.swingType = SwingMinorHigh
                    mSwingHighCache.updateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
                Else
                    If prevSwingHigh.swingType = SwingUnknownHigh Then
                        prevSwingHigh.swingType = SwingMajorHigh
                        mSwingHighCache.updateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
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

            mSwingHighCache.addValue mSwingHigh, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            If Not mSwingHighLineBuilder Is Nothing Then
                mSwingHighLineBuilder.point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
                notifySwingHighLine
            End If
            If Not mSwingLineBuilder Is Nothing Then
                mSwingLineBuilder.point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
                notifySwingLine
            End If
            
            Set mSwingHighLineBuilder = New StudyLineBuilder
            mSwingHighLineBuilder.point1 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingHighLineBuilder.point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingHighLineCache.addValue mSwingHighLineBuilder.StudyLine, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            mSwingCache.addValue mSwingHigh, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            Set mSwingLineBuilder = New StudyLineBuilder
            mSwingLineBuilder.point1 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingLineBuilder.point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingLineCache.addValue mSwingLineBuilder.StudyLine, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            notifySwingPoint
            notifySwingHighPoint
            
            notifySwingLine
            notifySwingHighLine
            
            prevSwingHigh = mSwingHigh
            mSwingHigh.value = DummyHigh
            mSwingHigh.barNumber = ev.sVal.barNumber
            mSwingHigh.barStartTime = ev.sVal.barStartTime
            mSwingHigh.timestamp = ev.sVal.timestamp
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
                mSwingLow.barNumber = ev.sVal.barNumber
                mSwingLow.barStartTime = ev.sVal.barStartTime
                mSwingLow.timestamp = ev.sVal.timestamp
                mSwingLow.swingType = SwingUnknownLow
                mSwingLow.implied = False
            End If
        End If
    End If
    If mSwingLow.value <> DummyLow Then
        If mStudyFoundation.barNumber - mSwingLow.barNumber >= 3 And _
            (mPrevLow - mSwingLow.value) >= mMinimumSwing _
        Then
            If Not IsEmpty(prevSwingLow.value) Then
                If mSwingLow.value <= prevSwingLow.value Then
                    prevSwingLow.swingType = SwingMinorLow
                    mSwingLowCache.updateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
                Else
                    If prevSwingLow.swingType = SwingUnknownLow Then
                        prevSwingLow.swingType = SwingMajorLow
                        mSwingLowCache.updateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
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

            mSwingLowCache.addValue mSwingLow, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            If Not mSwingLowLineBuilder Is Nothing Then
                mSwingLowLineBuilder.point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
                notifySwingLowLine
            End If
            If Not mSwingLineBuilder Is Nothing Then
                mSwingLineBuilder.point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
                notifySwingLine
            End If
            
            Set mSwingLowLineBuilder = New StudyLineBuilder
            mSwingLowLineBuilder.point1 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLowLineBuilder.point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLowLineCache.addValue mSwingLowLineBuilder.StudyLine, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            mSwingCache.addValue mSwingLow, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            Set mSwingLineBuilder = New StudyLineBuilder
            mSwingLineBuilder.point1 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLineBuilder.point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLineCache.addValue mSwingLineBuilder.StudyLine, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            notifySwingPoint
            notifySwingLowPoint
            
            notifySwingLine
            notifySwingLowLine
            
            prevSwingLow = mSwingLow
            mSwingLow.value = DummyLow
            mSwingLow.barNumber = ev.sVal.barNumber
            mSwingLow.barStartTime = ev.sVal.barStartTime
            mSwingLow.timestamp = ev.sVal.timestamp
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
                mSwingHigh.barNumber = ev.sVal.barNumber
                mSwingHigh.barStartTime = ev.sVal.barStartTime
                mSwingHigh.timestamp = ev.sVal.timestamp
                mSwingHigh.swingType = SwingUnknownHigh
                mSwingHigh.implied = False
            End If
        End If
    End If

    calcTrendDirection

End If

If ev.sVal.value < mLow Then mLow = ev.sVal.value
If ev.sVal.value > mHigh Then mHigh = ev.sVal.value

If Not IsEmpty(prevSwingHigh.value) Then
    If ev.sVal.value > prevSwingHigh.value Then
        If prevSwingHigh.swingType = SwingUnknownHigh Then
            prevSwingHigh.swingType = SwingMinorHigh
            mSwingHighCache.updateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
        End If
    End If
End If
If Not IsEmpty(prevSwingLow.value) Then
    If ev.sVal.value < prevSwingLow.value Then
        If prevSwingLow.swingType = SwingUnknownLow Then
            prevSwingLow.swingType = SwingMinorLow
            mSwingLowCache.updateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
        End If
    End If
End If

If (mSwingHigh.value = DummyHigh And _
    ev.sVal.value >= mPrevHigh And _
    ev.sVal.value >= mPrevPrevHigh) Or _
    (mSwingHigh.value <> DummyHigh And ev.sVal.value >= mSwingHigh.value) _
Then
    mSwingHigh.value = ev.sVal.value
    mSwingHigh.barNumber = ev.sVal.barNumber
    mSwingHigh.barStartTime = ev.sVal.barStartTime
    mSwingHigh.timestamp = ev.sVal.timestamp
    
    If mLowBeforeHighPoint.value >= mLowAfterHighPoint.value Then
        mLowBeforeHighPoint = mLowAfterHighPoint
    End If
    mLowAfterHighPoint.value = ev.sVal.value
    mLowAfterHighPoint.timestamp = ev.sVal.timestamp
    mLowAfterHighPoint.barNumber = ev.sVal.barNumber
    mLowAfterHighPoint.barStartTime = ev.sVal.barStartTime
End If

If (mSwingLow.value = DummyLow And _
    ev.sVal.value <= mPrevLow And _
    ev.sVal.value <= mPrevPrevLow) Or _
    (mSwingLow.value <> DummyLow And ev.sVal.value <= mSwingLow.value) _
Then
    mSwingLow.value = ev.sVal.value
    mSwingLow.barNumber = ev.sVal.barNumber
    mSwingLow.barStartTime = ev.sVal.barStartTime
    mSwingLow.timestamp = ev.sVal.timestamp
    
    If mHighBeforeLowPoint.value <= mHighAfterLowPoint.value Then
        mHighBeforeLowPoint = mHighAfterLowPoint
    End If
    mHighAfterLowPoint.value = ev.sVal.value
    mHighAfterLowPoint.timestamp = ev.sVal.timestamp
    mHighAfterLowPoint.barNumber = ev.sVal.barNumber
    mHighAfterLowPoint.barStartTime = ev.sVal.barStartTime
End If

If ev.sVal.value >= mHighAfterLowPoint.value Then
    mHighAfterLowPoint.value = ev.sVal.value
    mHighAfterLowPoint.timestamp = ev.sVal.timestamp
    mHighAfterLowPoint.barNumber = ev.sVal.barNumber
    mHighAfterLowPoint.barStartTime = ev.sVal.barStartTime
End If
If ev.sVal.value <= mLowAfterHighPoint.value Then
    mLowAfterHighPoint.value = ev.sVal.value
    mLowAfterHighPoint.timestamp = ev.sVal.timestamp
    mLowAfterHighPoint.barNumber = ev.sVal.barNumber
    mLowAfterHighPoint.barStartTime = ev.sVal.barStartTime
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customNotify" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function customNumberOfBarsRequired() As Long
customNumberOfBarsRequired = 0
End Function

Private Sub customProcessParameters(ByVal params As Parameters)
Dim param As Parameter

Dim failpoint As Long
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
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "customProcessParameters" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Private Function customStudyDefinition() As StudyDefinition
Set customStudyDefinition = GSwing.StudyDefinition
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
thePoint = HighAfterLowPoint(excludingCurrentBar).value
HighAfterLow.barNumber = thePoint.barNumber
HighAfterLow.timestamp = thePoint.timestamp
HighAfterLow.value = thePoint.value
End Property

Public Property Get HighAfterLowPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
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
HighAfterLowPoint.barNumber = thePoint.barNumber
HighAfterLowPoint.timestamp = thePoint.timestamp
HighAfterLowPoint.value = thePoint
End Property

Public Property Get LowAfterHigh( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
Dim thePoint As SwingPoint
thePoint = LowAfterHighPoint(excludingCurrentBar).value
LowAfterHigh.barNumber = thePoint.barNumber
LowAfterHigh.timestamp = thePoint.timestamp
LowAfterHigh.value = thePoint.value
End Property

Public Property Get LowAfterHighPoint( _
                Optional ByVal excludingCurrentBar As Boolean) As SValue
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
LowAfterHighPoint.barNumber = thePoint.barNumber
LowAfterHighPoint.timestamp = thePoint.timestamp
LowAfterHighPoint.value = thePoint
End Property

Public Property Get swingHighPoint(Optional ByVal ref As Long) As SValue
swingHighPoint = mSwingHighCache.getSValue(ref)
End Property

Public Property Get swingHigh(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
sh = mSwingHighCache.getValue(ref)
swingHigh.barNumber = sh.barNumber
swingHigh.timestamp = sh.timestamp
swingHigh.value = sh.value
End Property

Public Property Get swingLowPoint(Optional ByVal ref As Long) As SValue
swingLowPoint = mSwingLowCache.getSValue(ref)
End Property

Public Property Get swingLow(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
sh = mSwingLowCache.getValue(ref)
swingLow.barNumber = sh.barNumber
swingLow.timestamp = sh.timestamp
swingLow.value = sh.value
End Property

Public Property Get Swing(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
sh = mSwingCache.getValue(ref)
Swing.barNumber = sh.barNumber
Swing.timestamp = sh.timestamp
Swing.value = sh.value
End Property

Public Property Get SwingPoint(Optional ByVal ref As Long) As SValue
Dim sh As SwingPoint
sh = mSwingCache.getValue(ref)
Swing.barNumber = sh.barNumber
Swing.timestamp = sh.timestamp
Swing.value = sh
End Property

Public Property Get trend() As TrendDirections
trend = mTrend
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

If Not mIncludeImplicitSwingPoints Then Exit Sub

tempSwingHigh.implied = True
tempSwingHigh.value = mHighBeforeLowPoint.value
tempSwingHigh.timestamp = mHighBeforeLowPoint.timestamp
tempSwingHigh.barNumber = mHighBeforeLowPoint.barNumber
tempSwingHigh.barStartTime = mHighBeforeLowPoint.barStartTime

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
mSwingHighCache.addValue tempSwingHigh, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime
notifySwingHighPoint

If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
    notifySwingHighLine
End If

Set mSwingHighLineBuilder = New StudyLineBuilder
mSwingHighLineBuilder.point1 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingHighLineBuilder.point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingHighLineCache.addValue mSwingHighLineBuilder.StudyLine, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime

mSwingCache.addValue tempSwingHigh, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
    notifySwingLine
End If

Set mSwingLineBuilder = New StudyLineBuilder
mSwingLineBuilder.point1 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingLineBuilder.point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingLineCache.addValue mSwingLineBuilder.StudyLine, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime

notifySwingLine
notifySwingHighLine
End Sub

Private Sub createImpliedSwingLow(ByRef prevSwingLow As SwingPoint)
Dim tempSwingLow As SwingPoint

If Not mIncludeImplicitSwingPoints Then Exit Sub

tempSwingLow.implied = True
tempSwingLow.value = mLowBeforeHighPoint.value
tempSwingLow.timestamp = mLowBeforeHighPoint.timestamp
tempSwingLow.barNumber = mLowBeforeHighPoint.barNumber
tempSwingLow.barStartTime = mLowBeforeHighPoint.barStartTime

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
mSwingLowCache.addValue tempSwingLow, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime
notifySwingLowPoint

If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
    notifySwingLowLine
End If

Set mSwingLowLineBuilder = New StudyLineBuilder
mSwingLowLineBuilder.point1 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLowLineBuilder.point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLowLineCache.addValue mSwingLowLineBuilder.StudyLine, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime

mSwingCache.addValue tempSwingLow, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
    notifySwingLine
End If

Set mSwingLineBuilder = New StudyLineBuilder
mSwingLineBuilder.point1 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLineBuilder.point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLineCache.addValue mSwingLineBuilder.StudyLine, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime

notifySwingLine
notifySwingLowLine
End Sub

Private Function getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
Select Case UCase$(valueName)
Case UCase$(SwingValueSwingPoint)
    getStudyValue = Swing(ref)
Case UCase$(SwingValueSwingHighPoint)
    getStudyValue = swingHigh(ref)
Case UCase$(SwingValueSwingLowPoint)
    getStudyValue = swingLow(ref)
Case UCase$(SwingValueSwingLine)
    getStudyValue = mSwingLineCache.getSValue(ref)
Case UCase$(SwingValueSwingHighLine)
    getStudyValue = mSwingHighLineCache.getSValue(ref)
Case UCase$(SwingValueSwingLowLine)
    getStudyValue = mSwingLowLineCache.getSValue(ref)
End Select
End Function

Private Sub notifySwingHighLine()
Dim sv As SValue
Dim evOut As StudyValueEvent
sv = mSwingHighLineCache.getSValue(0)
Set evOut.Source = Me
evOut.sVal.barNumber = sv.barNumber
evOut.sVal.barStartTime = sv.barStartTime
evOut.sVal.timestamp = sv.timestamp
Set evOut.sVal.value = mSwingHighLineBuilder.StudyLine
evOut.valueName = SwingValueSwingHighLine
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifySwingHighPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEvent
sp = mSwingHighCache.getSValue(0).value
Set evOut.Source = Me
evOut.sVal.barNumber = sp.barNumber
evOut.sVal.barStartTime = sp.barStartTime
evOut.sVal.timestamp = sp.timestamp
evOut.sVal.value = sp.value
evOut.valueName = SwingValueSwingHighPoint
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifySwingLine()
Dim sv As SValue
Dim evOut As StudyValueEvent
sv = mSwingLineCache.getSValue(0)
Set evOut.Source = Me
evOut.sVal.barNumber = sv.barNumber
evOut.sVal.barStartTime = sv.barStartTime
evOut.sVal.timestamp = sv.timestamp
Set evOut.sVal.value = mSwingLineBuilder.StudyLine
evOut.valueName = SwingValueSwingLine
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifySwingLowLine()
Dim sv As SValue
Dim evOut As StudyValueEvent
sv = mSwingLowLineCache.getSValue(0)
Set evOut.Source = Me
evOut.sVal.barNumber = sv.barNumber
evOut.sVal.barStartTime = sv.barStartTime
evOut.sVal.timestamp = sv.timestamp
Set evOut.sVal.value = mSwingLowLineBuilder.StudyLine
evOut.valueName = SwingValueSwingLowLine
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifySwingLowPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEvent
sp = mSwingLowCache.getSValue(0).value
Set evOut.Source = Me
evOut.sVal.barNumber = sp.barNumber
evOut.sVal.barStartTime = sp.barStartTime
evOut.sVal.timestamp = sp.timestamp
evOut.sVal.value = sp.value
evOut.valueName = SwingValueSwingLowPoint
mStudyFoundation.notifyValue evOut
End Sub

Private Sub notifySwingPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEvent
sp = mSwingCache.getSValue(0).value
Set evOut.Source = Me
evOut.sVal.barNumber = sp.barNumber
evOut.sVal.barStartTime = sp.barStartTime
evOut.sVal.timestamp = sp.timestamp
evOut.sVal.value = sp.value
evOut.valueName = SwingValueSwingPoint
mStudyFoundation.notifyValue evOut
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

Set Study_addStudy = mStudyFoundation.addStudy( _
                            Study, _
                            valueNames, _
                            numUnderlyingValuesToUse, _
                            taskName, _
                            taskData)
End Function

Private Function Study_addStudyValueListener( _
                ByVal listener As StudyValueListener, _
                ByVal valueName As String, _
                ByVal numberOfValuesToReplay As Long, _
                Optional ByVal taskName As String, _
                Optional ByVal taskData As Variant) As TaskController
Set Study_addStudyValueListener = mStudyFoundation.addStudyValueListener( _
                            listener, _
                            valueName, _
                            numberOfValuesToReplay, _
                            taskName, _
                            taskData)
End Function

Private Property Get Study_baseStudy() As Study
Set Study_baseStudy = Me
End Property

Private Function Study_getStudyValue( _
                ByVal valueName As String, _
                ByVal ref As Long) As SValue
'Study_getStudyValue = mStudyFoundation.getStudyValue(valueName, ref)
Study_getStudyValue = getStudyValue(valueName, ref)
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

End Function

Private Property Get Study_id() As String
Study_id = mStudyFoundation.id
End Property

Private Sub Study_initialise( _
                ByVal id As String, _
                ByVal Parameters As Parameters, _
                ByVal numberOfValuesToCache As Long, _
                ByRef inputValueNames() As String, _
                ByVal underlyingStudy As Study, _
                ByVal pSession As session)
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
End Sub

Private Property Get Study_instanceName() As String
Study_instanceName = mStudyFoundation.instanceName
End Property

Private Property Get Study_instancePath() As String
Study_instancePath = mStudyFoundation.instancePath
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
mStudyFoundation.notifyInput ev
customNotify ev
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Study_numberOfBarsRequired = mStudyFoundation.numberOfBarsRequired(customNumberOfBarsRequired)
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)
End Function

Private Property Get Study_parameters() As Parameters
Set Study_parameters = mStudyFoundation.Parameters
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
mStudyFoundation.removeStudyValueListener listener
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
If inputName = SwingInputValue Then mTickSize = tickSize
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Set Study_studyDefinition = customStudyDefinition
End Property



