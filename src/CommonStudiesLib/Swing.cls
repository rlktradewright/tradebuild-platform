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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Sub customNotify( _
                ev As StudyValueEvent)
Static sPrevVal As Variant
Dim prevSwingHigh As SwingPoint
Dim prevSwingLow As SwingPoint

Const ProcName As String = "customNotify"
On Error GoTo Err

If IsEmpty(ev.sVal.value) Then Exit Sub

If Not mStudyFoundation.IsNewBar Then
    If ev.sVal.value = sPrevVal Then Exit Sub
End If
sPrevVal = ev.sVal.value

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
    notifySwingLine
End If
If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.Point2 = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
    notifySwingHighLine
End If
If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.Point2 = CreateStudyPoint(ev.sVal.timestamp, ev.sVal.value)
    notifySwingLowLine
End If

If mSwingHighCache.Count >= 1 Then prevSwingHigh = mSwingHighCache.GetValue(0)
If mSwingLowCache.Count >= 1 Then prevSwingLow = mSwingLowCache.GetValue(0)

If mStudyFoundation.IsNewBar Then
    mPrevPrevLow = mPrevLow
    mPrevLow = mLow
    mLow = ev.sVal.value
    mPrevPrevHigh = mPrevHigh
    mPrevHigh = mHigh
    mHigh = ev.sVal.value

    mHighAfterLowPointExcl = mHighAfterLowPoint
    mLowAfterHighPointExcl = mLowAfterHighPoint
    
    If mStudyFoundation.LocalBarNumber = 1 Then
        mSwingHigh.barNumber = ev.sVal.barNumber
        mSwingHigh.barStartTime = ev.sVal.barStartTime
        mSwingHigh.timestamp = ev.sVal.timestamp
        mSwingHigh.value = Empty
        mSwingCache.AddValue mSwingHigh, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLineCache.AddValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHighCache.AddValue mSwingHigh, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHighLineCache.AddValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingHigh.value = DummyHigh
        
        mSwingLow.barNumber = ev.sVal.barNumber
        mSwingLow.barStartTime = ev.sVal.barStartTime
        mSwingLow.timestamp = ev.sVal.timestamp
        mSwingLow.value = Empty
        mSwingLowCache.AddValue mSwingLow, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLowLineCache.AddValue Nothing, 1, ev.sVal.timestamp, ev.sVal.barStartTime
        mSwingLow.value = DummyLow
    End If
    
    If mSwingHigh.value <> DummyHigh Then
        If mStudyFoundation.barNumber - mSwingHigh.barNumber >= 3 And _
            (mSwingHigh.value - mPrevHigh) >= mMinimumSwing _
        Then
            If Not IsEmpty(prevSwingHigh.value) Then
                If mSwingHigh.value >= prevSwingHigh.value Then
                    prevSwingHigh.swingType = SwingMinorHigh
                    mSwingHighCache.UpdateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
                Else
                    If prevSwingHigh.swingType = SwingUnknownHigh Then
                        prevSwingHigh.swingType = SwingMajorHigh
                        mSwingHighCache.UpdateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
                    End If
                    mSwingHigh.swingType = SwingMinorHigh
                End If
            End If

            If Not IsEmpty(prevSwingHigh.value) Then
                If prevSwingLow.timestamp <= prevSwingHigh.timestamp Then
                    createImpliedSwingLow prevSwingLow
                    prevSwingLow = mSwingLowCache.GetValue(0)
                End If
            End If

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
            
            Set mSwingHighLineBuilder = New StudyLineBuilder
            mSwingHighLineBuilder.Point1 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingHighLineBuilder.Point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingHighLineCache.AddValue mSwingHighLineBuilder.StudyLine, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            mSwingCache.AddValue mSwingHigh, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
            Set mSwingLineBuilder = New StudyLineBuilder
            mSwingLineBuilder.Point1 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingLineBuilder.Point2 = CreateStudyPoint(mSwingHigh.timestamp, mSwingHigh.value)
            mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, mSwingHigh.barNumber, mSwingHigh.timestamp, mSwingHigh.barStartTime
            
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
                    mSwingLowCache.UpdateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
                Else
                    If prevSwingLow.swingType = SwingUnknownLow Then
                        prevSwingLow.swingType = SwingMajorLow
                        mSwingLowCache.UpdateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
                    End If
                    mSwingLow.swingType = SwingMinorLow
                End If
            End If

            If Not IsEmpty(prevSwingLow.value) Then
                If prevSwingHigh.timestamp <= prevSwingLow.timestamp Then
                    createImpliedSwingHigh prevSwingHigh
                    prevSwingHigh = mSwingHighCache.GetValue(0)
                End If
            End If

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
            
            Set mSwingLowLineBuilder = New StudyLineBuilder
            mSwingLowLineBuilder.Point1 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLowLineBuilder.Point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLowLineCache.AddValue mSwingLowLineBuilder.StudyLine, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            mSwingCache.AddValue mSwingLow, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
            Set mSwingLineBuilder = New StudyLineBuilder
            mSwingLineBuilder.Point1 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLineBuilder.Point2 = CreateStudyPoint(mSwingLow.timestamp, mSwingLow.value)
            mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, mSwingLow.barNumber, mSwingLow.timestamp, mSwingLow.barStartTime
            
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
            mSwingHighCache.UpdateValue prevSwingHigh, prevSwingHigh.barNumber, prevSwingHigh.timestamp
        End If
    End If
End If
If Not IsEmpty(prevSwingLow.value) Then
    If ev.sVal.value < prevSwingLow.value Then
        If prevSwingLow.swingType = SwingUnknownLow Then
            prevSwingLow.swingType = SwingMinorLow
            mSwingLowCache.UpdateValue prevSwingLow, prevSwingLow.barNumber, prevSwingLow.timestamp
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function customNumberOfBarsRequired() As Long
Const ProcName As String = "customNumberOfBarsRequired"
On Error GoTo Err

customNumberOfBarsRequired = 0

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function customStudyDefinition() As StudyDefinition
Const ProcName As String = "customStudyDefinition"
On Error GoTo Err

Set customStudyDefinition = GSwing.StudyDefinition

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get swingHighPoint(Optional ByVal ref As Long) As SValue
Const ProcName As String = "swingHighPoint"
On Error GoTo Err

swingHighPoint = mSwingHighCache.GetSValue(ref)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get swingLowPoint(Optional ByVal ref As Long) As SValue
Const ProcName As String = "swingLowPoint"
On Error GoTo Err

swingLowPoint = mSwingLowCache.GetSValue(ref)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get trend() As TrendDirections
Const ProcName As String = "trend"
On Error GoTo Err

trend = mTrend

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub createImpliedSwingHigh(ByRef prevSwingHigh As SwingPoint)
Dim tempSwingHigh As SwingPoint

Const ProcName As String = "createImpliedSwingHigh"
On Error GoTo Err

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
mSwingHighCache.AddValue tempSwingHigh, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime
notifySwingHighPoint

If Not mSwingHighLineBuilder Is Nothing Then
    mSwingHighLineBuilder.Point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
    notifySwingHighLine
End If

Set mSwingHighLineBuilder = New StudyLineBuilder
mSwingHighLineBuilder.Point1 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingHighLineBuilder.Point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingHighLineCache.AddValue mSwingHighLineBuilder.StudyLine, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime

mSwingCache.AddValue tempSwingHigh, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
    notifySwingLine
End If

Set mSwingLineBuilder = New StudyLineBuilder
mSwingLineBuilder.Point1 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingLineBuilder.Point2 = CreateStudyPoint(tempSwingHigh.timestamp, tempSwingHigh.value)
mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, tempSwingHigh.barNumber, tempSwingHigh.timestamp, tempSwingHigh.barStartTime

notifySwingLine
notifySwingHighLine

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub createImpliedSwingLow(ByRef prevSwingLow As SwingPoint)
Dim tempSwingLow As SwingPoint

Const ProcName As String = "createImpliedSwingLow"
On Error GoTo Err

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
mSwingLowCache.AddValue tempSwingLow, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime
notifySwingLowPoint

If Not mSwingLowLineBuilder Is Nothing Then
    mSwingLowLineBuilder.Point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
    notifySwingLowLine
End If

Set mSwingLowLineBuilder = New StudyLineBuilder
mSwingLowLineBuilder.Point1 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLowLineBuilder.Point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLowLineCache.AddValue mSwingLowLineBuilder.StudyLine, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime

mSwingCache.AddValue tempSwingLow, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime
notifySwingPoint

If Not mSwingLineBuilder Is Nothing Then
    mSwingLineBuilder.Point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
    notifySwingLine
End If

Set mSwingLineBuilder = New StudyLineBuilder
mSwingLineBuilder.Point1 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLineBuilder.Point2 = CreateStudyPoint(tempSwingLow.timestamp, tempSwingLow.value)
mSwingLineCache.AddValue mSwingLineBuilder.StudyLine, tempSwingLow.barNumber, tempSwingLow.timestamp, tempSwingLow.barStartTime

notifySwingLine
notifySwingLowLine

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub notifySwingHighLine()
Dim sv As SValue
Dim evOut As StudyValueEvent
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifySwingHighPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEvent
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifySwingLine()
Dim sv As SValue
Dim evOut As StudyValueEvent
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifySwingLowLine()
Dim sv As SValue
Dim evOut As StudyValueEvent
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifySwingLowPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEvent
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifySwingPoint()
Dim sp As SwingPoint
Dim evOut As StudyValueEvent
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_baseStudy() As Study
Const ProcName As String = "Study_baseStudy"
On Error GoTo Err

Set Study_baseStudy = Me

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function Study_getValueTicksize( _
                ByVal valueName As String) As Double

Const ProcName As String = "Study_getValueTicksize"
On Error GoTo Err



Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_id() As String
Const ProcName As String = "Study_id"
On Error GoTo Err

Study_id = mStudyFoundation.id

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_instanceName() As String
Const ProcName As String = "Study_instanceName"
On Error GoTo Err

Study_instanceName = mStudyFoundation.InstanceName

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get Study_instancePath() As String
Const ProcName As String = "Study_instancePath"
On Error GoTo Err

Study_instancePath = mStudyFoundation.InstancePath

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_Notify( _
                ev As StudyValueEvent)
Const ProcName As String = "Study_Notify"
On Error GoTo Err

mStudyFoundation.NotifyInput ev
customNotify ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_numberOfBarsRequired() As Long
Const ProcName As String = "Study_numberOfBarsRequired"
On Error GoTo Err

Study_numberOfBarsRequired = mStudyFoundation.NumberOfBarsRequired(customNumberOfBarsRequired)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Function Study_numberOfCachedValues( _
                Optional ByVal valueName As String) As Long
Const ProcName As String = "Study_numberOfCachedValues"
On Error GoTo Err

Study_numberOfCachedValues = mStudyFoundation.numberOfCachedValues(valueName)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Property Get Study_parameters() As Parameters
Const ProcName As String = "Study_parameters"
On Error GoTo Err

Set Study_parameters = mStudyFoundation.Parameters

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub Study_removeStudyValueListener( _
                ByVal listener As StudyValueListener)
Const ProcName As String = "Study_removeStudyValueListener"
On Error GoTo Err

mStudyFoundation.RemoveStudyValueListener listener

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub Study_setInputTicksize( _
                ByVal inputName As String, _
                ByVal tickSize As Double)
Const ProcName As String = "Study_setInputTicksize"
On Error GoTo Err

If UCase$(inputName) = UCase$(SwingInputValue) Then mTickSize = tickSize

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Property Get Study_studyDefinition() As StudyDefinition
Const ProcName As String = "Study_studyDefinition"
On Error GoTo Err

Set Study_studyDefinition = customStudyDefinition

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property



