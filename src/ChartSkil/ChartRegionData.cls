VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements Task

'================================================================================
' Events
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "ChartRegionData"

'================================================================================
' Types
'================================================================================

Private Type PeriodTableEntry
    firstIndex                  As Long
    lastindex                   As Long
    BackColor                   As Long
End Type

Private Type GraphicObjectTableEntry
    graphObj            As IGraphicObject
    PeriodNumber        As Long
    Layer               As Long
    prevPeriodIndex     As Long
    nextPeriodIndex     As Long
    prevLayerIndex      As Long
    nextLayerIndex      As Long
    prevExtendedIndex   As Long
    nextExtendedIndex   As Long
    inScopeCycle        As Long
    prevInScopeIndex    As Long
    nextInScopeIndex    As Long
    visibleCycle        As Long
    prevVisibleIndex    As Long
    nextVisibleIndex    As Long
End Type

Private Type VisibleTableEntry
    firstIndex          As Long
    lastindex           As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mChartRef As WeakReference
Private mRegionRef As WeakReference

Private mPeriodTable() As PeriodTableEntry

Private mGraphicObjects() As GraphicObjectTableEntry
Private mFirstGraphicObjectIndex As Long    ' the index of the first table entry
                                            ' that actually points to a graphic
                                            ' object - all prior entries are
                                            ' merely list start and end entries
Private mNextGraphicObjectIndex As Long

Private mMaxPeriodNumber As Long

Private mInScopeCycle As Long
Private mFirstInScopeIndex As Long
Private mLastInScopeIndex As Long

Private mVisibleCycle As Long
Private mVisibleTable(MinLayer To MaxLayer) As VisibleTableEntry

Private mFirstExtendedIndex As Long
Private mLastExtendedIndex As Long

Private mTaskContext As TaskContext

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Dim i As Long

ReDim mPeriodTable(255) As PeriodTableEntry
ReDim mGraphicObjects(1023) As GraphicObjectTableEntry

mFirstInScopeIndex = allocateGraphicObjectTableEntry
mLastInScopeIndex = allocateGraphicObjectTableEntry
initialiseInScopeObjectList

mFirstExtendedIndex = allocateGraphicObjectTableEntry
mLastExtendedIndex = allocateGraphicObjectTableEntry
mGraphicObjects(mFirstExtendedIndex).nextExtendedIndex = mLastExtendedIndex
mGraphicObjects(mLastExtendedIndex).prevExtendedIndex = mFirstExtendedIndex

For i = MinLayer To MaxLayer
    mVisibleTable(i).firstIndex = allocateGraphicObjectTableEntry
    mVisibleTable(i).lastindex = allocateGraphicObjectTableEntry
Next
initialiseVisibleObjectList

End Sub

Private Sub Class_Terminate()
Debug.Print "ChartRegionData terminated"
End Sub

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub Task_run()
Static i As Long
Dim j As Long

Do While i < mNextGraphicObjectIndex
    If Not mGraphicObjects(i).graphObj Is Nothing Then Set mGraphicObjects(i).graphObj = Nothing
    i = i + 1
    j = j + 1
    If j = 100 Then Exit Sub
Loop

Erase mGraphicObjects

mTaskContext.Finish 0, False

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TWUtilities30.TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String
Task_TaskName = mTaskContext.Name
End Property

'================================================================================
' Properties
'================================================================================

Friend Property Get MaxPeriodNumber() As Long
MaxPeriodNumber = mMaxPeriodNumber
End Property

Friend Property Let PeriodBackColor( _
                ByVal PeriodNumber As Long, _
                ByVal value As Long)
mPeriodTable(PeriodNumber).BackColor = value
End Property

Friend Property Get PeriodBackColor( _
                ByVal PeriodNumber As Long) As Long
PeriodBackColor = mPeriodTable(PeriodNumber).BackColor
End Property

'================================================================================
' Methods
'================================================================================

Friend Function AddGraphicObject( _
                ByVal graphObj As IGraphicObject) As Long
Dim handle As Long

handle = allocateGraphicObjectTableEntry
If mFirstGraphicObjectIndex = 0 Then
    mFirstGraphicObjectIndex = handle
End If

graphObj.handle = handle

With mGraphicObjects(handle)
    Set .graphObj = graphObj
    .Layer = graphObj.Layer
End With

If Not graphObj.ExtendedObject Then
    ' NB: the following call to addEntryToPeriodList must not be within a With block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    addEntryToPeriodList handle
Else
    addEntryToExtendedList handle
End If

If graphicObjectIsInScope(handle) Then
    addEntryToInScopeList handle
    If graphicObjectIsVisible(handle) Then
        addEntryToVisibleList handle
    End If
End If
AddGraphicObject = handle
End Function

Friend Sub AddPeriod(ByVal PeriodNumber As Long)
Dim failpoint As Long
On Error GoTo Err

addNewPeriod PeriodNumber

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "AddPeriod" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Friend Function CalculateInscopeGaugeDependentObjects()
Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
     With mGraphicObjects(index)
        If .graphObj.GaugeDependent Then
            Dim boundingRect As TRectangle
            boundingRect = .graphObj.BoundingRectangle
            If graphicObjectIsVisible(index) Then
                addEntryToVisibleList index
            End If
        End If
        index = .nextInScopeIndex
    End With
Loop

End Function

Friend Sub CalcVerticalBoundsForGaugeDependentObjects( _
                ByRef scaleMin As Double, _
                ByRef scaleMax As Double)
Dim objMax As Double
Dim objMin As Double
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex

Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .graphObj.IncludeInAutoscale And _
            .graphObj.GaugeDependent _
        Then
            Dim boundingRect As TRectangle
            boundingRect = .graphObj.BoundingRectangle

            calculateObjectHighAndLowInInterval index, _
                                            rectGetXInterval(mCanvas.Boundary), _
                                            objMax, _
                                            objMin
            If objMax > scaleMax Then
                scaleMax = objMax
            End If
            If objMin < scaleMin Then
                scaleMin = objMin
            End If
        End If

        index = .nextInScopeIndex
    End With
Loop
End Sub

Friend Function CalcVerticalBoundsForNonGaugeDependentObjects( _
                ByRef scaleMin As Double, _
                ByRef scaleMax As Double) As Boolean
Dim objMax As Double
Dim objMin As Double
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex

Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .graphObj.IncludeInAutoscale _
        Then
            If .graphObj.GaugeDependent Then
                CalcVerticalBoundsForNonGaugeDependentObjects = True
            Else

                calculateObjectHighAndLowInInterval index, _
                                                rectGetXInterval(mCanvas.Boundary), _
                                                objMax, _
                                                objMin
                If objMax > scaleMax Then
                    scaleMax = objMax
                End If
                If objMin < scaleMin Then
                    scaleMin = objMin
                End If
            End If
        End If

        index = .nextInScopeIndex
    End With
Loop
End Function

Friend Sub Clear()

Dim failpoint As Long
On Error GoTo Err

Set mChartRef = Nothing
Set mRegionRef = Nothing
StartTask Me, PriorityLow

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "clearRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub DetermineInScopeObjects( _
                ByVal minPeriod As Long, _
                ByVal maxPeriod As Long)
Dim i As Long
Dim index As Long
Dim et As New ElapsedTimer
Dim numberExamined As Long
Dim numberExtended As Long
Dim numberInScope As Long

et.StartTiming

clearInScopeObjectList

For i = minPeriod To maxPeriod
    index = mGraphicObjects(mPeriodTable(i).firstIndex).nextPeriodIndex
    Do While index <> mPeriodTable(i).lastindex
        numberExamined = numberExamined + 1
        With mGraphicObjects(index)
            If .graphObj.BoundingRectangle.isValid Then
                addEntryToInScopeList index
                numberInScope = numberInScope + 1
            End If
            index = .nextPeriodIndex
        End With
    Loop
Next

Debug.Print "DetermineInScopeObjects: nonextended=" & numberExamined & " time=" & 1000 * (Timer - timerIn) & "ms"
timerIn = Timer

index = mGraphicObjects(mFirstExtendedIndex).nextExtendedIndex
Do While index <> mLastExtendedIndex
    numberExamined = numberExamined + 1
    numberExtended = numberExtended + 1
    With mGraphicObjects(index)
        Dim boundingRect As TRectangle
        boundingRect = .graphObj.BoundingRectangle
        If graphicObjectIsInScope(index) Then
            addEntryToInScopeList index
            numberInScope = numberInScope + 1
        End If
        index = .nextExtendedIndex
    End With
Loop

Debug.Print "DetermineInScopeObjects: extended=" & numberExtended & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"
Debug.Print "DetermineInScopeObjects: total=" & numberExamined & " in-scope=" & numberInScope
End Sub

Friend Sub DetermineVisibleObjects()
Dim index As Long
Dim et As New ElapsedTimer
Dim numberExamined As Long
Dim numberVisible As Long

et.StartTiming

clearVisibleObjectList

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    numberExamined = numberExamined + 1
    With mGraphicObjects(index)
        Dim boundingRect As TRectangle
        boundingRect = .graphObj.BoundingRectangle
        If graphicObjectIsVisible(index) Then
            addEntryToVisibleList index
            numberVisible = numberVisible + 1
        End If
        index = .nextInScopeIndex
    End With
Loop

Debug.Print "DetermineVisibleObjects: total=" & numberExamined & " visible=" & numberVisible & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"
End Sub

Friend Function GetGraphicObject( _
                ByVal handle As Long) As IGraphicObject
Set GetGraphicObject = mGraphicObjects(handle).graphObj
End Function

Friend Function GetNextUserLayer(Optional ByVal Layer As Long = -1) As Long
Static lNextLayer As Long

If lNextLayer = 0 Then lNextLayer = LayerNumbers.LayerLowestUser

If Layer = -1 Then
    GetLayer = lNextLayer
    If lNextLayer <> LayerNumbers.LayerHighestUser Then
        lNextLayer = lNextLayer + 1
    End If
Else
    If Layer < MinLayer Or Layer > MaxLayer Then
        GetLayer = GetLayer()
    Else
        GetLayer = Layer
    End If
End If
End Function

Friend Sub Initialise( _
                ByVal pChart As Chart, _
                ByVal pRegion As ChartRegion)
Set mChartRef = CreateWeakReference(pChart)
Set mRegionRef = CreateWeakReference(pRegion)

End Sub

Friend Sub PaintAllObjects()
Dim i As Long
Dim counter As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

Dim drawAreas(0) As TRectangle
drawAreas(0) = .Boundary
drawAreas(0).Right = drawAreas(0).Right + 1
For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            .graphObj.Draw drawAreas
            index = .nextVisibleIndex
            counter = counter + 1
        End With
    Loop
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "PaintAllObjects" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription

End Sub

Friend Sub RecalcBoundaryForObjectsInSeries( _
                ByVal SeriesID As Long)
Dim i As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            If .graphObj.SeriesID = SeriesID Then
                Dim boundingRect As TRectangle
                ' force recalculation of the bounding rect
                boundingRect = .graphObj.BoundingRectangle
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_recalcBoundary" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription


End Sub

Friend Sub RedrawArea( _
                ByRef areaToRedraw As TRectangle, _
                ByVal startLayer As Long, _
                Optional ByVal objToExclude As IGraphicObject)
Dim Layer As Long
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

' redraw objects in the undrawn area, preserving layering

Dim rects(0) As TRectangle

rects(0) = areaToRedraw

For Layer = startLayer To MaxLayer

    failpoint = 400

    index = mGraphicObjects(mVisibleTable(Layer).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(Layer).lastindex
        With mGraphicObjects(index)

            failpoint = 500

            If Not .graphObj Is objToExclude Then
                If rectOverlaps(.graphObj.BoundingRectangle, areaToRedraw) Then

                    failpoint = 600

                    .graphObj.Draw rects
                End If
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next


Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "redrawArea" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription
End Sub

Friend Sub RemoveGraphicObject( _
                ByVal graphObj As IGraphicObject)
Dim handle As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

handle = graphObj.handle

If Not graphObj.ExtendedObject Then

    failpoint = 400

    removeEntryFromPeriodList handle
Else

    failpoint = 500

    removeEntryFromExtendedList handle
End If

If graphicObjectIsInScope(handle) Then

failpoint = 600

    removeEntryFromInScopeList handle
    If graphicObjectIsVisible(handle) Then

        failpoint = 700

        removeEntryFromVisibleList handle
    End If
End If

failpoint = 800

Set mGraphicObjects(handle).graphObj = Nothing
graphObj.handle = -1

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RemoveGraphicObject" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

' returns true if the object is in scope
Friend Function UpdateGraphicObject( _
                ByVal graphObj As IGraphicObject) As Boolean
Dim handle As Long

Dim failpoint As Long
On Error GoTo Err

handle = graphObj.handle

If Not graphObj.BoundingRectangle.isValid Then
    removeEntryFromInScopeList handle
    removeEntryFromVisibleList handle
End If

With mGraphicObjects(handle)
    
    ' get the new properties for the changed object
    
    If .Layer <> graphObj.Layer Then removeEntryFromVisibleList handle
    .Layer = graphObj.Layer
End With

If graphObj.ExtendedObject Then
    removeEntryFromPeriodList handle
    addEntryToExtendedList handle
Else
    removeEntryFromExtendedList handle
    
    ' The following call to addEntryToPeriodList must not be in a With block
    ' involving mGraphicObjects because in rare circumstances it can result
    ' in mGraphicObjects being ReDim'd which leads to an error 10
    addEntryToPeriodList handle
End If

If graphicObjectIsInScope(handle) Then
    addEntryToInScopeList handle
    If graphicObjectIsVisible(handle) Then
        addEntryToVisibleList handle
    Else
        removeEntryFromVisibleList handle
    End If
Else
    removeEntryFromInScopeList handle
    UpdateGraphicObject = False
    Exit Function
End If

UpdateGraphicObject = True
    
Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "GOSP_objectChanged" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription



End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub addEntryToExtendedList(ByVal index As Long)
With mGraphicObjects(index)
    If .nextExtendedIndex <> 0 Then Exit Sub
    .nextExtendedIndex = mLastExtendedIndex
    .prevExtendedIndex = mGraphicObjects(mLastExtendedIndex).prevExtendedIndex
    mGraphicObjects(.prevExtendedIndex).nextExtendedIndex = index
    mGraphicObjects(mLastExtendedIndex).prevExtendedIndex = index
End With
End Sub

Private Sub addEntryToInScopeList(ByVal index As Long)
With mGraphicObjects(index)
    If .inScopeCycle <> mInScopeCycle Then
        .graphObj.InScope = True
        .inScopeCycle = mInScopeCycle
        .nextInScopeIndex = mLastInScopeIndex
        .prevInScopeIndex = mGraphicObjects(mLastInScopeIndex).prevInScopeIndex
        mGraphicObjects(.prevInScopeIndex).nextInScopeIndex = index
        mGraphicObjects(mLastInScopeIndex).prevInScopeIndex = index
    End If
End With
End Sub

Private Sub addEntryToPeriodList( _
                ByVal index As Long)
Dim lastindex As Long
Dim newPeriodNumber As Long

newPeriodNumber = mGraphicObjects(index).graphObj.PeriodNumber
If newPeriodNumber <> mGraphicObjects(index).PeriodNumber Then
    If mGraphicObjects(index).PeriodNumber >= 1 Then
        removeEntryFromPeriodList index
    End If
    If newPeriodNumber < 1 Then Exit Sub

    ' NB: the following call to AddPeriod must not be within a 'With' block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    AddPeriod newPeriodNumber, _
            mGraphicObjects(index).graphObj.Timestamp ' in case the period hasn't been added yet

    lastindex = mPeriodTable(newPeriodNumber).lastindex
    With mGraphicObjects(index)
        .PeriodNumber = newPeriodNumber
        .nextPeriodIndex = lastindex
        .prevPeriodIndex = mGraphicObjects(lastindex).prevPeriodIndex
        mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = index
        mGraphicObjects(lastindex).prevPeriodIndex = index
    End With
End If

End Sub

Private Sub addEntryToVisibleList(ByVal index As Long)
Dim lastindex As Long
With mGraphicObjects(index)
    If .visibleCycle <> mVisibleCycle Then
        .graphObj.Visible = True
        .visibleCycle = mVisibleCycle
        lastindex = mVisibleTable(.Layer).lastindex
        .nextVisibleIndex = lastindex
        .prevVisibleIndex = mGraphicObjects(lastindex).prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = index
        mGraphicObjects(lastindex).prevVisibleIndex = index
    End If
End With
End Sub

Private Sub addNewPeriod(ByVal PeriodNumber As Long)
Dim i As Long

If PeriodNumber > UBound(mPeriodTable) Then
    ReDim Preserve mPeriodTable(PeriodNumber + 200) As PeriodTableEntry
End If
If PeriodNumber > mMaxPeriodNumber Then
    For i = mMaxPeriodNumber + 1 To PeriodNumber
        initialisePeriodTableEntry i
    Next
    mMaxPeriodNumber = PeriodNumber
Else
    initialisePeriodTableEntry PeriodNumber
End If

End Sub

Private Sub addVerticalGridLine(ByVal PeriodNumber As Long, _
                                ByVal Key As String, _
                                ByVal label As String, _
                                ByVal coarseKey As String, _
                                ByVal coarseLabel As String)
Dim vGridLine As ChartSkil26.Line
Dim vGridText As Text
Dim vGridTextCoarse As Text

If HasGrid Then
    On Error Resume Next
    Set vGridLine = mXGridLines.Item(Key)
    On Error GoTo 0
    If vGridLine Is Nothing Then
        Set vGridLine = mXGridLines.Add(Key)
        vGridLine.SetPosition NewPoint(PeriodNumber - 0.5, _
                                    0, _
                                    CoordsLogical, _
                                    CoordsRelative), _
                                NewPoint(PeriodNumber - 0.5, _
                                    100, _
                                    CoordsLogical, _
                                    CoordsRelative)
    Else
        If vGridLine.Point1.XLogical > PeriodNumber Then
            vGridLine.SetPosition NewPoint(PeriodNumber - 0.5, _
                                        0, _
                                        CoordsLogical, _
                                        CoordsRelative), _
                                    NewPoint(PeriodNumber - 0.5, _
                                        100, _
                                        CoordsLogical, _
                                        CoordsRelative)
        End If
    End If
End If

If HasGridText Then
    On Error Resume Next
    Set vGridText = mXGridTexts.Item(Key)
    On Error GoTo 0
    If vGridText Is Nothing Then
        Set vGridText = mXGridTexts.Add(Key)
        vGridText.FixedX = False
        vGridText.FixedY = True
        vGridText.position = NewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsDistance)
        vGridText.Text = label
    End If

    On Error Resume Next
    Set vGridTextCoarse = mXGridTextsCoarse.Item(coarseKey)
    On Error GoTo 0
    If vGridTextCoarse Is Nothing Then
        Set vGridTextCoarse = mXGridTextsCoarse.Add(coarseKey)
        vGridTextCoarse.FixedX = False
        vGridTextCoarse.FixedY = True
        vGridTextCoarse.position = NewPoint(PeriodNumber - 0.5, _
                                    0#, _
                                    CoordsLogical, _
                                    CoordsCounterDistance)
        vGridTextCoarse.Text = coarseLabel
    End If
End If

End Sub

Private Function allocateGraphicObjectTableEntry() As Long
On Error GoTo Err
allocateGraphicObjectTableEntry = mNextGraphicObjectIndex

If mNextGraphicObjectIndex > UBound(mGraphicObjects) Then
    ReDim Preserve mGraphicObjects(UBound(mGraphicObjects) + 500) As GraphicObjectTableEntry
End If
mNextGraphicObjectIndex = mNextGraphicObjectIndex + 1
Exit Function
Err:
Stop
End Function

Private Sub calculateObjectHighAndLowInInterval( _
                            ByVal index As Long, _
                            ByRef interval As TInterval, _
                            ByRef high As Double, _
                            ByRef low As Double)

Dim j As Long
Dim k As Long
Dim rects() As TRectangle
Dim numrects As Long
Dim overlapXInterval As TInterval

high = MinusInfinityDouble
low = PlusInfinityDouble

With mGraphicObjects(index)
    Dim boundingRect As TRectangle
    boundingRect = .graphObj.BoundingRectangle
    If (.graphObj.Capabilities And GraphicObjectCapabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(interval, rectGetXInterval(boundingRect))
        For j = Int(overlapXInterval.startValue + 0.9) To _
                Int(overlapXInterval.endValue)
            rects = .graphObj.BoundingRectanglesAt(j)
            numrects = -1
            On Error Resume Next
            numrects = UBound(rects)
            On Error GoTo 0
            If numrects >= 0 Then
                For k = 0 To numrects
                    If rects(k).Top > high Then
                        high = rects(k).Top
                    End If
                    If rects(k).Bottom < low Then
                        low = rects(k).Bottom
                    End If
                Next
            End If
        Next
    Else
        high = boundingRect.Top
        low = boundingRect.Bottom
    End If
End With
End Sub

Private Sub calculateVerticalBounds()
Dim scaleMax As Double
Dim scaleMin As Double

Dim failpoint As Long
On Error GoTo Err

If Not Autoscaling Then Exit Sub

scaleMax = MinusInfinityDouble
scaleMin = PlusInfinityDouble

If CalcVerticalBoundsForNonGaugeDependentObjects(scaleMin, scaleMax) Then
    ' there are gauge-dependent objects that may affect the max and min scales.
    ' We need to recalculate their bounds with the new gauge determined from
    ' the other objects to get a better estimate of their sizes.
    setCanvasVerticalBounds scaleMin, scaleMax
    CalcVerticalBoundsForGaugeDependentObjects scaleMin, scaleMax
End If

If scaleMin = PlusInfinityDouble Or scaleMax = MinusInfinityDouble Then Exit Sub

setCanvasVerticalBounds scaleMin - (scaleMax - scaleMin) * 0.025, _
                        scaleMax + (scaleMax - scaleMin) * 0.025

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "calculateVerticalBounds" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource & vbCrLf & ".Bottom=" & CStr(scaleMin - (scaleMax - scaleMin) * 0.025)
Err.Raise errNumber, errSource, errDescription
End Sub

Private Function chartObj() As Chart
Set chartObj = mChartRef.Target
End Function

Private Sub clearInScopeObjectList()
Dim index As Long

index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        .graphObj.InScope = False
        index = .nextInScopeIndex
    End With
Loop

initialiseInScopeObjectList
End Sub

Private Sub clearVisibleObjectList()
Dim i As Long
Dim index As Long

For i = MinLayer To MaxLayer
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            .graphObj.Visible = False
            index = .nextVisibleIndex
        End With
    Loop
Next

initialiseVisibleObjectList
End Sub

Private Function getObjectOverlaps(ByVal index1 As Long, _
                                ByVal index2 As Long, _
                                ByRef clippingRect As TRectangle, _
                                ByRef useComplexBounding1 As Boolean, _
                                ByRef useComplexBounding2 As Boolean, _
                                ByRef overlapRects() As TRectangle) As Long
Dim overlaps() As TRectangle
Dim graphobj1 As IGraphicObject
Dim graphobj2 As IGraphicObject
Dim obj1Rect As TRectangle
Dim obj2Rect As TRectangle
Dim XInterval As TInterval
Dim complexBounding1 As Boolean
Dim complexBounding2 As Boolean
Dim boundingRects1() As TRectangle
Dim numRects1 As Long
Dim boundingRects2() As TRectangle
Dim numRects2 As Long
Dim i As Long
Dim j As Long
Dim k As Long

Set graphobj1 = mGraphicObjects(index1).graphObj
obj1Rect = rectIntersection(graphobj1.BoundingRectangle, clippingRect)
complexBounding1 = (graphobj1.Capabilities And GraphicObjectCapabilities.ComplexBounding)

Set graphobj2 = mGraphicObjects(index2).graphObj
obj2Rect = rectIntersection(graphobj2.BoundingRectangle, clippingRect)
complexBounding2 = (graphobj2.Capabilities And GraphicObjectCapabilities.ComplexBounding)

ReDim overlaps(0) As TRectangle
overlaps(0) = rectIntersection(obj1Rect, obj2Rect)
If Not overlaps(0).isValid Then
    getObjectOverlaps = 0
    Exit Function
End If

If Not (useComplexBounding1 Or useComplexBounding2) Or _
    Not (complexBounding1 Or complexBounding2) Or _
    (complexBounding1 And Not useComplexBounding1) Or _
    (complexBounding2 And Not useComplexBounding2) _
Then
    overlapRects = overlaps
    getObjectOverlaps = 1
    Exit Function
End If

XInterval = rectGetXInterval(rectIntersection(obj1Rect, obj2Rect))

If complexBounding1 And complexBounding2 And useComplexBounding1 And useComplexBounding2 Then
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be Extended
            
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects1 = graphobj1.BoundingRectanglesAt(i)
        numRects1 = -1
        On Error Resume Next
        numRects1 = UBound(boundingRects1) + 1
        On Error GoTo 0
        
        boundingRects2 = graphobj2.BoundingRectanglesAt(i)
        numRects2 = -1
        On Error Resume Next
        numRects2 = UBound(boundingRects2) + 1
        On Error GoTo 0
        
        If numRects1 > 0 And numRects2 > 0 Then
            For j = 0 To UBound(boundingRects1)
                If boundingRects1(j).isValid Then
                    For k = 0 To UBound(boundingRects2)
                        If boundingRects2(k).isValid Then
                            If getObjectOverlaps > UBound(overlaps) Then
                                ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                            End If
                            overlaps(getObjectOverlaps) = _
                                rectIntersection(boundingRects1(j), boundingRects2(k))
                            If overlaps(getObjectOverlaps).isValid Then
                                getObjectOverlaps = getObjectOverlaps + 1
                            End If
                        End If
                    Next
                End If
            Next
        End If
    If getObjectOverlaps > 0 Then
        ReDim Preserve overlaps(getObjectOverlaps - 1) As TRectangle
        overlapRects = overlaps
    End If
    Next
ElseIf complexBounding1 Then
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be Extended
    
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects1 = graphobj1.BoundingRectanglesAt(i)
        numRects1 = -1
        On Error Resume Next
        numRects1 = UBound(boundingRects1) + 1
        On Error GoTo 0
        
        If numRects1 > 0 Then
            For j = 0 To UBound(boundingRects1)
                If getObjectOverlaps > UBound(overlaps) Then
                    ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                End If
                overlaps(getObjectOverlaps) = rectIntersection(boundingRects1(j), obj2Rect)
                If overlaps(getObjectOverlaps).isValid Then
                    getObjectOverlaps = getObjectOverlaps + 1
                End If
            Next
        End If
    Next
    If getObjectOverlaps > 0 Then
        ReDim Preserve overlaps(getObjectOverlaps - 1) As TRectangle
        overlapRects = overlaps
    End If
Else
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be Extended
    
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects2 = graphobj2.BoundingRectanglesAt(i)
        numRects2 = -1
        On Error Resume Next
        numRects2 = UBound(boundingRects2) + 1
        On Error GoTo 0
        
        If numRects2 > 0 Then
            For j = 0 To UBound(boundingRects2)
                If getObjectOverlaps > UBound(overlaps) Then
                    ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                End If
                overlaps(getObjectOverlaps) = rectIntersection(boundingRects2(j), obj1Rect)
                If overlaps(getObjectOverlaps).isValid Then
                    getObjectOverlaps = getObjectOverlaps + 1
                End If
            Next
        End If
    Next
    If getObjectOverlaps > 0 Then
        ReDim Preserve overlaps(getObjectOverlaps - 1) As TRectangle
        overlapRects = overlaps
    End If
End If

End Function

Private Function graphicObjectIsInScope(ByVal index As Long) As Boolean
With mGraphicObjects(index)
    If Not .graphObj.BoundingRectangle.isValid Then Exit Function
    If .graphObj.ExtendedObject Then
        graphicObjectIsInScope = intOverlaps(rectGetXInterval(mCanvas.Boundary), rectGetXInterval(.graphObj.BoundingRectangle))
    Else
        graphicObjectIsInScope = (.PeriodNumber >= mMinInViewPeriod And _
                                .PeriodNumber <= mMaxInViewPeriod)
    End If
End With
End Function

Private Function graphicObjectIsVisible(ByVal index As Long) As Boolean
Dim boundingRects() As TRectangle
Dim numrects As Long
Dim overlapXInterval As TInterval
Dim i As Long
Dim j As Long

With mGraphicObjects(index).graphObj
    If Not .BoundingRectangle.isValid Or .NoDraw Then Exit Function
    If .ExtendedObject And (.Capabilities And GraphicObjectCapabilities.ComplexBounding) Then
        overlapXInterval = intIntersection(rectGetXInterval(mCanvas.Boundary), rectGetXInterval(.BoundingRectangle))
        For i = Int(overlapXInterval.startValue) To _
                Int(overlapXInterval.endValue)
            boundingRects = .BoundingRectanglesAt(i)
            numrects = -1
            On Error Resume Next
            numrects = UBound(boundingRects)
            On Error GoTo 0
            If numrects >= 0 Then
                For j = 0 To numrects
                    If rectOverlaps(boundingRects(j), mCanvas.Boundary) Then
                        graphicObjectIsVisible = True
                        Exit Function
                    End If
                Next
            End If
        Next
    Else
        graphicObjectIsVisible = rectOverlaps(.BoundingRectangle, mCanvas.Boundary)
    End If
End With
End Function

Private Sub initialiseInScopeObjectList()
mInScopeCycle = mInScopeCycle + 1
mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex = mLastInScopeIndex
mGraphicObjects(mLastInScopeIndex).prevInScopeIndex = mFirstInScopeIndex
End Sub

Private Sub initialisePeriodTableEntry( _
                ByVal index As Long)
'With mPeriodTable(index)
    If mPeriodTable(index).firstIndex = 0 Then
        mPeriodTable(index).firstIndex = allocateGraphicObjectTableEntry
        mPeriodTable(index).lastindex = allocateGraphicObjectTableEntry
        mGraphicObjects(mPeriodTable(index).firstIndex).nextPeriodIndex = mPeriodTable(index).lastindex
        mGraphicObjects(mPeriodTable(index).lastindex).prevPeriodIndex = mPeriodTable(index).firstIndex
    End If
'End With
End Sub

Private Sub initialiseVisibleObjectList()
Dim i As Long

mVisibleCycle = mVisibleCycle + 1

For i = MinLayer To MaxLayer
    mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex = mVisibleTable(i).lastindex
    mGraphicObjects(mVisibleTable(i).lastindex).prevVisibleIndex = mVisibleTable(i).firstIndex
Next

End Sub

Private Sub removeEntryFromExtendedList(ByVal index As Long)
With mGraphicObjects(index)
    If .nextExtendedIndex = 0 Then Exit Sub
    mGraphicObjects(.nextExtendedIndex).prevExtendedIndex = .prevExtendedIndex
    mGraphicObjects(.prevExtendedIndex).nextExtendedIndex = .nextExtendedIndex
    .nextExtendedIndex = 0
    .prevExtendedIndex = 0
End With
End Sub

Private Sub removeEntryFromInScopeList(ByVal index As Long)
With mGraphicObjects(index)
    If .inScopeCycle = mInScopeCycle Then
        .inScopeCycle = 0
        mGraphicObjects(.nextInScopeIndex).prevInScopeIndex = .prevInScopeIndex
        mGraphicObjects(.prevInScopeIndex).nextInScopeIndex = .nextInScopeIndex
        .nextInScopeIndex = 0
        .prevInScopeIndex = 0
    End If
End With
End Sub

Private Sub removeEntryFromPeriodList(ByVal index As Long)
With mGraphicObjects(index)
    If .nextPeriodIndex = 0 Then Exit Sub
    mGraphicObjects(.nextPeriodIndex).prevPeriodIndex = .prevPeriodIndex
    mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = .nextPeriodIndex
    .nextPeriodIndex = 0
    .prevPeriodIndex = 0
End With
End Sub

Private Sub removeEntryFromVisibleList(ByVal index As Long)
With mGraphicObjects(index)
    If .visibleCycle = mVisibleCycle Then
        .visibleCycle = 0
        mGraphicObjects(.nextVisibleIndex).prevVisibleIndex = .prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = .nextVisibleIndex
        .nextVisibleIndex = 0
        .prevVisibleIndex = 0
    End If
End With
End Sub

Private Sub undrawCrosshairsPointer()

If Not mIsDrawingEnabled Then Exit Sub

mCrosshairLineHoriz.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)
mCrosshairLineVert.SetPosition NewPoint(-1000, 0), _
                                NewPoint(-1000, 0)

End Sub




