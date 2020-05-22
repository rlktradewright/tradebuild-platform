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

Implements ITask
Implements ITaskCompletionListener

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
    Redraw              As Byte
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

Private mPeriodTable()                  As PeriodTableEntry

Private mGraphicObjects()               As GraphicObjectTableEntry
Private mFirstGraphicObjectIndex        As Long ' the index of the first table entry
                                                ' that actually points to a graphic
                                                ' object - all prior entries are
                                                ' merely list start and end entries
Private mNextGraphicObjectIndex         As Long

Private mNumberOfObjects                As Long

Private mGraphObjSeriesCollection       As GraphObjSeriesCollection

Private mMaxPeriodNumber                As Long

Private mViewport                       As ViewPort

Private mInScopeCycle                   As Long
Private mFirstInScopeIndex              As Long
Private mLastInScopeIndex               As Long

Private mVisibleCycle                   As Long
Private mVisibleTable(LayerMin To LayerMax) As VisibleTableEntry

Private mFirstExtendedIndex             As Long
Private mLastExtendedIndex              As Long

Private mTaskContext                    As TaskContext

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

For i = LayerMin To LayerMax
    mVisibleTable(i).firstIndex = allocateGraphicObjectTableEntry
    mVisibleTable(i).lastindex = allocateGraphicObjectTableEntry
Next
initialiseVisibleObjectList

End Sub

'Private Sub Class_Terminate()
'Debug.Print "ChartRegionData terminated"
'End Sub

'================================================================================
' ITask Interface Members
'================================================================================

Private Sub ITask_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

Static sTotalNumberReleased As Long
Static sYieldCount As Long
Static sCurrentIndex As Long
Static sClearStartTime As Date

Dim results As Parameters

If sClearStartTime = 0 Then sClearStartTime = GetTimestamp

If mTaskContext.CancelPending Then
    gLogger.Log "Released (cancelled) " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    Set results = generateResults(False, sTotalNumberReleased, sYieldCount, sClearStartTime)
    Erase mGraphicObjects
    mTaskContext.Finish results, True
    Exit Sub
End If

Static sNumberToRelease As Long

Static sLoopCount As Long

Const HeuristicNumber As Long = 50
Const HeuristicTime As Long = 5

If sLoopCount Mod 50 = 0 Then
    sNumberToRelease = calibrate(HeuristicNumber, HeuristicTime, sCurrentIndex, sTotalNumberReleased, sYieldCount)
ElseIf Not releaseEntries(sNumberToRelease, sCurrentIndex, sTotalNumberReleased, sYieldCount) Then
    sNumberToRelease = 0
End If

If sNumberToRelease = 0 Then
    gLogger.Log "Released " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    Set results = generateResults(False, sTotalNumberReleased, sYieldCount, sClearStartTime)
    Erase mGraphicObjects
    mTaskContext.Finish results, False
End If

sLoopCount = sLoopCount + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'================================================================================
' ITaskCompletionListener Interface Members
'================================================================================

Private Sub ITaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim params As Parameters
Set params = ev.Result

If Not gLogger.IsLoggable(LogLevelHighDetail) Then Exit Sub

ReDim s(params.Count - 1) As String
Dim i As Long
For i = 0 To UBound(s)
    Dim p As Parameter
    Set p = params.Item(i + 1)
    s(i) = p.Name & "=" & p.Value
Next
gLogger.Log "Data cleared: " & Join(s, "; "), ModuleName, ProcName, LogLevelHighDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Get MaxPeriodNumber() As Long
MaxPeriodNumber = mMaxPeriodNumber
End Property

Friend Property Get MinPeriodNumber() As Long
MinPeriodNumber = LBound(mPeriodTable)
End Property

Friend Property Get NumberOfObjects() As Long
NumberOfObjects = mNumberOfObjects
End Property

Friend Property Let PeriodBackColor( _
                ByVal PeriodNumber As Long, _
                ByVal Value As Long)
Const ProcName As String = "PeriodBackColor"
On Error GoTo Err

mPeriodTable(PeriodNumber).BackColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get PeriodBackColor( _
                ByVal PeriodNumber As Long) As Long
Const ProcName As String = "PeriodBackColor"
On Error GoTo Err

PeriodBackColor = mPeriodTable(PeriodNumber).BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Friend Function AddGraphicObject( _
                ByVal graphObj As IGraphicObject) As Long
Const ProcName As String = "AddGraphicObject"
On Error GoTo Err

Dim Handle As Long
Handle = allocateGraphicObjectTableEntry
If mFirstGraphicObjectIndex = 0 Then
    mFirstGraphicObjectIndex = Handle
End If

graphObj.Handle = Handle

Set mGraphicObjects(Handle).graphObj = graphObj

If graphObj.ExtendedObject Then
    addEntryToExtendedList Handle
Else
    ' NB: the following call to addEntryToPeriodList must not be within a With block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    addEntryToPeriodList Handle
End If

If graphicObjectIsInScope(Handle) Then
    addEntryToInScopeList Handle
    If graphicObjectIsVisible(Handle) Then
        addEntryToVisibleList Handle
    End If
End If
mNumberOfObjects = mNumberOfObjects + 1
AddGraphicObject = Handle

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub AddPeriod( _
                ByVal pPeriodNumber As Long)
Const ProcName As String = "AddPeriod"
On Error GoTo Err

addNewPeriod pPeriodNumber

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function CalculateInscopeGaugeDependentObjects()
Const ProcName As String = "CalculateInscopeGaugeDependentObjects"
On Error GoTo Err

Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
     With mGraphicObjects(index)
        If .graphObj.GaugeDependent Then
            'Dim boundingRect As TRectangle
            'boundingRect = .graphObj.BoundingRectangle
            If graphicObjectIsVisible(index) Then
                addEntryToVisibleList index
            End If
        End If
        index = .nextInScopeIndex
    End With
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Sub CalcVerticalBoundsForGaugeDependentObjects( _
                ByRef scaleMin As Double, _
                ByRef scaleMax As Double)
Const ProcName As String = "CalcVerticalBoundsForGaugeDependentObjects"
On Error GoTo Err

Static et As New ElapsedTimer
et.StartTiming

Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex

Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        If Not .graphObj.NoDraw And _
            .graphObj.IncludeInAutoscale And _
            .graphObj.GaugeDependent _
        Then
            'Dim boundingRect As TRectangle
            'boundingRect = .graphObj.BoundingRectangle

            Dim objMax As Double
            Dim objMin As Double
            calculateObjectHighAndLowInPeriodRange index, _
                                            mViewport.Boundary.Left, _
                                            mViewport.Boundary.Right, _
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

Debug.Print "CalcVerticalBoundsForGaugeDependentObjects: time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function CalcVerticalBoundsForNonGaugeDependentObjects( _
                ByRef scaleMin As Double, _
                ByRef scaleMax As Double) As Boolean
Const ProcName As String = "CalcVerticalBoundsForNonGaugeDependentObjects"
On Error GoTo Err

Static et As New ElapsedTimer
et.StartTiming

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

                Dim objMax As Double
                Dim objMin As Double
                calculateObjectHighAndLowInPeriodRange index, _
                                                        mViewport.Boundary.Left, _
                                                        mViewport.Boundary.Right, _
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

Debug.Print "CalcVerticalBoundsForNonGaugeDependentObjects: time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

StartTask(Me, PriorityLow).AddTaskCompletionListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DetermineInScopeObjects()
Const ProcName As String = "DetermineInScopeObjects"
On Error GoTo Err

Dim et As New ElapsedTimer
et.StartTiming

clearInScopeObjectList

Dim numberInScope As Long
Dim numberExamined As Long
Dim numberExtended As Long

Dim i As Long
For i = Int(IIf(mViewport.Boundary.Left < LBound(mPeriodTable), _
                LBound(mPeriodTable), _
                Int(mViewport.Boundary.Left))) _
    To Int(IIf(mViewport.Boundary.Right > mMaxPeriodNumber, _
                mMaxPeriodNumber, _
                Int(mViewport.Boundary.Right)))
    
    Dim lNextPeriodIndex As Long
    lNextPeriodIndex = mGraphicObjects(mPeriodTable(i).firstIndex).nextPeriodIndex
    Do While lNextPeriodIndex <> mPeriodTable(i).lastindex
        numberExamined = numberExamined + 1
        With mGraphicObjects(lNextPeriodIndex)
            If graphicObjectIsInScope(lNextPeriodIndex) Then
                addEntryToInScopeList lNextPeriodIndex
                numberInScope = numberInScope + 1
            End If
            lNextPeriodIndex = .nextPeriodIndex
        End With
    Loop
Next

Debug.Print "DetermineInScopeObjects: nonextended=" & numberExamined & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

et.StartTiming

Dim lNextExtendedIndex As Long
lNextExtendedIndex = mGraphicObjects(mFirstExtendedIndex).nextExtendedIndex
Do While lNextExtendedIndex <> mLastExtendedIndex
    numberExamined = numberExamined + 1
    With mGraphicObjects(lNextExtendedIndex)
        If graphicObjectIsInScope(lNextExtendedIndex) Then
            addEntryToInScopeList lNextExtendedIndex
            numberInScope = numberInScope + 1
        End If
        lNextExtendedIndex = .nextExtendedIndex
    End With
Loop

Debug.Print "DetermineInScopeObjects: extended=" & numberExtended & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"
Debug.Print "DetermineInScopeObjects: total=" & numberExamined & " in-scope=" & numberInScope

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DetermineVisibleObjects()
Const ProcName As String = "DetermineVisibleObjects"
On Error GoTo Err

Dim et As New ElapsedTimer
et.StartTiming

clearVisibleObjectList

Dim numberExamined As Long
Dim numberVisible As Long

Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    numberExamined = numberExamined + 1
    With mGraphicObjects(index)
        'Dim boundingRect As TRectangle
        'boundingRect = .graphObj.BoundingRectangle
        If graphicObjectIsVisible(index) Then
            addEntryToVisibleList index
            numberVisible = numberVisible + 1
        End If
        index = .nextInScopeIndex
    End With
Loop

Debug.Print "DetermineVisibleObjects: total=" & numberExamined & " visible=" & numberVisible & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetGraphicObject( _
                ByVal Handle As Long) As IGraphicObject
Const ProcName As String = "GetGraphicObject"
On Error GoTo Err

Set GetGraphicObject = mGraphicObjects(Handle).graphObj

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetNextSelectableObjectOrSeriesAtPoint( _
                ByVal currObject As IGraphicObject, _
                ByVal X As Single, _
                ByVal Y As Single) As IGraphicObject
Const ProcName As String = "GetNextSelectableObjectOrSeriesAtPoint"
On Error GoTo Err

Const NumberOfLayers As Long = LayerNumbers.LayerMax - LayerNumbers.LayerMinVisible - 1

Dim startLayer As Long
If currObject Is Nothing Then
    startLayer = LayerMax
Else
    startLayer = currObject.Layer
End If

Dim currObjStillSelectable As Boolean

Dim i As Long
For i = 1 To NumberOfLayers
    Dim currLayer As Long
    currLayer = startLayer - i
    If currLayer < LayerMinVisible Then currLayer = currLayer + NumberOfLayers
    
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(currLayer).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(currLayer).lastindex
        With mGraphicObjects(index)
            If isObjectOrSeriesSelectable(.graphObj) Then
                If .graphObj.HitTest(X, Y) Then
                    If .graphObj Is currObject Then
                        currObjStillSelectable = True
                    Else
                        Set GetNextSelectableObjectOrSeriesAtPoint = .graphObj
                        ' we've found an object/series to select
                        
                        ' NB we can't use Exit For, because we are
                        ' within a With statement, and it causes the table lock
                        ' not to be released - see http://support.microsoft.com/kb/187553
                        GoTo ExitFor
                        
                    End If
                End If
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

ExitFor:

If GetNextSelectableObjectOrSeriesAtPoint Is Nothing And _
    currObjStillSelectable _
Then
    Set GetNextSelectableObjectOrSeriesAtPoint = currObject
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetNextUserLayer(Optional ByVal Layer As Long = LayerNone) As Long
Const ProcName As String = "GetNextUserLayer"
On Error GoTo Err

Static sNextLayer As Long

If sNextLayer = 0 Then sNextLayer = LayerNumbers.LayerLowestUser

If Layer < LayerNumbers.LayerLowestUser Or _
    Layer > LayerNumbers.LayerHighestUser _
Then
    GetNextUserLayer = sNextLayer
    If sNextLayer <> LayerNumbers.LayerHighestUser Then
        sNextLayer = sNextLayer + 1
    End If
Else
    GetNextUserLayer = Layer
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pViewport As ViewPort, _
                ByVal pGraphObjSeriesCollection As GraphObjSeriesCollection)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mGraphObjSeriesCollection = pGraphObjSeriesCollection
Set mViewport = pViewport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub MarkObjectsForRedraw( _
                ByRef areasToRedraw() As TRectangle, _
                ByVal pCount As Long)
Const ProcName As String = "MarkObjectsForRedraw"
On Error GoTo Err

' mark objects that overlap the redraw area

Dim Layer As Long
For Layer = LayerMinVisible To LayerMax
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(Layer).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(Layer).lastindex
        With mGraphicObjects(index)
            Dim br As TRectangle
            br = .graphObj.BoundingRectangle
            
            Dim i As Long
            For i = 0 To pCount - 1
                If areasToRedraw(i).isValid Then
                    If RectOverlaps(areasToRedraw(i), br) Then
                        .Redraw = CByte(True)
                        Exit For
                    End If
                End If
            Next
            
            index = .nextVisibleIndex
        End With
    Loop
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function PaintAllVisibleObjects() As Long
Const ProcName As String = "PaintAllVisibleObjects"
On Error GoTo Err

Static et As New ElapsedTimer
et.StartTiming

Dim counter As Long

Dim i As Long
For i = LayerMinVisible To LayerMax
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            Dim drawAreas(0) As TRectangle
            drawAreas(0) = .graphObj.BoundingRectangleInViewport
            .graphObj.Draw drawAreas
            index = .nextVisibleIndex
            counter = counter + 1
            .Redraw = 0
        End With
    Loop
Next

Debug.Print "PaintAllVisibleObjects: objects painted: " & counter & "; time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

PaintAllVisibleObjects = counter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function PaintMarkedObjects() As Long
Const ProcName As String = "PaintMarkedObjects"
On Error GoTo Err

Dim counter As Long

Dim i As Long
For i = LayerMinVisible To LayerMax
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            If .Redraw Then
                Dim drawAreas(0) As TRectangle
                drawAreas(0) = .graphObj.BoundingRectangleInViewport
                .graphObj.Draw drawAreas
                counter = counter + 1
                .Redraw = 0
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

PaintMarkedObjects = counter
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub RecalcBoundaryForObjectsInSeries( _
                ByVal SeriesID As Long)
Const ProcName As String = "RecalcBoundaryForObjectsInSeries"
On Error GoTo Err

Dim i As Long
For i = LayerMinVisible To LayerMax
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            If .graphObj.SeriesID = SeriesID Then
                Dim boundingRect As TRectangle
                ' force recalculation of the bounding GDI_RECT
                boundingRect = .graphObj.BoundingRectangle
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RedrawArea( _
                ByRef areaToRedraw As TRectangle, _
                ByVal startLayer As Long, _
                Optional ByVal objToExclude As IGraphicObject)
Const ProcName As String = "RedrawArea"
On Error GoTo Err

' redraw objects in the undrawn area, preserving layering

Dim rects(0) As TRectangle

Dim Layer As Long
For Layer = startLayer To LayerMax
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(Layer).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(Layer).lastindex
        With mGraphicObjects(index)
            If Not .graphObj Is objToExclude Then
                rects(0) = RectIntersection(areaToRedraw, .graphObj.BoundingRectangle)

                If rects(0).isValid Then
                    .graphObj.Draw rects
                End If
            End If
            index = .nextVisibleIndex
        End With
    Loop
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveGraphicObject( _
                ByVal graphObj As IGraphicObject)
Const ProcName As String = "RemoveGraphicObject"
On Error GoTo Err

Dim Handle As Long
Handle = graphObj.Handle

If Not graphObj.ExtendedObject Then
    removeEntryFromPeriodList Handle
Else
    removeEntryFromExtendedList Handle
End If

If graphicObjectIsInScope(Handle) Then
    removeEntryFromInScopeList Handle
    If graphicObjectIsVisible(Handle) Then
        removeEntryFromVisibleList Handle
    End If
End If

Set mGraphicObjects(Handle).graphObj = Nothing
graphObj.Handle = -1

mNumberOfObjects = mNumberOfObjects - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' returns true if the object is in scope
Friend Function UpdateGraphicObject( _
                ByVal graphObj As IGraphicObject) As Boolean
Const ProcName As String = "UpdateGraphicObject"
On Error GoTo Err

Dim Handle As Long
Handle = graphObj.Handle

If Not graphObj.BoundingRectangle.isValid Then
    removeEntryFromInScopeList Handle
    removeEntryFromVisibleList Handle
End If

removeEntryFromVisibleList Handle

If graphObj.ExtendedObject Then
    removeEntryFromPeriodList Handle
    addEntryToExtendedList Handle
Else
    removeEntryFromExtendedList Handle
    
    ' The following call to addEntryToPeriodList must not be in a With block
    ' involving mGraphicObjects because in rare circumstances it can result
    ' in mGraphicObjects being ReDim'd which leads to an error 10
    addEntryToPeriodList Handle
End If

If graphicObjectIsInScope(Handle) Then
    addEntryToInScopeList Handle
    If graphicObjectIsVisible(Handle) Then addEntryToVisibleList Handle
Else
    removeEntryFromInScopeList Handle
    UpdateGraphicObject = False
    Exit Function
End If

UpdateGraphicObject = True
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub addEntryToExtendedList(ByVal index As Long)
Const ProcName As String = "addEntryToExtendedList"
On Error GoTo Err

With mGraphicObjects(index)
    If .nextExtendedIndex <> 0 Then Exit Sub
    .nextExtendedIndex = mLastExtendedIndex
    .prevExtendedIndex = mGraphicObjects(mLastExtendedIndex).prevExtendedIndex
    mGraphicObjects(.prevExtendedIndex).nextExtendedIndex = index
    mGraphicObjects(mLastExtendedIndex).prevExtendedIndex = index
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addEntryToInScopeList(ByVal index As Long)
Const ProcName As String = "addEntryToInScopeList"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function addEntryToPeriodList( _
                ByVal index As Long) As Boolean
Const ProcName As String = "addEntryToPeriodList"
On Error GoTo Err

Dim newPeriodNumber As Long
newPeriodNumber = mGraphicObjects(index).graphObj.PeriodNumber
If newPeriodNumber <> mGraphicObjects(index).PeriodNumber Then
    If mGraphicObjects(index).PeriodNumber >= LBound(mPeriodTable) Then
        removeEntryFromPeriodList index
    End If
    If newPeriodNumber < LBound(mPeriodTable) Then Exit Function

    ' NB: the following call to AddPeriod must not be within a 'With' block
    ' that refers to mGraphicObjects, because in rare circumstances it can
    ' cause mGraphicObjects to be ReDim'd which causes an error 10
    AddPeriod newPeriodNumber ' in case the period hasn't been added yet

    Dim lastindex As Long
    lastindex = mPeriodTable(newPeriodNumber).lastindex
    With mGraphicObjects(index)
        .PeriodNumber = newPeriodNumber
        .nextPeriodIndex = lastindex
        .prevPeriodIndex = mGraphicObjects(lastindex).prevPeriodIndex
        mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = index
        mGraphicObjects(lastindex).prevPeriodIndex = index
    End With
End If

addEntryToPeriodList = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub addEntryToVisibleList(ByVal index As Long)
Const ProcName As String = "addEntryToVisibleList"
On Error GoTo Err

With mGraphicObjects(index)
    If .visibleCycle <> mVisibleCycle Then
        .graphObj.Visible = True
        .visibleCycle = mVisibleCycle
        Dim lastindex As Long
        lastindex = mVisibleTable(.graphObj.Layer).lastindex
        .nextVisibleIndex = lastindex
        .prevVisibleIndex = mGraphicObjects(lastindex).prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = index
        mGraphicObjects(lastindex).prevVisibleIndex = index
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addNewPeriod(ByVal PeriodNumber As Long)
Const ProcName As String = "addNewPeriod"
On Error GoTo Err

If PeriodNumber > UBound(mPeriodTable) Then
    ReDim Preserve mPeriodTable(PeriodNumber + 200) As PeriodTableEntry
End If
If PeriodNumber > mMaxPeriodNumber Then
    Dim i As Long
    For i = mMaxPeriodNumber + 1 To PeriodNumber
        initialisePeriodTableEntry i
    Next
    mMaxPeriodNumber = PeriodNumber
Else
    initialisePeriodTableEntry PeriodNumber
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function allocateGraphicObjectTableEntry() As Long
Const ProcName As String = "allocateGraphicObjectTableEntry"
On Error GoTo Err

allocateGraphicObjectTableEntry = mNextGraphicObjectIndex

If mNextGraphicObjectIndex > UBound(mGraphicObjects) Then
    ReDim Preserve mGraphicObjects(UBound(mGraphicObjects) + 500) As GraphicObjectTableEntry
End If
mNextGraphicObjectIndex = mNextGraphicObjectIndex + 1

mGraphicObjects(allocateGraphicObjectTableEntry).PeriodNumber = &H80000000
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub calculateObjectHighAndLowInPeriodRange( _
                ByVal index As Long, _
                ByVal minPeriod As Long, _
                ByVal maxPeriod As Long, _
                ByRef high As Double, _
                ByRef low As Double)
Const ProcName As String = "calculateObjectHighAndLowInPeriodRange"
On Error GoTo Err

high = MinusInfinityDouble
low = PlusInfinityDouble

Dim periodRange As TInterval
periodRange.startValue = minPeriod
periodRange.endValue = maxPeriod
periodRange.isValid = True

With mGraphicObjects(index)
    If (.graphObj.Capabilities And GraphicObjectCapabilities.ComplexBounding) Then
        
        Dim overlapXInterval As TInterval
        overlapXInterval = IntIntersection(periodRange, RectGetXInterval(.graphObj.BoundingRectangleInViewport))
        
        Dim j As Long
        For j = Int(overlapXInterval.startValue + 0.9) To _
                Int(overlapXInterval.endValue)
            
            Dim rects() As TRectangle
            rects = .graphObj.BoundingRectanglesAt(j)

            Dim numrects As Long
            numrects = -1
            On Error Resume Next
            numrects = UBound(rects)
            On Error GoTo Err
            If numrects >= 0 Then
                Dim k As Long
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
        Dim verticalRange As TInterval
        verticalRange = .graphObj.VerticalRangeInXinterval(periodRange)
        If verticalRange.isValid Then
            high = verticalRange.endValue
            low = verticalRange.startValue
        End If
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function calibrate( _
                ByVal pNumberToRelease As Long, _
                ByVal pHeuristicTime As Long, _
                ByRef pIndex As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Long
Const ProcName As String = "calibrate"
On Error GoTo Err

Dim lPrevCpuTime As Double: lPrevCpuTime = GetCurrentProcessCpuTime
Dim lAligned As Boolean

Dim lNumberReleased As Long
Do While releaseEntries(pNumberToRelease, pIndex, pTotalNumberReleased, pYieldCount)
    lNumberReleased = lNumberReleased + pNumberToRelease
    Dim lCurrCpuTime As Double: lCurrCpuTime = GetCurrentProcessCpuTime
    Dim lCpuTimeMillisecs As Double: lCpuTimeMillisecs = 1000 * (lCurrCpuTime - lPrevCpuTime)
    If lCpuTimeMillisecs > 0 Then
        If Not lAligned Then
            lNumberReleased = 0
            lPrevCpuTime = lCurrCpuTime
            lAligned = True
        Else
            gLogger.Log "Time to release " & lNumberReleased & " entries: " & lCpuTimeMillisecs, ProcName, ModuleName, LogLevelHighDetail
            calibrate = CLng(pHeuristicTime / lCpuTimeMillisecs * lNumberReleased)
            gLogger.Log "Number of entries to release per yield: " & calibrate, ProcName, ModuleName, LogLevelHighDetail
            Exit Function
        End If
    End If
Loop

calibrate = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub clearInScopeObjectList()
Const ProcName As String = "clearInScopeObjectList"
On Error GoTo Err

Dim index As Long
index = mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex
Do While index <> mLastInScopeIndex
    With mGraphicObjects(index)
        .graphObj.InScope = False
        index = .nextInScopeIndex
    End With
Loop

initialiseInScopeObjectList

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearVisibleObjectList()
Const ProcName As String = "clearVisibleObjectList"
On Error GoTo Err

Dim i As Long
For i = LayerMin To LayerMax
    Dim index As Long
    index = mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex
    Do While index <> mVisibleTable(i).lastindex
        With mGraphicObjects(index)
            .graphObj.Visible = False
            index = .nextVisibleIndex
        End With
    Loop
Next

initialiseVisibleObjectList

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateResults( _
                ByVal pCancelled As Boolean, _
                ByVal pNumberReleased As Long, _
                ByVal pNumberOfYields As Long, _
                ByVal pClearStartTime) As Parameters
Const ProcName As String = "generateResults"
On Error GoTo Err

Dim results As New Parameters
results.SetParameterValue "Cancelled", pCancelled
results.SetParameterValue "Number of entries released", pNumberReleased
results.SetParameterValue "Number of timeslots", pNumberOfYields

Dim lDuration As Double: lDuration = 86400# * (GetTimestamp - pClearStartTime)
If lDuration < 0.001 Then
    results.SetParameterValue "Duration (microsecs)", Format(lDuration * 1000000#, "0")
ElseIf lDuration < 1# Then
    results.SetParameterValue "Duration (millisecs)", Format(lDuration * 1000#, "0.000")
Else
    results.SetParameterValue "Duration (secs)", Format(lDuration, "0.000")
End If

Set generateResults = results

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getObjectOverlaps(ByVal index1 As Long, _
                                ByVal index2 As Long, _
                                ByRef clippingRect As TRectangle, _
                                ByRef useComplexBounding1 As Boolean, _
                                ByRef useComplexBounding2 As Boolean, _
                                ByRef overlapRects() As TRectangle) As Long
Const ProcName As String = "getObjectOverlaps"
On Error GoTo Err

Dim graphobj1 As IGraphicObject
Set graphobj1 = mGraphicObjects(index1).graphObj

Dim obj1Rect As TRectangle
obj1Rect = RectIntersection(graphobj1.BoundingRectangle, clippingRect)

Dim complexBounding1 As Boolean
complexBounding1 = (graphobj1.Capabilities And GraphicObjectCapabilities.ComplexBounding)

Dim graphobj2 As IGraphicObject
Set graphobj2 = mGraphicObjects(index2).graphObj

Dim obj2Rect As TRectangle
obj2Rect = RectIntersection(graphobj2.BoundingRectangle, clippingRect)

Dim complexBounding2 As Boolean
complexBounding2 = (graphobj2.Capabilities And GraphicObjectCapabilities.ComplexBounding)

ReDim overlaps(0) As TRectangle
overlaps(0) = RectIntersection(obj1Rect, obj2Rect)
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

Dim XInterval As TInterval
XInterval = RectGetXInterval(RectIntersection(obj1Rect, obj2Rect))

Dim boundingRects1() As TRectangle
Dim numRects1 As Long

Dim boundingRects2() As TRectangle
Dim numRects2 As Long

Dim i As Long
Dim j As Long
Dim k As Long

If complexBounding1 And complexBounding2 And useComplexBounding1 And useComplexBounding2 Then
    ReDim overlaps(5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
            ' size the array so it is unlikely to need to be Extended
            
    For i = Int(XInterval.startValue) To _
            Int(XInterval.endValue)
        
        boundingRects1 = graphobj1.BoundingRectanglesAt(i)
        numRects1 = -1
        On Error Resume Next
        numRects1 = UBound(boundingRects1) + 1
        On Error GoTo Err
        
        boundingRects2 = graphobj2.BoundingRectanglesAt(i)
        numRects2 = -1
        On Error Resume Next
        numRects2 = UBound(boundingRects2) + 1
        On Error GoTo Err
        
        If numRects1 > 0 And numRects2 > 0 Then
            For j = 0 To UBound(boundingRects1)
                If boundingRects1(j).isValid Then
                    For k = 0 To UBound(boundingRects2)
                        If boundingRects2(k).isValid Then
                            If getObjectOverlaps > UBound(overlaps) Then
                                ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                            End If
                            overlaps(getObjectOverlaps) = _
                                RectIntersection(boundingRects1(j), boundingRects2(k))
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
        On Error GoTo Err
        
        If numRects1 > 0 Then
            For j = 0 To UBound(boundingRects1)
                If getObjectOverlaps > UBound(overlaps) Then
                    ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                End If
                overlaps(getObjectOverlaps) = RectIntersection(boundingRects1(j), obj2Rect)
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
        On Error GoTo Err
        
        If numRects2 > 0 Then
            For j = 0 To UBound(boundingRects2)
                If getObjectOverlaps > UBound(overlaps) Then
                    ReDim Preserve overlaps(UBound(overlaps) + 5 * (Int(XInterval.endValue - XInterval.startValue) + 1)) As TRectangle
                End If
                overlaps(getObjectOverlaps) = RectIntersection(boundingRects2(j), obj1Rect)
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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function graphicObjectIsInScope(ByVal index As Long) As Boolean
Const ProcName As String = "graphicObjectIsInScope"
On Error GoTo Err

With mGraphicObjects(index)
    If Not .graphObj.BoundingRectangle.isValid Then Exit Function
    If .graphObj.ExtendedObject Then
        graphicObjectIsInScope = IntOverlaps(RectGetXInterval(mViewport.Boundary), RectGetXInterval(.graphObj.BoundingRectangle))
    Else
        graphicObjectIsInScope = (.PeriodNumber >= Int(mViewport.Boundary.Left) And _
                                .PeriodNumber <= Int(mViewport.Boundary.Right))
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function graphicObjectIsVisible(ByVal index As Long) As Boolean
Const ProcName As String = "graphicObjectIsVisible"
On Error GoTo Err

With mGraphicObjects(index).graphObj
    If Not .BoundingRectangle.isValid Or .NoDraw Then Exit Function
    If .ExtendedObject And (.Capabilities And GraphicObjectCapabilities.ComplexBounding) Then
        Dim overlapXInterval As TInterval
        overlapXInterval = IntIntersection(RectGetXInterval(mViewport.Boundary), RectGetXInterval(.BoundingRectangle))

        Dim i As Long
        For i = Int(overlapXInterval.startValue) To _
                Int(overlapXInterval.endValue)
            
            Dim boundingRects() As TRectangle
            boundingRects = .BoundingRectanglesAt(i)

            Dim numrects As Long
            numrects = -1
            On Error Resume Next
            numrects = UBound(boundingRects)
            On Error GoTo Err
            If numrects >= 0 Then
                Dim j As Long
                For j = 0 To numrects
                    If RectOverlaps(boundingRects(j), mViewport.Boundary) Then
                        graphicObjectIsVisible = True
                        Exit Function
                    End If
                Next
            End If
        Next
    Else
        graphicObjectIsVisible = RectIntersection(mGraphicObjects(index).graphObj.BoundingRectangleInViewport, mViewport.Boundary).isValid
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialiseInScopeObjectList()
Const ProcName As String = "initialiseInScopeObjectList"
On Error GoTo Err

mInScopeCycle = mInScopeCycle + 1
mGraphicObjects(mFirstInScopeIndex).nextInScopeIndex = mLastInScopeIndex
mGraphicObjects(mLastInScopeIndex).prevInScopeIndex = mFirstInScopeIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialisePeriodTableEntry( _
                ByVal index As Long)
Const ProcName As String = "initialisePeriodTableEntry"
On Error GoTo Err

If mPeriodTable(index).firstIndex = 0 Then
    mPeriodTable(index).firstIndex = allocateGraphicObjectTableEntry
    mPeriodTable(index).lastindex = allocateGraphicObjectTableEntry
    mGraphicObjects(mPeriodTable(index).firstIndex).nextPeriodIndex = mPeriodTable(index).lastindex
    mGraphicObjects(mPeriodTable(index).lastindex).prevPeriodIndex = mPeriodTable(index).firstIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseVisibleObjectList()
Const ProcName As String = "initialiseVisibleObjectList"
On Error GoTo Err

mVisibleCycle = mVisibleCycle + 1

Dim i As Long
For i = LayerMin To LayerMax
    mGraphicObjects(mVisibleTable(i).firstIndex).nextVisibleIndex = mVisibleTable(i).lastindex
    mGraphicObjects(mVisibleTable(i).lastindex).prevVisibleIndex = mVisibleTable(i).firstIndex
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function isObjectOrSeriesSelectable(ByVal graphObj As IGraphicObject) As Boolean
Const ProcName As String = "isObjectOrSeriesSelectable"
On Error GoTo Err

If graphObj.IsSelectable Then
    isObjectOrSeriesSelectable = True
ElseIf mGraphObjSeriesCollection.Item(graphObj.SeriesID).IsSelectable Then
    isObjectOrSeriesSelectable = True
Else
    isObjectOrSeriesSelectable = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function releaseEntries( _
                ByVal pNumberToRelease As Long, _
                ByRef pIndex As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Boolean
Const ProcName As String = "releaseEntries"
On Error GoTo Err

Dim lNumberReleased As Long

Do While pIndex < mNextGraphicObjectIndex
    If Not mGraphicObjects(pIndex).graphObj Is Nothing Then
        Set mGraphicObjects(pIndex).graphObj = Nothing
        mNumberOfObjects = mNumberOfObjects - 1
    End If
    pIndex = pIndex + 1
    lNumberReleased = lNumberReleased + 1
    If lNumberReleased >= pNumberToRelease Then
        releaseEntries = True
        pTotalNumberReleased = pTotalNumberReleased + lNumberReleased
        pYieldCount = pYieldCount + 1
        Exit Function
    End If
Loop

pTotalNumberReleased = pTotalNumberReleased + lNumberReleased
pYieldCount = pYieldCount + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeEntryFromExtendedList(ByVal index As Long)
Const ProcName As String = "removeEntryFromExtendedList"
On Error GoTo Err

With mGraphicObjects(index)
    If .nextExtendedIndex = 0 Then Exit Sub
    mGraphicObjects(.nextExtendedIndex).prevExtendedIndex = .prevExtendedIndex
    mGraphicObjects(.prevExtendedIndex).nextExtendedIndex = .nextExtendedIndex
    .nextExtendedIndex = 0
    .prevExtendedIndex = 0
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeEntryFromInScopeList(ByVal index As Long)
Const ProcName As String = "removeEntryFromInScopeList"
On Error GoTo Err

With mGraphicObjects(index)
    If .inScopeCycle = mInScopeCycle Then
        .inScopeCycle = 0
        mGraphicObjects(.nextInScopeIndex).prevInScopeIndex = .prevInScopeIndex
        mGraphicObjects(.prevInScopeIndex).nextInScopeIndex = .nextInScopeIndex
        .nextInScopeIndex = 0
        .prevInScopeIndex = 0
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeEntryFromPeriodList(ByVal index As Long)
Const ProcName As String = "removeEntryFromPeriodList"
On Error GoTo Err

With mGraphicObjects(index)
    If .nextPeriodIndex = 0 Then Exit Sub
    mGraphicObjects(.nextPeriodIndex).prevPeriodIndex = .prevPeriodIndex
    mGraphicObjects(.prevPeriodIndex).nextPeriodIndex = .nextPeriodIndex
    .nextPeriodIndex = 0
    .prevPeriodIndex = 0
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeEntryFromVisibleList(ByVal index As Long)
Const ProcName As String = "removeEntryFromVisibleList"
On Error GoTo Err

With mGraphicObjects(index)
    If .visibleCycle = mVisibleCycle Then
        .visibleCycle = 0
        mGraphicObjects(.nextVisibleIndex).prevVisibleIndex = .prevVisibleIndex
        mGraphicObjects(.prevVisibleIndex).nextVisibleIndex = .nextVisibleIndex
        .nextVisibleIndex = 0
        .prevVisibleIndex = 0
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




