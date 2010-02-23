VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements Enumerable

'@================================================================================
' Events
'@================================================================================

''
' Raised when an Item is added to or removed from this
' <code>ChartRegions</code> object, and when the collection is cleared.
'
' @remarks
'   The following members of the <code>CollectionChangeTypes</code> enum may
'   occur in the <code>ChangeType</code> field of the
'   <code>CollectionChangedEvent</code> parameter:
' <blockquote>
'    <tt>CollItemAdded</tt><br />
'    <tt>CollItemRemoved</tt><br />
'    <tt>CollCollectionCleared</tt>
' </blockquote>
' @param ev
'   contains details of the change.
'
'@/
Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type RegionTableEntry
    Region              As ChartRegion
    PercentHeight       As Double
'    ActualHeight        As Long
    useAvailableSpace   As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ChartRegions"

'@================================================================================
' Member variables
'@================================================================================

Private mChartRef                                   As WeakReference

Private mRegions()                                  As RegionTableEntry
Private mRegionsIndex                               As Long
Private mNumRegionsInUse                            As Long

Private mUpdateSequence                             As Long

Private mDefaultDataRegionStyle                     As ChartRegionStyle
Private mDefaultYAxisRegionStyle                    As ChartRegionStyle

Private mRegionHeightReductionFactor                As Double


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRegions(3) As RegionTableEntry
mRegionsIndex = 0
mRegionHeightReductionFactor = 1
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As TWUtilities30.Enumerator
Set Enumerable_Enumerator = getEnumerator
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
If cookie = &H7FFFFFFF Then Exit Function
If IsEmpty(cookie) Then cookie = -1

Do
    cookie = cookie + 2
    If cookie > mRegionsIndex Then
        cookie = &H7FFFFFFF
        Exit Function
    End If
Loop Until (Not mRegions(cookie).Region Is Nothing)

Set Item = mRegions(cookie).Region

Enumerable_NextItem = True
End Function

Private Sub Enumerable_Remove(cookie As Variant)
Remove mRegions(cookie).Region
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_Skip"
Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
        ProjectName & "." & ModuleName & ":" & ProcName, _
        "Skip not supported"
End Function

Private Property Get Enumerable_UpdateSequence() As Long
Enumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Chart(ByVal value As Chart)
Const ProcName As String = "Chart"
Dim failpoint As String
On Error GoTo Err

Set mChartRef = CreateWeakReference(value)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Count() As Long
Count = mNumRegionsInUse
End Property

Public Property Let DefaultDataRegionStyle( _
                ByVal value As ChartRegionStyle)
Set mDefaultDataRegionStyle = value
End Property

Public Property Get DefaultDataRegionStyle() As ChartRegionStyle
Const ProcName As String = "DefaultDataRegionStyle"
Dim failpoint As String
On Error GoTo Err

If mDefaultDataRegionStyle Is Nothing Then Set mDefaultDataRegionStyle = New ChartRegionStyle
Set DefaultDataRegionStyle = mDefaultDataRegionStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let DefaultYAxisRegionStyle( _
                ByVal value As ChartRegionStyle)
Set mDefaultYAxisRegionStyle = value
End Property

Public Property Get DefaultYAxisRegionStyle() As ChartRegionStyle
If mDefaultYAxisRegionStyle Is Nothing Then Set mDefaultYAxisRegionStyle = DefaultDataRegionStyle
Set DefaultYAxisRegionStyle = mDefaultYAxisRegionStyle
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Set Enumerator = getEnumerator
End Property

Public Property Get Item(ByVal Name As String) As ChartRegion
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "400"
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = findRegion(Name)

If Item Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Item does not exist"
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                Optional ByVal PercentHeight As Double = 100#, _
                Optional ByVal MinimumPercentHeight As Double, _
                Optional ByVal Style As ChartRegionStyle, _
                Optional ByVal yAxisStyle As ChartRegionStyle, _
                Optional ByVal Name As String) As ChartRegion

Dim YAxisRegion As ChartRegion
Dim availableheight As Long
Dim lRegion As ChartRegion


'
' NB: PercentHeight=100 means the region will use whatever space
' is available
'

Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

availableheight = chartObj.availableheight

If Style Is Nothing Then Set Style = DefaultDataRegionStyle
If yAxisStyle Is Nothing Then Set yAxisStyle = DefaultYAxisRegionStyle

If Name <> "" Then
    If Contains(Name) Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Region " & Name & " already exists"
    End If
Else
    Name = GenerateGUIDString
End If

mRegionsIndex = mRegionsIndex + 1

Set lRegion = New ChartRegion
lRegion.handle = mRegionsIndex

lRegion.Initialise Name, chartObj, chartObj.CreateDataRegionCanvas(lRegion.handle), RegionTypeData

lRegion.MinimumPercentHeight = MinimumPercentHeight
lRegion.PercentHeight = PercentHeight
lRegion.Style = Style

Set mRegions(mRegionsIndex).Region = lRegion
If PercentHeight <> 100 Then
    mRegions(mRegionsIndex).PercentHeight = mRegionHeightReductionFactor * PercentHeight
Else
    mRegions(mRegionsIndex).useAvailableSpace = True
End If

mRegionsIndex = mRegionsIndex + 1
If mRegionsIndex > UBound(mRegions) Then
    ReDim Preserve mRegions(2 * (UBound(mRegions) + 1) - 1) As RegionTableEntry
End If

Set YAxisRegion = New ChartRegion
YAxisRegion.handle = mRegionsIndex

YAxisRegion.Initialise "", chartObj, chartObj.CreateYAxisRegionCanvas(lRegion.handle), RegionTypeYAxis

YAxisRegion.Style = yAxisStyle
lRegion.YAxisRegion = YAxisRegion

Set mRegions(mRegionsIndex).Region = YAxisRegion

mNumRegionsInUse = mNumRegionsInUse + 1

If ResizeY(False) Then
    changed
    fireChange CollItemAdded, lRegion
Else
    ' can't fit this all in! So remove the added region,
    Set lRegion = Nothing
    Set mRegions(mRegionsIndex).Region = Nothing
    mRegions(mRegionsIndex).PercentHeight = 0
    mRegions(mRegionsIndex).useAvailableSpace = False
    mRegionsIndex = mRegionsIndex - 2
    mNumRegionsInUse = mNumRegionsInUse - 1
    Exit Function
End If

Set Add = lRegion

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function Contains(ByVal Name As String) As Boolean
Dim lRegion As ChartRegion

Const ProcName As String = "Contains"
Dim failpoint As String
On Error GoTo Err

Set lRegion = findRegion(Name)
If Not lRegion Is Nothing Then Contains = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Sub Finish()
Set mChartRef = Nothing
End Sub

Friend Function ItemFromHandle(ByVal pHandle As Long) As ChartRegion
Const ProcName As String = "ItemFromHandle"
Dim failpoint As String
On Error GoTo Err

Set ItemFromHandle = mRegions(pHandle).Region

If ItemFromHandle Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid Handle"
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
Dim failpoint As String
On Error GoTo Err

Set NewEnum = getEnumerator

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Sub Remove( _
                ByVal pRegion As ChartRegion)
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

If pRegion.IsXAxisRegion Or pRegion.IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Cannot remove an axis region"
End If

If mRegions(pRegion.handle).Region Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Region has already been removed"
End If

pRegion.ClearRegion
'pRegion.YAxisRegion.ClearRegion
Set mRegions(pRegion.handle).Region = Nothing
Set mRegions(pRegion.handle + 1).Region = Nothing

mNumRegionsInUse = mNumRegionsInUse - 1

changed
fireChange CollItemRemoved, pRegion

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function ResizeRegion( _
                ByVal pRegion As ChartRegion, _
                ByVal vertChange As Long) As Boolean
Dim handle As Long
Dim newHeight As Long
Dim prevPercentHeight As Double
Dim availableheight As Long

Const ProcName As String = "ResizeRegion"
Dim failpoint As String
On Error GoTo Err

handle = pRegion.handle
availableheight = chartObj.availableheight

newHeight = pRegion.ActualHeight + vertChange
If newHeight < 0 Then newHeight = 0

' the region table indicates the requested percentage used by each region
' and the actual Height allocation. We need to work out the new percentage
' for the region to be resized.

prevPercentHeight = mRegions(handle).PercentHeight
mRegions(handle).PercentHeight = 100 * newHeight / availableheight
If mRegions(handle).useAvailableSpace Then
    ' this is a 'use available space' region that's being resized. Now change
    ' it to use a specific percentage
    mRegions(handle).useAvailableSpace = False
    mRegions(handle).Region.PercentHeight = mRegions(handle).PercentHeight
End If

If ResizeY(True) Then
    ResizeRegion = True
Else
    ' the regions couldn't be resized so reset the region's percent Height
    mRegions(handle).PercentHeight = prevPercentHeight
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Function ResizeY( _
                ByVal userResizing As Boolean) As Boolean
'
' NB: PercentHeight=100 means the region will use whatever space
' is available
'
Dim i As Long
Dim aRegion As ChartRegion
Dim numAvailableSpaceRegions As Long
Dim totalMinimumPercents As Double
Dim nonFixedAvailableSpacePercent As Double
Dim availableSpacePercent As Double
Dim availableheight As Long

Const ProcName As String = "ResizeY"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

availableSpacePercent = 100
nonFixedAvailableSpacePercent = 100
availableheight = chartObj.availableheight

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
        If Not mRegions(i).useAvailableSpace Then
            availableSpacePercent = availableSpacePercent - mRegions(i).PercentHeight
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).PercentHeight
        Else
            If aRegion.MinimumPercentHeight <> 0 Then
                availableSpacePercent = availableSpacePercent - aRegion.MinimumPercentHeight
            End If
            numAvailableSpaceRegions = numAvailableSpaceRegions + 1
        End If
    End If
Next

If availableSpacePercent < 0 And userResizing Then
    ResizeY = False
    Exit Function
End If


failpoint = 200

mRegionHeightReductionFactor = 1
Do While availableSpacePercent < 0 And mRegionHeightReductionFactor >= 0.0625
    availableSpacePercent = 100
    nonFixedAvailableSpacePercent = 100
    totalMinimumPercents = 0
    mRegionHeightReductionFactor = mRegionHeightReductionFactor - 0.0625
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).Region Is Nothing Then
            Set aRegion = mRegions(i).Region
            If Not mRegions(i).useAvailableSpace Then
                If aRegion.MinimumPercentHeight <> 0 Then
                    If mRegions(i).PercentHeight * mRegionHeightReductionFactor >= _
                        aRegion.MinimumPercentHeight _
                    Then
                        mRegions(i).PercentHeight = mRegions(i).PercentHeight * mRegionHeightReductionFactor
                    Else
                        mRegions(i).PercentHeight = aRegion.MinimumPercentHeight
                    End If
                    totalMinimumPercents = totalMinimumPercents + aRegion.MinimumPercentHeight
                Else
                    mRegions(i).PercentHeight = mRegions(i).PercentHeight * mRegionHeightReductionFactor
                End If
                availableSpacePercent = availableSpacePercent - mRegions(i).PercentHeight
                nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).PercentHeight
            Else
                If aRegion.MinimumPercentHeight <> 0 Then
                    availableSpacePercent = availableSpacePercent - aRegion.MinimumPercentHeight
                    totalMinimumPercents = totalMinimumPercents + aRegion.MinimumPercentHeight
                End If
            End If
        End If
    Next
    If totalMinimumPercents > 100 Then
        ' can't possibly fit this all in!
        ResizeY = False
        Exit Function
    End If
Loop


failpoint = 300

If numAvailableSpaceRegions = 0 Then
    ' we must adjust the percentages on the other regions so they
    ' total 100.
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).Region Is Nothing Then
            mRegions(i).PercentHeight = 100 * mRegions(i).PercentHeight / (100 - nonFixedAvailableSpacePercent)
        End If
    Next
End If

' first set Heights for fixed Height regions

failpoint = 400

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        If Not mRegions(i).useAvailableSpace Then
            mRegions(i).Region.ActualHeight = mRegions(i).PercentHeight * availableheight / 100
            Debug.Assert mRegions(i).Region.ActualHeight >= 0
        End If
    End If
Next


failpoint = 500

' now set Heights for 'available space' regions with a minimum Height
' that needs to be respected
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
        If mRegions(i).useAvailableSpace Then
            mRegions(i).Region.ActualHeight = 0
            If aRegion.MinimumPercentHeight <> 0 Then
                If (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) < aRegion.MinimumPercentHeight Then
                    mRegions(i).Region.ActualHeight = aRegion.MinimumPercentHeight * availableheight / 100
                    Debug.Assert mRegions(i).Region.ActualHeight >= 0
                    nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.MinimumPercentHeight
                    numAvailableSpaceRegions = numAvailableSpaceRegions - 1
                End If
            End If
        End If
    End If
Next


failpoint = 600

' finally set Heights for all other 'available space' regions
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        If mRegions(i).useAvailableSpace And _
            mRegions(i).Region.ActualHeight = 0 _
        Then
            mRegions(i).Region.ActualHeight = (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) * availableheight / 100
            Debug.Assert mRegions(i).Region.ActualHeight >= 0
        End If
    End If
Next


failpoint = 700

ResizeY = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub changed()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Function chartObj() As Chart
Const ProcName As String = "chartObj"
Dim failpoint As String
On Error GoTo Err

Set chartObj = mChartRef.Target

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedItem As Variant)
Dim ev As CollectionChangeEvent
Const ProcName As String = "fireChange"
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.changeType = changeType
If IsObject(affectedItem) Then
    If Not affectedItem Is Nothing Then Set ev.affectedItem = affectedItem
Else
    ev.affectedItem = affectedItem
End If
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
Dim failpoint As String
On Error GoTo Err

Set getEnumerator = CreateEnumerator(Me)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function findRegion( _
                ByVal Name As String) As ChartRegion
Dim en As Enumerator
Dim lRegion As ChartRegion

Const ProcName As String = "findRegion"
Dim failpoint As String
On Error GoTo Err

Name = UCase$(Name)

Set en = getEnumerator

Do While en.MoveNext
    Set lRegion = en.current
    If UCase$(lRegion.Name) = Name Then
        Set findRegion = lRegion
        Exit Do
    End If
Loop

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function
