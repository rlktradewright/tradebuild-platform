VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjSeriesCollection"
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GraphObjSeriesCollection"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphObjSeriesTable()                      As IGraphicObjectSeries
Private mGraphObjSeriesIndex                        As Long
Private mRegionRef                                  As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mGraphObjSeriesTable(3) As IGraphicObjectSeries
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal pGraphObjSeries As IGraphicObjectSeries, _
                ByVal pLayer As Long, _
                ByVal pName As String, _
                ByVal pViewport As Viewport, _
                ByVal pChart As Chart, _
                ByVal pGOSP As GraphObjServiceProvider) As IGraphicObjectSeries
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

Set Add = pGraphObjSeries

If mGraphObjSeriesIndex > UBound(mGraphObjSeriesTable) Then
    ReDim Preserve mGraphObjSeriesTable(2 * (UBound(mGraphObjSeriesTable) + 1) - 1) As IGraphicObjectSeries
End If

Set mGraphObjSeriesTable(mGraphObjSeriesIndex) = pGraphObjSeries

If pName = "" Then pName = "Series" & mGraphObjSeriesIndex

pGraphObjSeries.Initialise pName, _
                        pViewport, _
                        pLayer, _
                        pChart, _
                        mGraphObjSeriesIndex, _
                        pGOSP

mGraphObjSeriesIndex = mGraphObjSeriesIndex + 1

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Finish()
Dim i As Long
Dim t As GraphObjSeriesRemovalTask

' remove all graphic object references from the various collections
Dim graphObjSeries As IGraphicObjectSeries
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

For i = 0 To mGraphObjSeriesIndex - 1
    If Not mGraphObjSeriesTable(i) Is Nothing Then
        mGraphObjSeriesTable(i).Finish
        
        Set t = New GraphObjSeriesRemovalTask
        t.Initialise mGraphObjSeriesTable(i), regionObj, False
        StartTask t, PriorityLow

        Set mGraphObjSeriesTable(i) = Nothing
    End If
Next
Erase mGraphObjSeriesTable

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Friend Sub Initialise(ByVal pRegion As ChartRegion)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mRegionRef = CreateWeakReference(pRegion)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function Item( _
                ByVal index As Long) As IGraphicObjectSeries
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mGraphObjSeriesTable(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Remove( _
                ByVal pSeries As IGraphicObjectSeries)
Dim t As New GraphObjSeriesRemovalTask
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

t.Initialise pSeries, regionObj, True
StartTask t, PriorityLow
Set mGraphObjSeriesTable(pSeries.SeriesID) = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function regionObj() As ChartRegion
Const ProcName As String = "regionObj"
Dim failpoint As String
On Error GoTo Err

Set regionObj = mRegionRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function
