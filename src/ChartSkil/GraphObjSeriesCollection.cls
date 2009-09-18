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
                ByVal pviewport As Viewport, _
                ByVal pChart As Chart, _
                ByVal pGOSP As GraphObjServiceProvider) As IGraphicObjectSeries
Dim failpoint As Long
On Error GoTo Err

Set Add = pGraphObjSeries

If mGraphObjSeriesIndex > UBound(mGraphObjSeriesTable) Then
    ReDim Preserve mGraphObjSeriesTable(2 * (UBound(mGraphObjSeriesTable) + 1) - 1) As IGraphicObjectSeries
End If

Set mGraphObjSeriesTable(mGraphObjSeriesIndex) = pGraphObjSeries

If pName = "" Then pName = "Series" & mGraphObjSeriesIndex

pGraphObjSeries.Initialise pName, _
                        pviewport, _
                        pLayer, _
                        pChart, _
                        mGraphObjSeriesIndex, _
                        pGOSP

mGraphObjSeriesIndex = mGraphObjSeriesIndex + 1

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "addGraphicObjectSeries" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription

End Function

Friend Sub Finish()
Dim i As Long
Dim t As GraphObjSeriesRemovalTask

' remove all graphic object references from the various collections
Dim graphObjSeries As IGraphicObjectSeries
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

End Sub

Friend Sub Initialise(ByVal pRegion As ChartRegion)
Set mRegionRef = CreateWeakReference(pRegion)
End Sub

Friend Function Item( _
                ByVal index As Long) As IGraphicObjectSeries
Set Item = mGraphObjSeriesTable(index)
End Function

Friend Sub Remove( _
                ByVal pSeries As IGraphicObjectSeries)
Dim t As New GraphObjSeriesRemovalTask
t.Initialise pSeries, regionObj, True
StartTask t, PriorityLow
Set mGraphObjSeriesTable(pSeries.SeriesID) = Nothing
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function
