VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DunniganBarFormatter"
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

Implements IBarFormatter

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

Private Const ModuleName                    As String = "DunniganBarFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mBars                               As Bars
Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
 
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IBarFormatter Interface Members
'@================================================================================

Private Sub IBarFormatter_FormatBar( _
                ByVal sourceBar As BarUtils27.Bar, _
                ByVal chartBar As ChartSkil27.Bar)
Const ProcName As String = "IBarFormatter_FormatBar"
On Error GoTo Err

If mBars.Count < 2 Then Exit Sub

Dim lFirstBarNumber As Long
lFirstBarNumber = mBars.Bar(1).BarNumber

Dim lPrevBarIndex As Long
lPrevBarIndex = sourceBar.BarNumber - lFirstBarNumber

If lPrevBarIndex < 1 Then Exit Sub

Dim lPrevBar As BarUtils27.Bar
Set lPrevBar = mBars.Bar(lPrevBarIndex)

If sourceBar.HighValue > lPrevBar.HighValue And _
    sourceBar.LowValue >= lPrevBar.LowValue _
Then
    chartBar.Color = &H1D9311
    chartBar.UpColor = &H1D9311
    chartBar.DownColor = &H1D9311
ElseIf sourceBar.HighValue <= lPrevBar.HighValue And _
    sourceBar.LowValue < lPrevBar.LowValue _
Then
    chartBar.Color = &H43FC2
    chartBar.UpColor = &H43FC2
    chartBar.DownColor = &H43FC2
ElseIf sourceBar.HighValue <= lPrevBar.HighValue And _
    sourceBar.LowValue >= lPrevBar.LowValue _
Then
    chartBar.Color = vbMagenta
    chartBar.UpColor = vbMagenta
    chartBar.DownColor = vbMagenta
ElseIf sourceBar.HighValue >= lPrevBar.HighValue And _
    sourceBar.LowValue <= lPrevBar.LowValue _
Then
    chartBar.Color = vbBlue
    chartBar.UpColor = vbBlue
    chartBar.DownColor = vbBlue
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then Set mBars = ev.Future.Value

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let BarsFuture(ByVal Value As IFuture)
Const ProcName As String = "BarsFuture"
On Error GoTo Err

mFutureWaiter.Add Value

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


