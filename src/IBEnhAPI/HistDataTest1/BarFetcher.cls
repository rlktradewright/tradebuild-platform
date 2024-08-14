VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IBarFetchListener

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

Private Const ModuleName                            As String = "BarFetcher"

'@================================================================================
' Member variables
'@================================================================================

Private mFirstBar                                   As Bar
Private mLastBar                                    As Bar

Private mNumberOfBars                               As Long

Private mShowEachBar                                As Boolean
Private mShowAllBarsAtEnd                           As Boolean
Private mContractSpec                               As IContractSpecifier

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBars                                       As Bars

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Sub IBarFetchListener_FetchCancelled(ByVal pCookie As Variant)
End Sub

Private Sub IBarFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchCompleted"
On Error GoTo Err

processReturnedBars

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchDataRetrieved(ByVal pCookie As Variant)
End Sub

Private Sub IBarFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
End Sub

Private Sub IBarFetchListener_FetchStarted(ByVal pCookie As Variant)
End Sub

Private Sub IBarFetchListener_NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)
Const ProcName As String = "IBarFetchListener_NotifyBar"
On Error GoTo Err

Static sGotFirstBar As Boolean

mNumberOfBars = mNumberOfBars + 1

If Not sGotFirstBar Then
    sGotFirstBar = True
    Set mFirstBar = pBar
Else
    Set mLastBar = pBar
End If

If mShowEachBar Then LogMessage mContractSpec.LocalSymbol & "." & mNumberOfBars & ": " & pBar.ToString

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

Dim lContractSpec As IContractSpecifier: Set lContractSpec = ev.Future.Cookie
If ev.Future.IsFaulted Then
    LogMessage "Error " & ev.Future.ErrorNumber & " for " & lContractSpec.ToString & vbCrLf & _
                        ev.Future.ErrorMessage & vbCrLf & _
                        ev.Future.ErrorSource
ElseIf ev.Future.IsCancelled Then
    LogMessage "Historical data fetch cancelled for " & lContractSpec.ToString
ElseIf TypeOf ev.Future.Value Is IContract Then
    Dim lContract As IContract: Set lContract = ev.Future.Value
    Set mContractSpec = lContract.Specifier
ElseIf TypeOf ev.Future.Value Is Bars Then
    Set mBars = ev.Future.Value
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function Fetch( _
                ByVal pBarSpec As BarDataSpecifier, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pContractFuture As IFuture, _
                ByVal pShowEachBar As Boolean, _
                ByVal pShowAllBarsAtEnd As Boolean, _
                ByVal pCookie As Variant) As IFuture
Const ProcName As String = "Fetch"
On Error GoTo Err

mShowEachBar = pShowEachBar
mShowAllBarsAtEnd = pShowAllBarsAtEnd
Set Fetch = pHistDataStore.FetchBars( _
                            pContractFuture, _
                            pBarSpec, _
                            pListener:=Me, _
                            pCookie:=pCookie)
mFutureWaiter.Add Fetch
mFutureWaiter.Add pContractFuture

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub processReturnedBars()
Const ProcName As String = "processReturnedBars"
On Error GoTo Err

LogMessage "Historical data fetch completed: " & mBars.Count & " bars retrieved for " & mContractSpec.ToString
If mBars.Count = 0 Then Exit Sub

LogMessage "First bar: " & mBars.Bar(1).ToString
LogMessage "Last bar: " & mBars.Bar(0).ToString

If mShowAllBarsAtEnd Then
    Dim i As Long
    For i = 1 To mBars.Count
        LogMessage mContractSpec.LocalSymbol & "." & i & ": " & mBars.Bar(i).ToString
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




