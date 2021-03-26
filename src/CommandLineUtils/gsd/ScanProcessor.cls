VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScanProcessor"
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

Private Const ModuleName                            As String = "ScanProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mScanName                                   As String
Private mScanParameters                             As Parameters

Private mScanResultFormat                           As String
Private mScanEchoResultFormat                       As String

Private mOutputPath                                 As String
Private mOutputFilename                             As String
Private mAppend                                     As Boolean

Private mEchoToStdOut                               As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim lMessage As String

If ev.Future.IsCancelled Then
    lMessage = "Scan cancelled"
ElseIf ev.Future.IsFaulted Then
    lMessage = ev.Future.ErrorMessage
Else
    Dim lScanResults As IScanResults
    Set lScanResults = ev.Future.Value
    lMessage = "scan results: " & lScanResults.Count
    
    Dim lStream As TextStream
    
    If mOutputPath = "" And mOutputFilename = "" Then
    ElseIf lScanResults.Count = 0 Then
    Else
        Set lStream = gCreateOutputStream( _
                            mOutputPath, _
                            mOutputFilename, _
                            Me, _
                            mAppend, _
                            lMessage)
        If lStream Is Nothing Then
            gNotifyProcessorCompleted Me, lMessage
            Exit Sub
        End If
    End If
        
    Dim lScanResult As IScanResult
    For Each lScanResult In lScanResults
        If lStream Is Nothing Or mEchoToStdOut Then gWriteLineToStdOut formatScanResult(lScanResult, True)
        If Not lStream Is Nothing Then lStream.WriteLine formatScanResult(lScanResult, False)
    Next
End If

gNotifyProcessorCompleted Me, lMessage
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ScanParameters() As Parameters
Set ScanParameters = mScanParameters
End Property

Friend Property Get ScanName() As String
ScanName = mScanName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Scan( _
                ByVal pHistDataStore As HistoricalDataStore, _
                ByVal pScanName As String, _
                ByVal pScanParameters As Parameters, _
                ByVal pOptions As Parameters, _
                ByVal pFilters As Parameters, _
                ByVal pAutoRefresh As Boolean, _
                ByVal pScanResultFormat As String, _
                ByVal pScanEchoResultFormat As String, _
                ByVal pOutputPath As String, _
                ByVal pOutputFilename As String, _
                ByVal pAppend As Boolean, _
                ByVal pEchoToStdOut As Boolean)
Const ProcName As String = "Scan"
On Error GoTo Err

mScanName = pScanName
Set mScanParameters = pScanParameters
mScanResultFormat = pScanResultFormat
mScanEchoResultFormat = pScanEchoResultFormat

mOutputPath = pOutputPath
mOutputFilename = pOutputFilename

mAppend = pAppend
mEchoToStdOut = pEchoToStdOut

Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add pHistDataStore.ScanMarket( _
                                    mScanName, _
                                    pScanParameters, _
                                    pOptions, _
                                    pFilters, _
                                    pAutoRefresh, _
                                    IIf(pAutoRefresh, GenerateGUIDString, ""))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatScanResult( _
                ByVal pScanResult As IScanResult, _
                ByVal pUseEchoFormat As Boolean) As String
Const ProcName As String = "formatScanResult"
On Error GoTo Err

formatScanResult = gPerformResultVariableSubstitution( _
                        IIf(pUseEchoFormat, mScanEchoResultFormat, mScanResultFormat), _
                        pScanResult, _
                        Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



