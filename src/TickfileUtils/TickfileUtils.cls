VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

Public Enum TickfileAccessModes
    TickfileReadOnly
    TickfileWriteOnly
    TickfileReadWrite
End Enum

Public Enum ReplayContinueModes
    ReplayContinue
    ReplayCancelThisFile
    ReplayCancelAllFiles
End Enum

Public Enum TickfileEventCodes
    TickfileEventFileDoesNotExist = 1
    TickfileEventFileIsEmpty
    TickfileEventFileIsInvalid
    TickfileEventFileFormatNotSupported
    TickfileEventNoContractDetails
    TickfileEventDataSourceNotAvailable
    TickfileEventAmbiguousContractDetails
End Enum

Public Enum TickfileGranularities
    TickfileGranularityDay
    TickfileGranularityWeek
    TickfileGranularitySession
    TickfileGranularityExecution
End Enum

Public Enum TickfileModes
    TickfileModeFileBased = 1
    TickfileModeStreamBased = 2
End Enum

Public Enum TickfileStoreCapabilities
    TickfileStoreCanRecord = 1&
    TickfileStoreCanRecordMarketDepth = 2&
    TickfileStoreCanReplay = 4&
    TickfileStoreCanPositionExact = 8&
    TickfileStoreCanReplayMarketDepth = 16&
    TickfileStoreCanReportReplayProgress = 32&
    TickfileStoreCanSaveContractInformation = 64&
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TickfileFormatSpecifier
    ' Indicates the tickfile mode
    FormatType As TickfileModes

    ' Bitmask of TickfileStoreCapabilities values indicating the capabilities supported by this tickfile format
    Capabilities As TickfileStoreCapabilities

    ' The unique id identifying this tickfile format
    FormalID As String

    ' The informal (human readable but not necessarily unique) name for this tickfile format
    Name As String

    ' For file-based tickfile formats, this is the filename extension used to identify tickfiles in this format
    FileExtension As String
End Type

Public Type TickfileEventData
    Source              As Object
    TickStream          As ITickStream
    Specifier           As TickfileSpecifier
    SizeInBytes         As Long
End Type

Public Type TickWriterPendingComment
    Comment             As String
    Timestamp           As Date
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickfileUtils"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

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

Public Function CreateBufferedTickfileWriter( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As ITickfileWriter
Const ProcName As String = "CreateBufferedTickfileWriter"
On Error GoTo Err

Set CreateBufferedTickfileWriter = gCreateBufferedTickfileWriter(pTickfileStore, pOutputMonitor, pContractFuture, pFormatIdentifier, pLocation)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickfileReplayController( _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pReplaySpeed As Long = 1, _
                Optional ByVal pReplayProgressEventInterval As Long = 1000, _
                Optional ByVal pTimestampAdjustmentStart As Double = 0#, _
                Optional ByVal pTimestampAdjustmentEnd As Double = 0#) As ReplayController
Const ProcName As String = "CreateTickfileReplayController"
On Error GoTo Err

AssertArgument Not pTickfileStore Is Nothing, "pTickfileStore is Nothing"
AssertArgument pReplayProgressEventInterval >= 50, "pReplayProgressEventInterval cannot be less than 50"

Dim clr As New ReplayController
clr.Intialise pTickfileStore, pPrimaryContractStore, pSecondaryContractStore, pReplaySpeed, pTimestampAdjustmentStart, pTimestampAdjustmentEnd, pReplayProgressEventInterval

Set CreateTickfileReplayController = clr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTickfileStore( _
                ByVal pMode As TickfileAccessModes, _
                Optional ByVal pOutputTickfilePath As String, _
                Optional ByVal pTickfileGranularity As TickfileGranularities = TickfileGranularityWeek) As ITickfileStore
Const ProcName As String = "CreateTickfileStore"
On Error GoTo Err

Dim lStore As New TickfileStore
lStore.Initialise pMode, pOutputTickfilePath, pTickfileGranularity
Set CreateTickfileStore = lStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GenerateTickfileSpecifiers( _
                ByVal pContracts As IContracts, _
                ByVal pTickfileFormatID As String, _
                ByVal pStartDate As Date, _
                ByVal pEndDate As Date, _
                Optional ByVal pCompleteSessionsOnly As Boolean = True, _
                Optional ByVal pUseExchangeTimezone As Boolean = True, _
                Optional ByVal pCustomSessionStartTime As Date, _
                Optional ByVal pCustomSessionEndTime As Date) As TickFileSpecifiers
Const ProcName As String = "GenerateTickfileSpecifiers"
On Error GoTo Err

Dim tfsg As New TickfileSpecGenerator
tfsg.Initialise pContracts, _
                pTickfileFormatID, _
                pStartDate, _
                pEndDate, _
                pCompleteSessionsOnly, _
                pUseExchangeTimezone, _
                pCustomSessionStartTime, _
                pCustomSessionEndTime

Set GenerateTickfileSpecifiers = tfsg.Generate()

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GenerateTickfileSpecifiersFromFile( _
                ByVal pFilename As String) As TickFileSpecifiers
Const ProcName As String = "GenerateTickfileSpecifiersFromFile"
On Error GoTo Err

Set GenerateTickfileSpecifiersFromFile = gParseTickfileListFile(pFilename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecordTickData( _
                ByVal pTickSource As IGenericTickSource, _
                ByVal pContractFuture As IFuture, _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pOutputMonitor As ITickfileOutputMonitor, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As TickDataWriter
Const ProcName As String = "RecordTickData"
On Error GoTo Err

Set RecordTickData = New TickDataWriter
RecordTickData.Initialise pTickSource, pContractFuture, pOutputMonitor, pTickfileStore, pFormatIdentifier, pLocation

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================



