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

AssertArgument (pCustomSessionStartTime = 0 And pCustomSessionEndTime = 0) Or _
    (pCustomSessionStartTime <> 0 And pCustomSessionEndTime <> 0), "Session start and end times must be either both supplied or both omitted"
    
Dim lUseCustomSessionTimes As Boolean
lUseCustomSessionTimes = (pCustomSessionStartTime <> 0)

' get the most recent contract (though they should all have the same
' info regarding Session times)
Dim lContract As IContract
Set lContract = pContracts.ItemAtIndex(pContracts.Count)

Dim lContractTimezone As TimeZone
Set lContractTimezone = GetTimeZone(lContract.TimezoneName)

' the from and to dates converted (if necessary) to the contract's timezone
Dim FromDate As Date
Dim ToDate As Date
If pUseExchangeTimezone Then
    FromDate = pStartDate
    If pEndDate <> 0 Then ToDate = pEndDate
Else
    FromDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pStartDate), lContractTimezone)
    If pEndDate <> 0 Then ToDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pEndDate), lContractTimezone)
End If

' determine start and end times ---------------------------------------------------

Dim lSessionStartTime As Date
Dim lSessionEndTime As Date

' note that the custom times are in the contract's timezone
If lUseCustomSessionTimes Then
    lSessionStartTime = pCustomSessionStartTime
    lSessionEndTime = pCustomSessionEndTime
Else
    lSessionStartTime = lContract.sessionStartTime
    lSessionEndTime = lContract.sessionEndTime
End If

' set the Session start and end times for the starting date (in the contract's
' local timezone)...
Dim fromSessionTimes As SessionTimes
fromSessionTimes = GetSessionTimes(FromDate, lSessionStartTime, lSessionEndTime)

' the from and to dates, Session-oriented if required
Dim replayFromDate As Date
If pCompleteSessionsOnly Then
    replayFromDate = fromSessionTimes.StartTime
Else
    replayFromDate = FromDate
End If

Dim toSessionTimes As SessionTimes
Dim replayToDate As Date
If ToDate <> 0 Then
    ' set the Session start and end times for the ending date (in the contract's
    ' local timezone)...
    toSessionTimes = GetSessionTimes(ToDate, lSessionStartTime, lSessionEndTime)
    
    If pCompleteSessionsOnly Then
        replayToDate = toSessionTimes.EndTime
    Else
        replayToDate = ToDate
    End If
Else
    toSessionTimes = fromSessionTimes
    replayToDate = toSessionTimes.EndTime
End If

' find contract for start date ------------------------------------------------------

Dim en As Enumerator
Set en = pContracts.Enumerator

Dim lSecType As SecurityTypes
lSecType = pContracts.ItemAtIndex(1).Specifier.SecType

Dim lCurrContract As IContract
Dim aContract As IContract

If lSecType <> SecurityTypes.SecTypeFuture And _
    lSecType <> SecurityTypes.SecTypeOption And _
    lSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    Set lCurrContract = pContracts.ItemAtIndex(1)
Else
    Do While en.MoveNext
    Set aContract = en.Current
        If DateValue(FromDate) <= _
            (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
        Then
            Set lCurrContract = aContract
            Exit Do
        End If
    Loop
    
    AssertArgument Not lCurrContract Is Nothing, "No contract for this from date"
End If

If pCompleteSessionsOnly Then
    If lUseCustomSessionTimes Then
        Set lCurrContract = editContractSessionTimes(lCurrContract, pCustomSessionStartTime, pCustomSessionEndTime)
    End If
End If

Dim currSessionStart As Date
Dim thisSessionTimes As SessionTimes

currSessionStart = FromDate

Dim lTickFileSpecifiers As New TickFileSpecifiers
Dim tfs As TickfileSpecifier

If pCompleteSessionsOnly Then
    Do While currSessionStart < replayToDate
        Set tfs = New TickfileSpecifier
        lTickFileSpecifiers.Add tfs
        tfs.Contract = lCurrContract
        tfs.TickfileFormatID = pTickfileFormatID
        
        If lUseCustomSessionTimes Then
            thisSessionTimes = GetSessionTimes(currSessionStart, lSessionStartTime, lSessionEndTime)
            tfs.FromDate = thisSessionTimes.StartTime
            tfs.ToDate = thisSessionTimes.EndTime
            tfs.EntireSession = False
            tfs.Filename = generateNonSessionFilename(lCurrContract, tfs.FromDate, tfs.ToDate)
        Else
            tfs.FromDate = currSessionStart
            tfs.EntireSession = True
            tfs.Filename = generateSessionFilename(lCurrContract, currSessionStart)
            
        End If
        
        currSessionStart = currSessionStart + 1
        
        If lSecType = SecurityTypes.SecTypeFuture Or _
            lSecType = SecurityTypes.SecTypeOption Or _
            lSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (lCurrContract.ExpiryDate - lCurrContract.DaysBeforeExpiryToSwitch) _
            Then
                Do While en.MoveNext
                    Set aContract = en.Current
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
                    Then
                        Set lCurrContract = aContract
                        If lUseCustomSessionTimes Then
                            Set lCurrContract = editContractSessionTimes(lCurrContract, _
                                                                        pCustomSessionStartTime, _
                                                                        pCustomSessionEndTime)
                        End If
                        Exit Do
                    End If
                Loop
                AssertArgument Not lCurrContract Is Nothing, "No contract from " & currSessionStart
            End If
        End If
        
    Loop
    AssertArgument lTickFileSpecifiers.Count <> 0, "No trading sessions in specified date range"
Else
    Set tfs = New TickfileSpecifier
    lTickFileSpecifiers.Add tfs
    tfs.Contract = lCurrContract
    tfs.TickfileFormatID = pTickfileFormatID

    tfs.FromDate = FromDate
    currSessionStart = currSessionStart + 1
    Do While currSessionStart < replayToDate
        
        If lSecType = SecurityTypes.SecTypeFuture Or _
            lSecType = SecurityTypes.SecTypeOption Or _
            lSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (lCurrContract.ExpiryDate - lCurrContract.DaysBeforeExpiryToSwitch) _
            Then
                Do While en.MoveNext
                    Set aContract = en.Current
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
                    Then
                        tfs.ToDate = currSessionStart
                        tfs.Filename = generateNonSessionFilename(lCurrContract, tfs.FromDate, tfs.ToDate)
                        
                        Set lCurrContract = aContract
                        
                        Set tfs = New TickfileSpecifier
                        lTickFileSpecifiers.Add tfs
                        
                        tfs.Contract = lCurrContract
                        tfs.TickfileFormatID = pTickfileFormatID
                    
                        tfs.FromDate = currSessionStart
                        Exit Do
                    End If
                Loop
                AssertArgument Not lCurrContract Is Nothing, "No contract from " & currSessionStart
            End If
        End If
        
        currSessionStart = currSessionStart + 1
        
    Loop
        
    tfs.ToDate = replayToDate
    tfs.Filename = generateNonSessionFilename(lCurrContract, tfs.FromDate, tfs.ToDate)

End If

Set GenerateTickfileSpecifiers = lTickFileSpecifiers

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

Private Function editContractSessionTimes( _
                ByVal pContract As IContract, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As IContract
Const ProcName As String = "editContractSessionTimes"
On Error GoTo Err

Dim lContractBuilder As ContractBuilder

Set lContractBuilder = CreateContractBuilderFromContract(pContract)
lContractBuilder.sessionEndTime = sessionEndTime
lContractBuilder.sessionStartTime = sessionStartTime
Set editContractSessionTimes = lContractBuilder.Contract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateNonSessionFilename( _
                ByVal pContract As IContract, _
                ByVal pFromDate As Date, _
                ByVal pToDate As Date) As String
Const ProcName As String = "generateNonSessionFilename"
On Error GoTo Err

generateNonSessionFilename = pContract.Specifier.LocalSymbol & " (" & _
                        FormatDateTime(pFromDate, vbGeneralDate) & _
                        "-" & _
                        FormatDateTime(pToDate, vbGeneralDate) & _
                        ")"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateSessionFilename( _
                ByVal pContract As IContract, _
                ByVal pSessionStart As Date) As String
Const ProcName As String = "generateSessionFilename"
On Error GoTo Err

generateSessionFilename = pContract.Specifier.LocalSymbol & " (" & _
                        "Session: " & _
                        FormatDateTime(DateValue(pSessionStart), vbShortDate) & _
                        ")"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


