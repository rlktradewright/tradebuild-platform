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

Public Function CreateTickfileManager( _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pReplaySpeed As Long = 1, _
                Optional ByVal pReplayProgressEventInterval As Long = 1000, _
                Optional ByVal pTimestampAdjustmentStart As Double = 0#, _
                Optional ByVal pTimestampAdjustmentEnd As Double = 0#) As TickFileManager
Const ProcName As String = "CreateTickfileManager"
On Error GoTo Err

AssertArgument Not pTickfileStore Is Nothing, "pTickfileStore cannot be Nothing"
AssertArgument pReplayProgressEventInterval >= 50, "pReplayProgressEventInterval cannot be less than 50"

Set CreateTickfileManager = New TickFileManager
CreateTickfileManager.Initialise pTickfileStore, pPrimaryContractStore, pSecondaryContractStore, pReplaySpeed, pTimestampAdjustmentStart, pTimestampAdjustmentEnd, pReplayProgressEventInterval

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

Dim useCustomSessionTimes As Boolean
Dim lTickfileSpecifiers As New TickFileSpecifiers
Dim tfs As TickfileSpecifier
Dim lContract As IContract

Dim lSecType As SecurityTypes
Dim currContract As IContract
Dim aContract As IContract
Dim lContractTimezone As TimeZone

' the from and to dates converted (if necessary) to the contract's timezone
Dim FromDate As Date
Dim ToDate As Date

' the from and to dates, Session-oriented if required
Dim replayFromDate As Date
Dim replayToDate As Date


AssertArgument (pCustomSessionStartTime = 0 And pCustomSessionEndTime = 0) Or _
    (pCustomSessionStartTime <> 0 And pCustomSessionEndTime <> 0), "Session start and end times must be either both supplied or both omitted"
    
If pCustomSessionStartTime <> 0 Then useCustomSessionTimes = True

' get the most recent contract (though they should all have the same
' info regarding Session times)
Set lContract = pContracts.ItemAtIndex(pContracts.Count)
Set lContractTimezone = GetTimeZone(lContract.TimezoneName)

If pUseExchangeTimezone Then
    FromDate = pStartDate
    If pEndDate <> 0 Then ToDate = pEndDate
Else
    FromDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pStartDate), lContractTimezone)
    If pEndDate <> 0 Then ToDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(pEndDate), lContractTimezone)
End If

' determine start and end times ---------------------------------------------------

Dim lSessionBuilder As SessionBuilder
Set lSessionBuilder = New SessionBuilder

' note that the custom times are in the contract's timezone
If useCustomSessionTimes Then
    lSessionBuilder.sessionStartTime = pCustomSessionStartTime
    lSessionBuilder.sessionEndTime = pCustomSessionEndTime
Else
    lSessionBuilder.sessionStartTime = lContract.sessionStartTime
    lSessionBuilder.sessionEndTime = lContract.sessionEndTime
End If
lSessionBuilder.CompleteSessionCreation

Dim lSession As Session
Set lSession = lSessionBuilder.SessionFuture.Value

' set the Session start and end times for the starting date (in the contract's
' local timezone)...
Dim fromSessionTimes As SessionTimes
fromSessionTimes = lSession.GetSessionTimes(FromDate)
If pCompleteSessionsOnly Then
    replayFromDate = fromSessionTimes.StartTime
Else
    replayFromDate = FromDate
End If

Dim toSessionTimes As SessionTimes
If ToDate <> 0 Then
    ' set the Session start and end times for the ending date (in the contract's
    ' local timezone)...
    toSessionTimes = lSession.GetSessionTimes(ToDate)
    
    If pCompleteSessionsOnly Then
        replayToDate = toSessionTimes.EndTime
    Else
        replayToDate = ToDate
    End If
Else
    toSessionTimes.StartTime = fromSessionTimes.StartTime
    toSessionTimes.EndTime = fromSessionTimes.EndTime
    replayToDate = toSessionTimes.EndTime
End If

' find contract for start date ------------------------------------------------------

Dim en As Enumerator
Set en = pContracts.Enumerator

lSecType = pContracts.ItemAtIndex(1).Specifier.SecType

If lSecType <> SecurityTypes.SecTypeFuture And _
    lSecType <> SecurityTypes.SecTypeOption And _
    lSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    Set currContract = pContracts.ItemAtIndex(1)
Else
    Do While en.MoveNext
    Set aContract = en.Current
        If DateValue(FromDate) <= _
            (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
        Then
            Set currContract = aContract
            Exit Do
        End If
    Loop
    
    AssertArgument Not currContract Is Nothing, "No contract for this from date"
End If

If pCompleteSessionsOnly Then
    If useCustomSessionTimes Then
        Set currContract = editContractSessionTimes(currContract, pCustomSessionStartTime, pCustomSessionEndTime)
    End If
End If

Dim currSessionStart As Date
Dim thisSessionTimes As SessionTimes

currSessionStart = FromDate

If pCompleteSessionsOnly Then
    Do While currSessionStart < replayToDate
        Set tfs = New TickfileSpecifier
        lTickfileSpecifiers.Add tfs
        tfs.ContractFuture = CreateFuture(currContract)
        tfs.TickfileFormatID = pTickfileFormatID
        
        If useCustomSessionTimes Then
            thisSessionTimes = lSession.GetSessionTimes(currSessionStart)
            tfs.FromDate = thisSessionTimes.StartTime
            tfs.ToDate = thisSessionTimes.EndTime
            tfs.EntireSession = False
            tfs.Filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(tfs.ToDate, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.Specifier.ToString, vbCrLf, "; ")
        Else
            tfs.FromDate = currSessionStart
            tfs.EntireSession = True
            tfs.Filename = "Session " & _
                            FormatDateTime(DateValue(currSessionStart), vbShortDate) & _
                            " " & _
                            Replace(currContract.Specifier.ToString, vbCrLf, "; ")
            
        End If
        
        currSessionStart = currSessionStart + 1
        
        If lSecType = SecurityTypes.SecTypeFuture Or _
            lSecType = SecurityTypes.SecTypeOption Or _
            lSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.ExpiryDate - currContract.DaysBeforeExpiryToSwitch) _
            Then
                Do While en.MoveNext
                    Set aContract = en.Current
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
                    Then
                        Set currContract = aContract
                        If useCustomSessionTimes Then
                            Set currContract = editContractSessionTimes(currContract, _
                                                                        pCustomSessionStartTime, _
                                                                        pCustomSessionEndTime)
                        End If
                        Exit Do
                    End If
                Loop
                AssertArgument Not currContract Is Nothing, "No contract from " & currSessionStart
            End If
        End If
        
        Loop
    AssertArgument lTickfileSpecifiers.Count <> 0, "No trading sessions in specified date range"
Else
    Set tfs = New TickfileSpecifier
    lTickfileSpecifiers.Add tfs
    tfs.ContractFuture = CreateFuture(currContract)
    tfs.TickfileFormatID = pTickfileFormatID

    tfs.FromDate = FromDate
    currSessionStart = currSessionStart + 1
    Do While currSessionStart < replayToDate
        
        If lSecType = SecurityTypes.SecTypeFuture Or _
            lSecType = SecurityTypes.SecTypeOption Or _
            lSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.ExpiryDate - currContract.DaysBeforeExpiryToSwitch) _
            Then
                Do While en.MoveNext
                    Set aContract = en.Current
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.DaysBeforeExpiryToSwitch) _
                    Then
                        tfs.ToDate = currSessionStart
                        tfs.Filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(tfs.ToDate, vbGeneralDate) & " " & _
                                                    Replace(currContract.Specifier.ToString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        Set tfs = New TickfileSpecifier
                        lTickfileSpecifiers.Add tfs
                        
                        tfs.ContractFuture = CreateFuture(currContract)
                        tfs.TickfileFormatID = pTickfileFormatID
                    
                        tfs.FromDate = currSessionStart
                        Exit Do
                    End If
                Loop
                AssertArgument Not currContract Is Nothing, "No contract from " & currSessionStart
            End If
        End If
        
        currSessionStart = currSessionStart + 1
        
    Loop
        
    tfs.ToDate = replayToDate
    tfs.Filename = FormatDateTime(tfs.FromDate, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(tfs.ToDate, vbGeneralDate) & " " & _
                                Replace(currContract.Specifier.ToString, vbCrLf, "; ")

End If

Set GenerateTickfileSpecifiers = lTickfileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RecordTickData( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTickfileStore As ITickfileStore, _
                Optional ByVal pOutputMonitor As ITickfileOutputMonitor, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As TickDataWriter
Const ProcName As String = "RecordTickData"
On Error GoTo Err

Set RecordTickData = New TickDataWriter
RecordTickData.Initialise pDataSource, pOutputMonitor, pTickfileStore, pFormatIdentifier, pLocation

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




