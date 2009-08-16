VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractRequester"
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

Implements ContractDetailsRequester

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

Private Const ModuleName                            As String = "ContractRequester"

'@================================================================================
' Member variables
'@================================================================================

Private mContractInfoServiceProvider As ContractInfoServiceProvider
Private mCommonServiceConsumer As ICommonServiceConsumer
Private mHandle As Long
Private mContractInfoServiceConsumer As IContractInfoServiceConsumer

Private mName As String

Private mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mNumOutstandingRequests As Long

Private mContractsBuilder As ContractsBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

'================================================================================
' ContractDetailsRequester Interface Members
'================================================================================

Public Sub ContractDetailsRequester_cancelRequest( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
On Error GoTo err
mNumOutstandingRequests = mNumOutstandingRequests - 1
mContractsBuilder.unableToLoad reason, mName
checkIfFinished
Exit Sub

err:
handleFatalError err.number, _
                "ContractDetailsRequester_cancelRequest", _
                err.Description
End Sub

Private Sub ContractDetailsRequester_contractsLoaded( _
                ByRef request As TWSContractDetailsRequestQueueEntry)
Dim lContractWrapper As ContractWrapper

On Error GoTo err
mNumOutstandingRequests = mNumOutstandingRequests - 1
For Each lContractWrapper In request.contractWrappers
    mContractsBuilder.AddContract lContractWrapper.tradeBuildContract
Next
If mNumOutstandingRequests = 0 Then mContractInfoServiceConsumer.LoadComplete mHandle
checkIfFinished
Exit Sub
err:
handleFatalError err.number, _
                "ContractDetailsRequester_contractsLoaded", _
                err.Description
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
On Error GoTo err
mNumOutstandingRequests = mNumOutstandingRequests - 1
If mNumOutstandingRequests = 0 Then
    If mContractsBuilder.Contracts.Count <> 0 Then
        mContractInfoServiceConsumer.LoadComplete mHandle
    Else
        mContractInfoServiceConsumer.notifyEvent StandardSPEventCodes.CIContractSpecifierInvalid, _
                                            reason, _
                                            mHandle
    End If
End If
checkIfFinished
Exit Sub

err:
handleFatalError err.number, _
                "ContractDetailsRequester_contractSpecifierInvalid", _
                err.Description
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub finish()
Set mContractInfoServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing
Set mContractsBuilder = Nothing
Set mTWSAPI = Nothing
End Sub

Friend Sub Initialise( _
                ByVal pContractInfoServiceProvider As ContractInfoServiceProvider, _
                ByVal commonServiceConsumer As ICommonServiceConsumer, _
                ByVal ContractInfoServiceConsumer As IContractInfoServiceConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal pContractsBuilder As ContractsBuilder, _
                ByVal name As String, _
                ByVal handle As Long)

Dim contractSpec As ContractSpecifier

mName = name
mHandle = handle

Set mContractInfoServiceProvider = pContractInfoServiceProvider
Set mCommonServiceConsumer = commonServiceConsumer
Set mContractInfoServiceConsumer = ContractInfoServiceConsumer

Set mContractsBuilder = pContractsBuilder
Set contractSpec = mContractsBuilder.Contracts.ContractSpecifier

Set mTWSAPI = pTWSAPI

If contractSpec.sectype = SecTypeNone Then
    ' no sectype supplied, so we'll place a separate request for each of the following:
    '       Futures
    '       Indexes
    '       Stocks
    mTWSAPI.RequestContract CreateContractSpecifier(contractSpec.localSymbol, _
                                                    contractSpec.symbol, _
                                                    contractSpec.exchange, _
                                                    SecTypeFuture, _
                                                    contractSpec.currencyCode, _
                                                    contractSpec.expiry), _
                           Me
    mNumOutstandingRequests = 1
    
    If contractSpec.localSymbol = "" Then
        mTWSAPI.RequestContract CreateContractSpecifier(contractSpec.symbol, _
                                                        "", _
                                                        contractSpec.exchange, _
                                                        SecTypeFuture, _
                                                        contractSpec.currencyCode, _
                                                        contractSpec.expiry), _
                               Me
        mNumOutstandingRequests = mNumOutstandingRequests + 1
    End If
    
    mTWSAPI.RequestContract CreateContractSpecifier(contractSpec.localSymbol, _
                                                    contractSpec.symbol, _
                                                    contractSpec.exchange, _
                                                    SecTypeIndex, _
                                                    contractSpec.currencyCode, _
                                                    contractSpec.expiry), _
                           Me
    mNumOutstandingRequests = mNumOutstandingRequests + 1

    mTWSAPI.RequestContract CreateContractSpecifier(contractSpec.localSymbol, _
                                                    contractSpec.symbol, _
                                                    contractSpec.exchange, _
                                                    SecTypeStock, _
                                                    contractSpec.currencyCode, _
                                                    contractSpec.expiry), _
                           Me
    mNumOutstandingRequests = mNumOutstandingRequests + 1

    mTWSAPI.RequestContract CreateContractSpecifier(contractSpec.localSymbol, _
                                                    contractSpec.symbol, _
                                                    contractSpec.exchange, _
                                                    SecTypeCash, _
                                                    contractSpec.currencyCode, _
                                                    contractSpec.expiry), _
                           Me
    mNumOutstandingRequests = mNumOutstandingRequests + 1

Else
    mTWSAPI.RequestContract contractSpec, _
                            Me
    mNumOutstandingRequests = 1
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfFinished()
If mNumOutstandingRequests = 0 Then
    mContractInfoServiceProvider.contractRequesterFinished Me
    finish
End If
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
logMessage "Error " & err.number & ": " & _
            ProjectName & "." & ModuleName & ":" & source & vbCrLf & _
            err.Description, _
            LogLevelSevere

gReleaseAllTWSAPIInstances

mCommonServiceConsumer.NotifyFatalError number, _
                            source, _
                            Description, _
                            mHandle

End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            ProjectName & "." & ModuleName & ": " & message
End Sub


