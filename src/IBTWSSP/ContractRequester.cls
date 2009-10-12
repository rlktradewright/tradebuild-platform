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
On Error GoTo Err
mNumOutstandingRequests = mNumOutstandingRequests - 1
mContractsBuilder.unableToLoad reason, mName
checkIfFinished
Exit Sub

Err:
handleFatalError "ContractDetailsRequester_cancelRequest"
End Sub

Private Sub ContractDetailsRequester_contractsLoaded( _
                ByRef request As TWSContractDetailsRequestQueueEntry)
Dim lContractWrapper As ContractWrapper

On Error GoTo Err
mNumOutstandingRequests = mNumOutstandingRequests - 1
For Each lContractWrapper In request.contractWrappers
    mContractsBuilder.AddContract lContractWrapper.tradeBuildContract
Next
If mNumOutstandingRequests = 0 Then mContractInfoServiceConsumer.LoadComplete mHandle
checkIfFinished
Exit Sub
Err:
handleFatalError "ContractDetailsRequester_contractsLoaded"
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
On Error GoTo Err
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

Err:
handleFatalError "ContractDetailsRequester_contractSpecifierInvalid"
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

Friend Sub Finish()
Dim failpoint As Long
On Error GoTo Err

Set mContractInfoServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing
Set mContractsBuilder = Nothing
Set mTWSAPI = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Finish", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub Initialise( _
                ByVal pContractInfoServiceProvider As ContractInfoServiceProvider, _
                ByVal commonServiceConsumer As ICommonServiceConsumer, _
                ByVal ContractInfoServiceConsumer As IContractInfoServiceConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal pContractsBuilder As ContractsBuilder, _
                ByVal Name As String, _
                ByVal Handle As Long)

Dim contractSpec As ContractSpecifier

Dim failpoint As Long
On Error GoTo Err

mName = Name
mHandle = Handle

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Initialise", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfFinished()
Dim failpoint As Long
On Error GoTo Err

If mNumOutstandingRequests = 0 Then
    mContractInfoServiceProvider.contractRequesterFinished Me
    Finish
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="checkIfFinished", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
HandleUnexpectedError pReRaise:=True, pLog:=True, pProcedureName:=pProcName, pNumber:=errNum, pSource:=errSource, pDescription:=errDesc, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=pFailpoint

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub


