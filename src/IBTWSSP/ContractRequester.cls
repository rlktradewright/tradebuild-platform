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
Const ProcName As String = "ContractDetailsRequester_cancelRequest"
On Error GoTo Err
mNumOutstandingRequests = mNumOutstandingRequests - 1
mContractsBuilder.UnableToLoad reason, mName
checkIfFinished
Exit Sub

Err:
handleFatalError ProcName
End Sub

Private Sub ContractDetailsRequester_contractsLoaded( _
                ByRef request As TWSContractDetailsRequestQueueEntry)
Const ProcName As String = "ContractDetailsRequester_contractsLoaded"
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
handleFatalError ProcName
End Sub

Private Sub ContractDetailsRequester_contractSpecifierInvalid( _
                ByRef request As TWSContractDetailsRequestQueueEntry, _
                ByVal reason As String)
Const ProcName As String = "ContractDetailsRequester_contractSpecifierInvalid"
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
handleFatalError ProcName
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
Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Set mContractInfoServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mContractInfoServiceConsumer = Nothing
Set mContractsBuilder = Nothing
Set mTWSAPI = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractInfoServiceProvider As ContractInfoServiceProvider, _
                ByVal CommonServiceConsumer As ICommonServiceConsumer, _
                ByVal ContractInfoServiceConsumer As IContractInfoServiceConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal pContractsBuilder As ContractsBuilder, _
                ByVal Name As String, _
                ByVal Handle As Long)
Const ProcName As String = "Initialise"


Dim contractSpec As ContractSpecifier

Dim failpoint As Long
On Error GoTo Err

mName = Name
mHandle = Handle

Set mContractInfoServiceProvider = pContractInfoServiceProvider
Set mCommonServiceConsumer = CommonServiceConsumer
Set mContractInfoServiceConsumer = ContractInfoServiceConsumer

Set mContractsBuilder = pContractsBuilder
Set contractSpec = mContractsBuilder.Contracts.ContractSpecifier

Set mTWSAPI = pTWSAPI

If contractSpec.sectype = SecTypeNone Then
    ' no sectype supplied, so we'll place a separate request for each of the following:
    '       Futures
    '       Indexes
    '       Stocks
    '       Cash
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkIfFinished()
Const ProcName As String = "checkIfFinished"
Dim failpoint As Long
On Error GoTo Err

If mNumOutstandingRequests = 0 Then
    mContractInfoServiceProvider.contractRequesterFinished Me
    Finish
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

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


