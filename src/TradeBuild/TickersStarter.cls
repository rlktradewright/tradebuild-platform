VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickersStarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
' @see
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

Private Const ProjectName                   As String = "TradeBuild26"
Private Const ModuleName                    As String = "TickersStarter"

'@================================================================================
' Member variables
'@================================================================================

' need this to keep alive - nothing else has a reference to me
Private mSelfRef                            As TickersStarter

Private WithEvents mContracts               As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mTickersRef                         As WeakReference
Private mTickerOptions                      As Long
Private mDOMEventsRequired                  As DOMEvents
Private mOutputTickfileFormat               As String
Private mOutputTickfilePath                 As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSelfRef = Me
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_CollectionChanged( _
                ev As TWUtilities30.CollectionChangeEvent)
Dim lContract As Contract
Dim lTicker As Ticker

Set lContract = ev.affectedItem

'check contract is still current
If lContract.expiryDate <> 0 And lContract.expiryDate < Date Then Exit Sub

Set lTicker = tckrs.add(mTickerOptions)
lTicker.DOMEventsRequired = mDOMEventsRequired
lTicker.outputTickfileFormat = mOutputTickfileFormat
lTicker.outputTickfilePath = mOutputTickfilePath

lTicker.StartTickerFromContract lContract
End Sub

Private Sub mContracts_ContractSpecifierInvalid( _
                ByVal reason As String)
Dim ev As NotificationEvent

Set ev.source = mContracts.contractSpecifier
ev.eventCode = ApiNotifyCodes.ApiNotifyContractSpecifierInvalid
ev.eventMessage = reason & ": " & mContracts.contractSpecifier.toString
tckrs.notify ev
End Sub

Private Sub mContracts_NoMoreContractDetails()
' all tickers have been started so let ourself die
Set mSelfRef = Nothing
End Sub

Private Sub mContracts_Notification( _
                ev As TWUtilities30.NotificationEvent)
' don't know what we can do to handle this - should log it perhaps?
End Sub

Private Sub mContracts_TryLater( _
                ByVal reason As String, _
                ByVal serviceProviderName As String)
' don't know what we can do to handle this - should log it perhaps?
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub startTickers( _
                ByVal pTickers As Tickers, _
                ByVal pTickerOptions As Long, _
                ByVal pContractSpec As contractSpecifier, _
                Optional ByVal pDOMEventsRequired As DOMEvents, _
                Optional ByVal poutputTickfileFormat As String, _
                Optional ByVal poutputTickfilePath As String)
                
Set mTickersRef = CreateWeakReference(pTickers)
mTickerOptions = pTickerOptions
mDOMEventsRequired = pDOMEventsRequired
mOutputTickfileFormat = poutputTickfileFormat
mOutputTickfilePath = poutputTickfilePath

Set mContracts = gTB.loadContracts(pContractSpec)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function tckrs() As Tickers
Set tckrs = mTickersRef.Target
End Function
