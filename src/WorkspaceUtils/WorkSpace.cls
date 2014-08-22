VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "WorkSpace"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName                           As String

Private mTickers                        As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mConfig                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "Workspace " & mName & " terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Name() As String
Name = mName
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pScopeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistoricalDataStore As IHistoricalDataStore, _
                ByVal pOrderManager As OrderManager, _
                ByVal pOrderSubmitterFactoryLive As IOrderSubmitterFactory, _
                ByVal pOrderSubmitterFactorySimulated As IOrderSubmitterFactory, _
                ByVal pMoneyManager As IMoneyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

mName = pName
Set mTickers = CreateTickers(pMarketDataManager, _
                            pStudyLibraryManager, _
                            pHistoricalDataStore, _
                            pOrderManager, _
                            pOrderSubmitterFactoryLive, _
                            pOrderSubmitterFactorySimulated, _
                            pName, _
                            pScopeName, _
                            pMoneyManager)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = pConfig

mTickers.LoadFromConfig mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


