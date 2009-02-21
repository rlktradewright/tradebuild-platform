VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartTickersTask"
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

Implements Task

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

Private Const ProjectName                   As String = "TBDataCollector26"
Private Const ModuleName                    As String = "StartTickersTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                    As TaskContext

Private mContractsConfig                                As ConfigurationSection

Private mTickers                                        As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mWriteticks                                     As Boolean
Private mWritebars                                      As Boolean

Private mOutputFormat                                   As String
Private mOutputPath                                     As String

Private mEnumerator                                     As Enumerator

Private WithEvents mTickersTC                           As TaskController
Attribute mTickersTC.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_cancel()

End Sub

Private Sub Task_run()
Dim contractSpec As contractSpecifier
Dim contractConfig As ConfigurationSection
Dim contractSpecConfig As ConfigurationSection

Dim localSymbol As String
Dim symbol As String
Dim exchange As String
Dim sectype As SecurityTypes
Dim currencyCode As String
Dim expiry As String
Dim strikePrice As Double
Dim optRight As OptionRights

Dim writeBidAskBars As Boolean
Dim includeMarketDepth As Boolean

Dim failpoint As Long
On Error GoTo Err

If Not mEnumerator.moveNext Then
    mTaskContext.finish "Yield count =" & mTaskContext.yieldCount, False
    Exit Sub
End If

Set contractConfig = mEnumerator.current

failpoint = 100 '=====================================================

If CBool(contractConfig.getAttribute(AttributeNameEnabled)) Then

    includeMarketDepth = CBool(contractConfig.getAttribute(AttributeNameIncludeMktDepth, "False"))
    writeBidAskBars = CBool(contractConfig.getAttribute(AttributeNameBidAskBars, "False"))

    failpoint = 200 '=====================================================
    
    With contractConfig
        localSymbol = .GetSetting(ConfigSettingContractSpecLocalSYmbol, "")
        symbol = .GetSetting(ConfigSettingContractSpecSymbol, "")
        exchange = .GetSetting(ConfigSettingContractSpecExchange, "")
        sectype = SecTypeFromString(.GetSetting(ConfigSettingContractSpecSecType, ""))
        currencyCode = .GetSetting(ConfigSettingContractSpecCurrency, "")
        expiry = .GetSetting(ConfigSettingContractSpecExpiry, "")
        strikePrice = CDbl("0" & .GetSetting(ConfigSettingContractSpecStrikePrice, "0.0"))
        optRight = OptionRightFromString(.GetSetting(ConfigSettingContractSpecRight, ""))
        
        Set contractSpec = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                sectype, _
                                                currencyCode, _
                                                expiry, _
                                                strikePrice, _
                                                optRight)
    End With
    
    failpoint = 400 '=====================================================
    
    gLogger.Log LogLevelNormal, mTaskContext.Name & ": starting tickers for: " & contractSpec.toString
    
    Set mTickersTC = mTickers.startTickers( _
                    TickerOptions.TickerOptNoImpliedTrades + _
                    TickerOptions.TickerOptNoVolumeAdjustments + _
                    IIf(mWritebars, TickerOptions.TickerOptWriteTradeBarData, 0) + _
                    IIf(mWritebars And writeBidAskBars, _
                        TickerOptions.TickerOptWriteBidAndAskBarData, 0) + _
                    IIf(mWritebars, TickerOptions.TickerOptUpdateLatestBar, 0) + _
                    IIf(mWriteticks, TickerOptions.TickerOptWriteTickData, 0) + _
                    IIf(mWriteticks And includeMarketDepth, _
                        TickerOptions.TickerOptIncludeMarketDepthInTickfile, 0), _
                    contractSpec, _
                    IIf(mWriteticks And includeMarketDepth, DOMEvents.DOMRawEvents, DOMNoEvents), _
                    mOutputFormat, _
                    mOutputPath, _
                    1000)
    
End If

mTaskContext.suspend -1 ' wait till this set of tickers is started to avoid creating too many database connections

Exit Sub

Err:
mTaskContext.Error Err.Number, _
                    Err.description, _
                    ProjectName & "." & ModuleName & ":" & "Task_run" & "." & failpoint & _
                    IIf(Err.Source <> "", vbCrLf & Err.Source, "")

End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
End Property

'@================================================================================
' mTickersTC Event Handlers
'@================================================================================

Private Sub mTickersTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.errorNumber <> 0 Then gLogger.Log LogLevelSevere, "Error " & ev.errorNumber & " starting tickers: " & ev.ErrorMessage
mTaskContext.continue
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise( _
                ByVal pTickers As Tickers, _
                ByVal contractsConfig As ConfigurationSection, _
                ByVal writeBars As Boolean, _
                ByVal writeTicks As Boolean, _
                ByVal outputFormat As String, _
                ByVal outputPath As String)
Set mTickers = pTickers
Set mContractsConfig = contractsConfig
mWritebars = writeBars
mWriteticks = writeTicks
mOutputFormat = outputFormat
mOutputPath = outputPath

Set mEnumerator = mContractsConfig.ChildEnumerator
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


