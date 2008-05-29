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

Private mContractsConfig                                As ConfigItem

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
Dim contractConfig As ConfigItem
Dim contractSpecConfig As ConfigItem

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

Dim failpoint As Integer


On Error GoTo Err

For Each contractConfig In mContractsConfig.childItems

    failpoint = 100 '=====================================================

    If CBool(contractConfig.getAttribute(AttributeNameEnabled)) Then
    
        includeMarketDepth = CBool(contractConfig.getDefaultableAttribute(AttributeNameIncludeMktDepth, "False"))
        writeBidAskBars = CBool(contractConfig.getDefaultableAttribute(AttributeNameBidAskBars, "False"))

        failpoint = 200 '=====================================================
        
        Set contractSpecConfig = contractConfig.childItems.Item(ConfigNameContractSpecifier)
        
        failpoint = 300 '=====================================================
        
        With contractSpecConfig
            localSymbol = .getDefaultableAttribute(AttributeNameLocalSYmbol, "")
            symbol = .getDefaultableAttribute(AttributeNameSymbol, "")
            exchange = .getDefaultableAttribute(AttributeNameExchange, "")
            sectype = SecTypeFromString(.getDefaultableAttribute(AttributeNameSecType, ""))
            currencyCode = .getDefaultableAttribute(AttributeNameCurrency, "")
            expiry = .getDefaultableAttribute(AttributeNameExpiry, "")
            strikePrice = CDbl("0" & .getDefaultableAttribute(AttributeNameStrikePrice, "0.0"))
            optRight = OptionRightFromString(.getDefaultableAttribute(AttributeNameRight, ""))
            
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
        
        mTickers.startTickers _
                        IIf(mWriteBars, TickerOptions.TickerOptWriteTradeBarData, 0) + _
                        IIf(mWriteBars And writeBidAskBars, _
                            TickerOptions.TickerOptWriteBidAndAskBarData, 0) + _
                        IIf(mWriteBars, TickerOptions.TickerOptUpdateLatestBar, 0) + _
                        IIf(mWriteTicks, TickerOptions.TickerOptWriteTickData, 0) + _
                        IIf(mWriteTicks And includeMarketDepth, _
                            TickerOptions.TickerOptIncludeMarketDepthInTickfile, 0), _
                        contractSpec, _
                        IIf(mWriteTicks And includeMarketDepth, DOMEvents.DOMRawEvents, DOMNoEvents), _
                        mOutputFormat, _
                        mOutputPath
        
    End If
Next

Exit Sub

Err:
MsgBox "Error " & Err.Number & ":" & vbCrLf & _
        "Failpoint: " & failpoint & vbCrLf & _
        Err.description & vbCrLf & _
        Err.Source, , "An error has occurred"

End Sub

Private Property Let Task_TaskContext(ByVal value As TWUtilities30.TaskContext)
Set mTaskContext = value
Set mContractsConfig = mTaskContext.Data
End Property

Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


