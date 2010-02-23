Attribute VB_Name = "Globals"
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                               As String = "TBDataCollector26"
Private Const ModuleName                                As String = "Globals"

Public Const AttributeNameBidAskBars                    As String = "WriteBidAndAskBars"
Public Const AttributeNameEnabled                       As String = "Enabled"
Public Const AttributeNameIncludeMktDepth               As String = "IncludeMarketDepth"

Public Const ConfigSectionCollectionControl             As String = "CollectionControl"
Public Const ConfigSectionContract                      As String = "Contract"
Public Const ConfigSectionContracts                     As String = "Contracts"
Public Const ConfigSectionContractspecifier             As String = "ContractSpecifier"
Public Const ConfigSectionTickdata                      As String = "TickData"

Public Const ConfigSettingContractSpecCurrency          As String = ConfigSectionContractspecifier & "&Currency"
Public Const ConfigSettingContractSpecExpiry            As String = ConfigSectionContractspecifier & "&Expiry"
Public Const ConfigSettingContractSpecExchange          As String = ConfigSectionContractspecifier & "&Exchange"
Public Const ConfigSettingContractSpecLocalSYmbol       As String = ConfigSectionContractspecifier & "&LocalSymbol"
Public Const ConfigSettingContractSpecRight             As String = ConfigSectionContractspecifier & "&Right"
Public Const ConfigSettingContractSpecSecType           As String = ConfigSectionContractspecifier & "&SecType"
Public Const ConfigSettingContractSpecStrikePrice       As String = ConfigSectionContractspecifier & "&StrikePrice"
Public Const ConfigSettingContractSpecSymbol            As String = ConfigSectionContractspecifier & "&Symbol"

Public Const ConfigFileVersion                          As String = "1.0"

Public Const ConfigSettingWriteBarData                  As String = ConfigSectionCollectionControl & ".WriteBarData"
Public Const ConfigSettingWriteTickData                 As String = ConfigSectionCollectionControl & ".WriteTickData"
Public Const ConfigSettingWriteTickDataFormat           As String = ConfigSectionTickdata & ".Format"
Public Const ConfigSettingWriteTickDataPath             As String = ConfigSectionTickdata & ".Path"

'@================================================================================
' Member variables
'@================================================================================

Private mLogger                             As Logger

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

Public Property Get gLogger() As Logger
Set gLogger = mLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Sub main()
Set mLogger = GetLogger("")
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


