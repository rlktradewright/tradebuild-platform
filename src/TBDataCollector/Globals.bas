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

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Sub main()
Set mLogger = GetLogger("")
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


