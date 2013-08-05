VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataSourceSettings"
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

Private Const ModuleName                            As String = "MarketDataSourceSettings"

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataSourceOptions                    As MarketDataSourceOptions
Private mNumberOfMarketDepthRows                    As Long

Private mConfig                                     As ConfigurationSection

Private mDefaultStateChangeListener                 As StateChangeListener

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mMarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone
mNumberOfMarketDepthRows = 20
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ConfigurationSection() As ConfigurationSection
Set ConfigurationSection = mConfig
End Property

Public Property Let DefaultStateChangeListener(ByVal Value As StateChangeListener)
Set mDefaultStateChangeListener = Value
End Property

Public Property Get DefaultStateChangeListener() As StateChangeListener
Set DefaultStateChangeListener = mDefaultStateChangeListener
End Property

Public Property Let MarketDataSourceOptions(ByVal Value As MarketDataSourceOptions)
Const ProcName As String = "MarketDataSourceOptions"
On Error GoTo Err

mMarketDataSourceOptions = Value

If mConfig Is Nothing Then Exit Property

mConfig.SetSetting ConfigSettingNoImpliedTrades, CStr((mMarketDataSourceOptions And MarketDataSourceOptNoImpliedTrades) = MarketDataSourceOptNoImpliedTrades)
mConfig.SetSetting ConfigSettingNoVolumeAdjustments, CStr((mMarketDataSourceOptions And MarketDataSourceOptNoVolumeAdjustments) = MarketDataSourceOptNoVolumeAdjustments)
mConfig.SetSetting ConfigSettingUseExchangeTimezone, CStr((mMarketDataSourceOptions And MarketDataSourceOptUseExchangeTimeZone) = MarketDataSourceOptUseExchangeTimeZone)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get MarketDataSourceOptions() As MarketDataSourceOptions
MarketDataSourceOptions = mMarketDataSourceOptions
End Property

Public Property Let NumberOfMarketDepthRows(ByVal Value As Long)
Const ProcName As String = "NumberOfMarketDepthRows"
On Error GoTo Err

mNumberOfMarketDepthRows = Value

If mConfig Is Nothing Then Exit Property

mConfig.SetSetting ConfigSettingNumberOfMarketDepthRows, CStr(mNumberOfMarketDepthRows)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NumberOfMarketDepthRows() As Long
NumberOfMarketDepthRows = mNumberOfMarketDepthRows
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = pConfig.AddConfigurationSection(ConfigSectionMarketDataSources)

Dim lOptions As MarketDataSourceOptions
If CBool(mConfig.GetSetting(ConfigSettingNoImpliedTrades, "False")) Then lOptions = MarketDataSourceOptNoImpliedTrades
If CBool(mConfig.GetSetting(ConfigSettingNoVolumeAdjustments, "False")) Then lOptions = lOptions Or MarketDataSourceOptNoVolumeAdjustments
If CBool(mConfig.GetSetting(ConfigSettingUseExchangeTimezone, "True")) Then lOptions = lOptions Or MarketDataSourceOptUseExchangeTimeZone

mMarketDataSourceOptions = lOptions

mNumberOfMarketDepthRows = CLng(mConfig.GetSetting(ConfigSettingNumberOfMarketDepthRows, "20"))
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




