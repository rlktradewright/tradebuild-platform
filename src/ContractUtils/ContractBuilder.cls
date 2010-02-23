VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
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

Private Const ModuleName                    As String = "ContractBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mContract                           As Contract

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

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Let DaysBeforeExpiryToSwitch(ByVal value As Long)
mContract.DaysBeforeExpiryToSwitch = value
End Property

Public Property Let Description(ByVal value As String)
mContract.Description = value
End Property

Public Property Let ExpiryDate(ByVal value As Date)
mContract.ExpiryDate = value
End Property

Public Property Let Multiplier(ByVal value As Double)
mContract.Multiplier = value
End Property

Public Property Let ProviderIDs(ByRef value As Parameters)
mContract.ProviderIDs = value
End Property

Public Property Let SessionStartTime(ByVal value As Date)
mContract.SessionStartTime = TimeValue(value)
End Property

Public Property Let SessionEndTime(ByVal value As Date)
mContract.SessionEndTime = TimeValue(value)
End Property

Public Property Let TickSize(ByVal value As Double)
mContract.TickSize = value
End Property

Public Property Let Timezone( _
                ByVal value As Timezone)
mContract.Timezone = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub BuildFrom( _
                ByVal pContract As Contract)
Const ProcName As String = "BuildFrom"
Dim failpoint As String
On Error GoTo Err

With mContract
    .DaysBeforeExpiryToSwitch = pContract.DaysBeforeExpiryToSwitch
    .Description = pContract.Description
    .ExpiryDate = pContract.ExpiryDate
    .Multiplier = pContract.Multiplier
    .ProviderIDs = pContract.ProviderIDs
    .SessionEndTime = pContract.SessionEndTime
    .SessionStartTime = pContract.SessionStartTime
    .Specifier = pContract.Specifier
    .TickSize = pContract.TickSize
    .Timezone = pContract.Timezone
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Sub

Friend Sub Initialise( _
                ByVal Specifier As ContractSpecifier)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mContract = New Contract
mContract.Specifier = Specifier
mContract.Timezone = GetTimeZone("")
mContract.DaysBeforeExpiryToSwitch = 1

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
                
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mContract = New Contract
With config
    mContract.Specifier = gCreateContractSpecifier(.GetSetting(ConfigSettingContractSpecLocalSYmbol, ""), _
                                            .GetSetting(ConfigSettingContractSpecSymbol, ""), _
                                            .GetSetting(ConfigSettingContractSpecExchange, ""), _
                                            gSecTypeFromString(.GetSetting(ConfigSettingContractSpecSecType, "")), _
                                            .GetSetting(ConfigSettingContractSpecCurrency, ""), _
                                            .GetSetting(ConfigSettingContractSpecExpiry, ""), _
                                            CDbl(.GetSetting(ConfigSettingContractSpecStrikePrice, "0.0")), _
                                            gOptionRightFromString(.GetSetting(ConfigSettingContractSpecRight, "")))

    mContract.DaysBeforeExpiryToSwitch = .GetSetting(ConfigSettingDaysBeforeExpiryToSwitch, "1")
    mContract.Description = .GetSetting(ConfigSettingDescription, "")
    mContract.ExpiryDate = CDate(.GetSetting(ConfigSettingExpiryDate, ""))
    mContract.Multiplier = .GetSetting(ConfigSettingMultiplier, "1")
    mContract.SessionEndTime = .GetSetting(ConfigSettingSessionEndTime, "00:00:00")
    mContract.SessionStartTime = .GetSetting(ConfigSettingSessionStartTime, "00:00:00")
    mContract.TickSize = .GetSetting(ConfigSettingSessionTickSize, "0.01")
    mContract.Timezone = GetTimeZone(.GetSetting(ConfigSettingSessionTimezone, ""))
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


