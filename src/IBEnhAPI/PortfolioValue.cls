VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PortfolioValue"
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

Implements IPortfolioValue

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

Private Const ModuleName                            As String = "PortfolioValue"

'@================================================================================
' Member variables
'@================================================================================

Private mAccountName                                As String
Private mAverageCost                                As Double
Private mContractSpecifier                          As IContractSpecifier
Private mMarketPrice                                As Double
Private mMarketValue                                As Double
Private mPosition                                   As Double
Private mRealizedPNL                                As Double
Private mUnrealizedPNL                              As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IPortfolioValue Interface Members
'@================================================================================

Private Property Get IPortfolioValue_AccountName() As String
IPortfolioValue_AccountName = mAccountName
End Property

Private Property Get IPortfolioValue_AverageCost() As Double
IPortfolioValue_AverageCost = mAverageCost
End Property

Private Property Get IPortfolioValue_ContractSpecifier() As IContractSpecifier
Set IPortfolioValue_ContractSpecifier = mContractSpecifier
End Property

Private Property Get IPortfolioValue_MarketPrice() As Double
IPortfolioValue_MarketPrice = mMarketPrice
End Property

Private Property Get IPortfolioValue_MarketValue() As Double
IPortfolioValue_MarketValue = mMarketValue
End Property

Private Property Get IPortfolioValue_Position() As Double
IPortfolioValue_Position = mPosition
End Property

Private Property Get IPortfolioValue_RealizedPNL() As Double
IPortfolioValue_RealizedPNL = mRealizedPNL
End Property

Private Property Get IPortfolioValue_UnrealizedPNL() As Double
IPortfolioValue_UnrealizedPNL = mUnrealizedPNL
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

Friend Sub Initialise( _
                ByVal pAccountName As String, _
                ByVal pAverageCost As Double, _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pMarketPrice As Double, _
                ByVal pMarketValue As Double, _
                ByVal pPosition As Double, _
                ByVal pRealizedPNL As Double, _
                ByVal pUnrealizedPNL As Double)
mAccountName = pAccountName
mAverageCost = pAverageCost
Set mContractSpecifier = pContractSpecifier
mMarketPrice = pMarketPrice
mMarketValue = pMarketValue
mPosition = pPosition
mRealizedPNL = pRealizedPNL
mUnrealizedPNL = pUnrealizedPNL
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




