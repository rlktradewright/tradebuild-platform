Attribute VB_Name = "gContract"
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

Private Const ModuleName                            As String = "gContract"

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateOptionContract( _
                ByVal pSymbol As String, _
                ByVal pLocalSymbol As String, _
                ByVal pExchange As String, _
                ByVal pExpiry As String, _
                ByVal pRight As OptionRights, _
                ByVal pStrike As Double) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(pLocalSymbol, pSymbol, pExchange, SecTypeOption, "USD", pExpiry, , pStrike, pRight)


Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("15:15")
lContractBuilder.SessionStartTime = CDate("08:30")
lContractBuilder.TickSize = 0.01
lContractBuilder.TimezoneName = "Central Standard Time"
Set gCreateOptionContract = lContractBuilder.Contract
End Function

Public Function gCreateStockContract(ByVal pLocalSymbol As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(pLocalSymbol, pLocalSymbol, "SMART", SecTypeStock, "USD", "")

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("16:15")
lContractBuilder.SessionStartTime = CDate("09:30")
lContractBuilder.TickSize = 0.01
lContractBuilder.TimezoneName = "Eastern Standard Time"
Set gCreateStockContract = lContractBuilder.Contract
End Function

Public Function gFetchOptionExpiries( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pStrike As Double = 0#, _
                Optional ByVal pCookie As Variant) As IFuture
Dim lExpiriesBuilder As New ExpiriesBuilder
lExpiriesBuilder.Add "20200731"
lExpiriesBuilder.Add "20200807"
lExpiriesBuilder.Add "20200814"
lExpiriesBuilder.Add "20200821"
Set gFetchOptionExpiries = CreateFuture(lExpiriesBuilder.Expiries, pCookie)
End Function

Public Function gFetchOptionStrikes( _
                ByVal pUnderlyingContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                Optional ByVal pExpiry As String, _
                Optional ByVal pCookie As Variant) As IFuture
Dim lStrikesBuilder As New StrikesBuilder
lStrikesBuilder.Add 187.5
lStrikesBuilder.Add 190
lStrikesBuilder.Add 192.5
lStrikesBuilder.Add 195
lStrikesBuilder.Add 197.5
lStrikesBuilder.Add 200
lStrikesBuilder.Add 202.5
lStrikesBuilder.Add 205
lStrikesBuilder.Add 207.5
lStrikesBuilder.Add 210#
lStrikesBuilder.Add 212.5
lStrikesBuilder.Add 215#
Set gFetchOptionStrikes = CreateFuture(lStrikesBuilder.Strikes, pCookie)
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createESContract(ByVal pLocalSymbol As String, ByVal pExpiry As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(pLocalSymbol, "ES", "GLOBEX", SecTypeFuture, "USD", pExpiry)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("16:15")
lContractBuilder.SessionStartTime = CDate("16:30")
lContractBuilder.TickSize = 0.25
lContractBuilder.ExpiryDate = CDate(Left$(pExpiry, 4) & "/" & Mid$(pExpiry, 5, 2) & "/" & Right$(pExpiry, 2))
lContractBuilder.TimezoneName = "Central Standard Time"
lContractBuilder.DaysBeforeExpiryToSwitch = 1
Set createESContract = lContractBuilder.Contract
End Function

Private Function createZContract(ByVal pLocalSymbol As String, ByVal pExpiry As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(pLocalSymbol, "Z", "ICEEU", SecTypeFuture, "GBP", pExpiry)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("17:30")
lContractBuilder.SessionStartTime = CDate("08:00")
lContractBuilder.TickSize = 0.5
lContractBuilder.TimezoneName = "GMT Standard Time"
lContractBuilder.ExpiryDate = CDate(Left$(pExpiry, 4) & "/" & Mid$(pExpiry, 5, 2) & "/" & Right$(pExpiry, 2))
lContractBuilder.DaysBeforeExpiryToSwitch = 1
Set createZContract = lContractBuilder.Contract
End Function




