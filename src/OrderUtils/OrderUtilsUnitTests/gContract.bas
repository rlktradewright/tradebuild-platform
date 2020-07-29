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

Public Function gCreateContractFromLocalSymbol(ByVal pLocalSymbol As String) As IContract
Dim lContract As IContract

Select Case pLocalSymbol
Case "ESM3"
    Set lContract = createESContract("ESM3", "20130621")
Case "ZM03"
    Set lContract = createZContract("ZM03", "20030620")
Case "ZZ2"
    Set lContract = createZContract("ZZ2", "20121221")
Case "ZH3"
    Set lContract = createZContract("ZH3", "20130315")
Case "ZM3"
    Set lContract = createZContract("ZM3", "20130621")
Case "ZU3"
    Set lContract = createZContract("ZU3", "20130920")
Case "ZZ3"
    Set lContract = createZContract("ZZ3", "20131220")
Case "ZU4"
    Set lContract = createZContract("ZU4", "20140918")
Case "IBM"
    Set lContract = createStockContract("IBM")
Case "MSFT"
    Set lContract = createStockContract("MSFT")
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Localname not known"
End Select
    
Set gCreateContractFromLocalSymbol = lContract
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createESContract(ByVal localSymbol As String, ByVal expiry As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(localSymbol, "ES", "GLOBEX", SecTypeFuture, "USD", expiry)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("16:15")
lContractBuilder.SessionStartTime = CDate("16:30")
lContractBuilder.TickSize = 0.25
lContractBuilder.ExpiryDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 4, 2) & "/" & Right$(expiry, 2))
lContractBuilder.TimezoneName = "Central Standard Time"
lContractBuilder.DaysBeforeExpiryToSwitch = 1
Set createESContract = lContractBuilder.Contract
End Function

Private Function createStockContract(ByVal localSymbol As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(localSymbol, localSymbol, "SMART", SecTypeStock, "USD", "")

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("16:15")
lContractBuilder.SessionStartTime = CDate("09:30")
lContractBuilder.TickSize = 0.01
lContractBuilder.TimezoneName = "Eastern Standard Time"
Set createStockContract = lContractBuilder.Contract
End Function

Private Function createZContract(ByVal localSymbol As String, ByVal expiry As String) As IContract
Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(localSymbol, "Z", "ICEEU", SecTypeFuture, "GBP", expiry)

Dim lContractBuilder As ContractBuilder
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractBuilder.SessionEndTime = CDate("17:30")
lContractBuilder.SessionStartTime = CDate("08:00")
lContractBuilder.TickSize = 0.5
lContractBuilder.TimezoneName = "GMT Standard Time"
lContractBuilder.ExpiryDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
lContractBuilder.DaysBeforeExpiryToSwitch = 1
Set createZContract = lContractBuilder.Contract
End Function




