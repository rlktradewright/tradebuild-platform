VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTickfileListGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mContracts As IContracts

Private Const TickfileFormatIdentifier As String = "wazzox"

'@================================================================================
' Setup Functions
'@================================================================================

Public Sub FixtureSetup()
End Sub

Public Sub FixtureTeardown()
End Sub

Public Sub Setup()
Dim lContractsBuilder As New ContractsBuilder

lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM03")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZZ13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU14")

Set mContracts = lContractsBuilder.Contracts
End Sub

Public Sub TearDown()
LogMessage "Test Completed -----------------------------------------------"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Test Methods
'@================================================================================

'@================================================================================
' Tests for non-expiring contracts, not whole sessions
'@================================================================================

Public Sub Generate010NonExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
lContractsBuilder.Add gCreateContractFromLocalSymbol("IBM")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

On Error Resume Next
Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/12 16:07", _
                                        pCompleteSessionsOnly:=False)

Assert.AreEqual ErrorCodes.ErrIllegalArgumentException, Err.Number
Assert.AreEqual "pContracts contains conflicting contracts", Err.Description

End Sub

Public Sub Generate020NonExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/12 16:07", _
                                        pCompleteSessionsOnly:=False)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:07"), "Incorrect ToDate"

End Sub

Public Sub Generate030NonExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 16:35", _
                                        "2013/08/12 21:07", _
                                        pCompleteSessionsOnly:=False, _
                                        pUseExchangeTimezone:=False)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:07"), "Incorrect ToDate"

End Sub

Public Sub Generate040NonExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        CDate(0), _
                                        pCompleteSessionsOnly:=False)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:15"), "Incorrect ToDate"

End Sub

Public Sub Generate050NonExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/13 15:22", _
                                        pCompleteSessionsOnly:=False)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/13 15:22"), "Incorrect ToDate"

End Sub


'@================================================================================
' Tests for non-expiring contracts, whole sessions
'@================================================================================

Public Sub Generate110NonExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/12 16:07", _
                                        pCompleteSessionsOnly:=True, _
                                        pCustomSessionStartTime:="10:00", _
                                        pCustomSessionEndTime:="15:00")

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 10:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 15:00"), "Incorrect ToDate"

End Sub

Public Sub Generate120NonExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        CDate(0), _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 09:30"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:15"), "Incorrect ToDate"

End Sub

Public Sub Generate130NonExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/11 11:35", _
                                        "2013/08/11 15:22", _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 0, "No TickfileSpecifiers should be created"

End Sub

Public Sub Generate140NonExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/11 11:35", _
                                        "2013/08/12 15:22", _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 09:30"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:15"), "Incorrect ToDate"

End Sub

Public Sub Generate150NonExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/13 15:22", _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 2, "2 TickfileSpecifiers should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 09:30"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:15"), "Incorrect ToDate"

en.MoveNext
Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/13 09:30"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/13 16:15"), "Incorrect ToDate"
End Sub

Public Sub Generate160NonExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("MSFT")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/09 11:35", _
                                        "2013/08/19 15:22", _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 7, "7 TickfileSpecifiers should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/09 09:30"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/09 16:15"), "Incorrect ToDate"

Do While en.MoveNext
    Set tfs = en.Current
Loop

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/19 09:30"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/19 16:15"), "Incorrect ToDate"
End Sub


'@================================================================================
' Tests for expiring contracts, not whole sessions
'@================================================================================

Public Sub Generate200ExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/03/05 11:35", _
                                        "2013/03/05 16:07", _
                                        pCompleteSessionsOnly:=False)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(1).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/03/05 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/03/05 16:07"), "Incorrect ToDate"

End Sub

Public Sub Generate210ExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

On Error Resume Next
Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/03/22 11:35", _
                                        "2013/03/22 16:07", _
                                        pCompleteSessionsOnly:=False)

Assert.AreEqual ErrorCodes.ErrIllegalStateException, Err.Number
Assert.AreEqual "No contract includes specifed from date", Err.Description

End Sub

Public Sub Generate220ExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZZ13")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/12 16:07", _
                                        pCompleteSessionsOnly:=False)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(3).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 16:07"), "Incorrect ToDate"

End Sub

Public Sub Generate230ExpiringNonSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZZ2")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU3")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/03/12 11:35", _
                                        "2013/09/16 16:07", _
                                        pCompleteSessionsOnly:=False)

Assert.IsTrue tfSpecs.Count = 3, "3 TickfileSpecifiers should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(2).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/03/12 11:35"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/03/14 08:00"), "Incorrect ToDate"

en.MoveNext
Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(3).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/03/14 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/06/20 08:00"), "Incorrect ToDate"

en.MoveNext
Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(4).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsFalse tfs.EntireSession, "TickfleSpecifier.EntireSession should be False"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/06/20 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/09/16 16:07"), "Incorrect ToDate"

End Sub


'@================================================================================
' Tests for expiring contracts, whole sessions
'@================================================================================

Public Sub Generate310ExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZZ13")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/08/12 11:35", _
                                        "2013/08/12 16:07", _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 1, "Only 1 TickfileSpecifier should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(3).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/08/12 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/08/12 17:30"), "Incorrect ToDate"

End Sub

Public Sub Generate320ExpiringWholeSession()
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZH13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZM13")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZU3")
lContractsBuilder.Add gCreateContractFromLocalSymbol("ZZ13")
Dim lContracts As IContracts: Set lContracts = lContractsBuilder.Contracts

Dim tfSpecs As TickFileSpecifiers
Set tfSpecs = GenerateTickfileSpecifiers(lContracts, _
                                        "wazzox", _
                                        "2013/09/17 11:35", _
                                        "2013/09/25 16:07", _
                                        pCompleteSessionsOnly:=True)

Assert.IsTrue tfSpecs.Count = 7, "7 TickfileSpecifiers should be created"

Dim en As Enumerator: Set en = tfSpecs.Enumerator
en.MoveNext

Dim tfs As TickfileSpecifier: Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(3).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/09/17 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/09/17 17:30"), "Incorrect ToDate"

en.MoveNext
Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(3).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/09/18 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/09/18 17:30"), "Incorrect ToDate"

en.MoveNext
Set tfs = en.Current

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(4).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/09/19 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/09/19 17:30"), "Incorrect ToDate"

Do While en.MoveNext
    Set tfs = en.Current
Loop

Assert.IsTrue ContractSpecsEqual(lContracts.ItemAtIndex(4).Specifier, tfs.Contract.Specifier), "TickfleSpecifier contains the wrong contract"
Assert.IsTrue tfs.EntireSession, "TickfleSpecifier.EntireSession should be True"
Assert.IsTrue tfs.TickfileFormatID = TickfileFormatIdentifier, "Incorrect tickfile format identifier"
Assert.IsTrue tfs.FromDate = CDate("2013/09/25 08:00"), "Incorrect FromDate"
Assert.IsTrue tfs.ToDate = CDate("2013/09/25 17:30"), "Incorrect ToDate"

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




