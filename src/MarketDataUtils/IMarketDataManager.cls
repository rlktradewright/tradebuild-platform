VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMarketDataManager"
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

Private Const ModuleName                            As String = "IMarketDataManager"

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

Public Property Get PrimaryContractStore() As IContractStore
End Property

'Public Property Get DataSources() As EnumerableCollection
'End Property

Public Property Get SecondaryContractStore() As IContractStore
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateMarketDataSource( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pContractSpec As IContractSpecifier, _
                Optional ByVal pKey As String, _
                Optional ByVal pStateChangeListener As StateChangeListener) As IMarketDataSource
End Function

Public Sub Finish()
End Sub

Public Function GetMarketDataSource(ByVal pKey As String) As IMarketDataSource
End Function

'@================================================================================
' Helper Functions
'@================================================================================




