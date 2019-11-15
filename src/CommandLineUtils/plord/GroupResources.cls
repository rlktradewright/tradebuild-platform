VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GroupResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "GroupResources"

'@================================================================================
' Member variables
'@================================================================================

Private mGroupName                                  As String

Private mContractProcessors                         As New EnumerableCollection

Private mCurrentContractProcessor                   As ContractProcessor

Private mOrderPlacers                               As New EnumerableCollection

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

Friend Property Get ContractProcessors() As EnumerableCollection
Set ContractProcessors = mContractProcessors
End Property

Friend Property Let CurrentContractProcessor(ByVal Value As ContractProcessor)
Set mCurrentContractProcessor = Value
End Property

Friend Property Get CurrentContractProcessor() As ContractProcessor
Set CurrentContractProcessor = mCurrentContractProcessor
End Property

Friend Property Let GroupName(ByVal Value As String)
mGroupName = Value
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
End Property

Friend Property Get OrderPlacers() As EnumerableCollection
Set OrderPlacers = mOrderPlacers
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




