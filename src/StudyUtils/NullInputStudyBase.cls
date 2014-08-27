VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NullInputStudyBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IStudyBase

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

Private Const ModuleName                            As String = "NullInputStudyBase"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyManager                               As StudyManager
Private mStudyInputHandler                          As StudyInputHandler

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStudyBase Interface Members
'@================================================================================

Private Property Get IStudyBase_BaseStudy() As IStudy
Const ProcName As String = "IStudyBase_BaseStudy"
On Error GoTo Err

Set IStudyBase_BaseStudy = mStudyInputHandler

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IStudyBase_StudyManager() As StudyManager
Const ProcName As String = "IStudyBase_StudyManager"
On Error GoTo Err

Set IStudyBase_StudyManager = mStudyManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContract As IContract
Set lContract = ev.Future.value

mStudyInputHandler.Name = lContract.Specifier.ToString

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Sub Initialise(ByVal pStudyManager As StudyManager, ByVal pContractFuture As IFuture)
Set mStudyManager = pStudyManager

Dim lName As String

Set mStudyManager = pStudyManager

If pContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = pContractFuture.value
    lName = lContract.Specifier.ToString
Else
    lName = CStr(GetObjectKey(pContractFuture))
    Set mFutureWaiter = New FutureWaiter
    mFutureWaiter.Add pContractFuture
End If

Set mStudyInputHandler = mStudyManager.CreateStudyInputHandler(lName)
End Sub

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




