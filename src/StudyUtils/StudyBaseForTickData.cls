VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyBaseForTickDataInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IGenericTickListener
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

Private Const ModuleName                            As String = "StudyBaseForTickDataInput"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyManager                               As StudyManager
Private mStudyInputHandler                          As StudyInputHandler

Private mInputHandleBarNumber                       As Long
Private mInputHandleBid                             As Long
Private mInputHandleAsk                             As Long
Private mInputHandleOpenInterest                    As Long
Private mInputHandleTickVolume                      As Long
Private mInputHandleTrade                           As Long
Private mInputHandleVolume                          As Long

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mIsInitialised                              As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
' nothing to do
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

NotifyTick ev.Tick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStudyBase Interface Members
'@================================================================================

Private Property Get IStudyBase_BaseStudy() As IStudy
Set IStudyBase_BaseStudy = BaseStudy
End Property

Private Property Get IStudyBase_StudyManager() As StudyManager
Set IStudyBase_StudyManager = StudyManager
End Property

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContract As IContract
Set lContract = ev.Future.Value

mStudyInputHandler.Name = lContract.Specifier.ToString

mStudyManager.SetInputTicksize mInputHandleAsk, lContract.TickSize
mStudyManager.SetInputTicksize mInputHandleBid, lContract.TickSize
mStudyManager.SetInputTicksize mInputHandleTrade, lContract.TickSize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BaseStudy() As IStudy
Set BaseStudy = mStudyInputHandler
End Property

Public Property Get StudyManager() As StudyManager
Set StudyManager = mStudyManager
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pStudyManager As StudyManager, _
                ByVal pContractFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = pContractFuture.Value
    init pStudyManager, lContract.Specifier.ToString, lContract.TickSize
Else
    waitForFuture pContractFuture
    init pStudyManager, CStr(GetObjectKey(pContractFuture)), 0#
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub InitialiseWithContract( _
                ByVal pStudyManager As StudyManager, _
                ByVal pContract As IContract)
Const ProcName As String = "InitialiseWithContract"
On Error GoTo Err

init pStudyManager, pContract.Specifier.ToString, pContract.TickSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyBarNumber(pBarNumber As Long, ByVal pTimestamp As Date)
Const ProcName As String = "NotifyBarNumber"
On Error GoTo Err

mStudyManager.NotifyInput mInputHandleBarNumber, pBarNumber, pTimestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub NotifyTick(pTick As GenericTick)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

Assert mIsInitialised, "Not initialised"

Select Case pTick.TickType
Case TickTypeBid
    mStudyManager.NotifyInput mInputHandleBid, pTick.Price, pTick.Timestamp
Case TickTypeAsk
    mStudyManager.NotifyInput mInputHandleAsk, pTick.Price, pTick.Timestamp
Case TickTypeTrade
    mStudyManager.NotifyInput mInputHandleTrade, pTick.Price, pTick.Timestamp
Case TickTypeVolume
    mStudyManager.NotifyInput mInputHandleVolume, pTick.Size, pTick.Timestamp
Case TickTypeOpenInterest
    mStudyManager.NotifyInput mInputHandleOpenInterest, pTick.Size, pTick.Timestamp
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub init( _
                ByVal pStudyManager As StudyManager, _
                ByVal pName As String, _
                ByVal pTickSize As Double)
Const ProcName As String = "init"
On Error GoTo Err

Assert Not mIsInitialised, "Already initialised"
mIsInitialised = True

Set mStudyManager = pStudyManager

Set mStudyInputHandler = mStudyManager.CreateStudyInputHandler(pName)
mInputHandleBid = mStudyManager.AddInput(mStudyInputHandler, _
                        BidInputName, _
                        "Bid prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        pTickSize)
mInputHandleAsk = mStudyManager.AddInput(mStudyInputHandler, _
                        AskInputName, _
                        "Ask prices", _
                        StudyInputTypes.InputTypeReal, _
                        False, _
                        pTickSize)
mInputHandleOpenInterest = mStudyManager.AddInput(mStudyInputHandler, _
                        OpenInterestInputName, _
                        "OpenInterest", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
mInputHandleTickVolume = mStudyManager.AddInput(mStudyInputHandler, _
                        TickVolumeInputName, _
                        "Tick Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
mInputHandleTrade = mStudyManager.AddInput(mStudyInputHandler, _
                        TradeInputName, _
                        "Trade prices", _
                        StudyInputTypes.InputTypeReal, _
                        True, _
                        pTickSize)
mInputHandleVolume = mStudyManager.AddInput(mStudyInputHandler, _
                        VolumeInputName, _
                        "Volume", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)
mInputHandleBarNumber = mStudyManager.AddInput(mStudyInputHandler, _
                        BarNumberInputName, _
                        "Bar number", _
                        StudyInputTypes.InputTypeInteger, _
                        False, _
                        1)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



Private Sub waitForFuture(ByVal pFuture As IFuture)
Const ProcName As String = "waitForFuture"
On Error GoTo Err

If mFutureWaiter Is Nothing Then Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add pFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


