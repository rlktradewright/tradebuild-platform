VERSION 5.00
Begin VB.UserControl ContractSearch 
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4800
   Begin VB.CommandButton ClearButton 
      Caption         =   "Clear"
      Height          =   330
      Left            =   1680
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton ActionButton 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3120
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   3863
      _ExtentY        =   6509
   End
   Begin TradingUI27.ContractSelector ContractSelector1 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _ExtentX        =   0
      _ExtentY        =   0
      RowBackColorOdd =   16316664
      RowBackColorEven=   15658734
   End
   Begin VB.Label MessageLabel 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E7D395&
      X1              =   0
      X2              =   4800
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "ContractSearch"
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

'@================================================================================
' Events
'@================================================================================

Event Action()

Event Error(ev As ErrorEventData)

Event NoContracts()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractSearch"

'@================================================================================
' Member variables
'@================================================================================

Private mContractStorePrimary                       As IContractStore
Private mContractStoreSecondary                     As IContractStore

Private mContracts                                  As IContracts

Private mLoadingContracts                           As Boolean

Private mAllowMultipleSelection                     As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_EnterFocus()
If ContractSpecBuilder1.Visible Then
    ContractSpecBuilder1.SetFocus
Else
    ContractSelector1.SetFocus
End If
End Sub

Private Sub UserControl_Initialize()
Set mFutureWaiter = New FutureWaiter
ContractSpecBuilder1.Visible = True
ContractSelector1.Visible = False
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

AllowMultipleSelection = True
ActionButtonCaption = "Start"
IncludeHistoricalContracts = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

ActionButtonCaption = PropBag.ReadProperty("ActionButtonCaption", "Start")
If Err.Number <> 0 Then
    ActionButtonCaption = "Start"
    Err.Clear
End If

AllowMultipleSelection = CBool(PropBag.ReadProperty("AllowMultipleSelection", "True"))
If Err.Number <> 0 Then
    AllowMultipleSelection = True
    Err.Clear
End If

IncludeHistoricalContracts = CBool(PropBag.ReadProperty("IncludeHistoricalContracts", "False"))
If Err.Number <> 0 Then
    IncludeHistoricalContracts = False
    Err.Clear
End If

End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"

On Error GoTo Err

If UserControl.Height <= ContractSelector1.Height Then UserControl.Height = ContractSelector1.Height

ActionButton.Left = UserControl.Width - ActionButton.Width
ClearButton.Left = ActionButton.Left - ClearButton.Width - 120

ActionButton.Top = UserControl.Height - ActionButton.Height
ClearButton.Top = ActionButton.Top
MessageLabel.Top = ActionButton.Top

Line1.Y1 = ActionButton.Top - 120
Line1.Y2 = Line1.Y1
Line1.X2 = UserControl.Width

ContractSpecBuilder1.Width = UserControl.Width
ContractSelector1.Width = UserControl.Width

ContractSelector1.Height = Line1.Y1 - 120

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty "ActionButtonCaption", ActionButtonCaption, "Start"
PropBag.WriteProperty "AllowMultipleSelection", AllowMultipleSelection, "True"
PropBag.WriteProperty "IncludeHistoricalContracts", IncludeHistoricalContracts, "False"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ActionButton_Click()
Const ProcName As String = "ActionButton_Click"
On Error GoTo Err

If ContractSelector1.Visible Then
    RaiseEvent Action
Else
    mFutureWaiter.Add FetchContracts(ContractSpecBuilder1.ContractSpecifier, mContractStorePrimary, mContractStoreSecondary)
    mLoadingContracts = True
    ActionButton.Enabled = False
    UserControl.MousePointer = MousePointerConstants.vbHourglass
    MessageLabel.caption = "Searching..."
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ClearButton_Click()
ContractSpecBuilder1.Visible = True
ContractSelector1.Visible = False
ContractSpecBuilder1.SetFocus
ClearButton.Visible = False
MessageLabel.caption = ""
End Sub

Private Sub ContractSpecBuilder1_NotReady()
If Not mLoadingContracts Then ActionButton.Enabled = False
End Sub

Private Sub ContractSpecBuilder1_Ready()
If Not mLoadingContracts Then
    ActionButton.Enabled = True
    ActionButton.Default = True
End If
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then Exit Sub

MessageLabel.caption = ""
UserControl.MousePointer = MousePointerConstants.vbDefault
ActionButton.Enabled = True
mLoadingContracts = False

If ContractSpecBuilder1.IsReady Then
    ActionButton.Enabled = True
    ActionButton.Default = True
Else
    ActionButton.Enabled = False
    ActionButton.Default = False
End If

If ev.Future.IsFaulted Then
    Dim lEv As ErrorEventData
    Set lEv.Source = Me
    lEv.ErrorCode = ev.Future.ErrorNumber
    lEv.ErrorMessage = ev.Future.ErrorMessage
    lEv.ErrorSource = ev.Future.ErrorSource
    RaiseEvent Error(lEv)
Else
    Set mContracts = ev.Future.value
    handleContractsLoaded
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let ActionButtonCaption( _
                ByVal value As String)
ActionButton.caption = value
PropertyChanged ActionButtonCaption
End Property

Public Property Get ActionButtonCaption() As String
ActionButtonCaption = ActionButton.caption
End Property

Public Property Get AllowMultipleSelection() As Boolean
AllowMultipleSelection = mAllowMultipleSelection
End Property

Public Property Let AllowMultipleSelection(ByVal value As Boolean)
mAllowMultipleSelection = value
PropertyChanged AllowMultipleSelection
End Property

Public Property Let IncludeHistoricalContracts( _
                ByVal value As Boolean)
Const ProcName As String = "IncludeHistoricalContracts"
On Error GoTo Err

ContractSelector1.IncludeHistoricalContracts = value

PropertyChanged IncludeHistoricalContracts
Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeHistoricalContracts() As Boolean
Const ProcName As String = "IncludeHistoricalContracts"
On Error GoTo Err

IncludeHistoricalContracts = ContractSelector1.IncludeHistoricalContracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Public Property Let InitialContracts(ByVal pContracts As IContracts)
'Const ProcName As String = "InitialContracts"
'On Error GoTo Err
'
'AssertArgument Not pContracts Is Nothing, "pContracts must not be Nothing"
'
'Set mContracts = pContracts
'ContractSpecBuilder1.ContractSpecifier = mContracts.ContractSpecifier
'
'If mContracts.Count > 0 Then setupContractSelector mContracts, mAllowMultipleSelection
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

Public Property Get SelectedContracts() As IContracts
Const ProcName As String = "SelectedContracts"
On Error GoTo Err

If mContracts.Count = 1 Then
    Set SelectedContracts = mContracts
Else
    Set SelectedContracts = ContractSelector1.SelectedContracts
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pContractStorePrimary As IContractStore, _
                ByVal pContractStoreSecondary As IContractStore)
Set mContractStorePrimary = pContractStorePrimary
Set mContractStoreSecondary = pContractStoreSecondary
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleContractsLoaded()
Const ProcName As String = "handleContractsLoaded"
On Error GoTo Err

If mContracts.Count = 0 Then
    MessageLabel.caption = "No contracts"
    RaiseEvent NoContracts
ElseIf mContracts.Count = 1 Then
    If IncludeHistoricalContracts Or Not IsContractExpired(mContracts.ItemAtIndex(1)) Then
        RaiseEvent Action
    Else
        MessageLabel.caption = "Contract expired"
        RaiseEvent NoContracts
    End If
Else
    setupContractSelector mContracts, mAllowMultipleSelection
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupContractSelector( _
                ByVal pContracts As IContracts, _
                ByVal pAllowMultipleSelection As Boolean)
Const ProcName As String = "setupContractSelector"
On Error GoTo Err

ContractSelector1.Initialise pContracts, pAllowMultipleSelection
MessageLabel.caption = ContractSelector1.Count & IIf(ContractSelector1.Count = 1, " contract", " contracts")
ContractSelector1.Visible = True
ContractSpecBuilder1.Visible = False
ClearButton.Visible = True

On Error Resume Next    ' because SetFocus gives an error if called during a Form_Load event
ContractSelector1.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
