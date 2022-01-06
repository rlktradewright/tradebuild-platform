VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#33.0#0"; "TWControls40.ocx"
Begin VB.UserControl ContractSearch 
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4800
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3150
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6191
      ForeColor       =   -2147483640
   End
   Begin TWControls40.TWButton ClearButton 
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "Clear"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton ActionButton 
      Height          =   330
      Left            =   3240
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TradingUI27.ContractSelector ContractSelector1 
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5741
   End
   Begin VB.Label MessageLabel 
      Height          =   375
      Left            =   0
      TabIndex        =   1
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

Implements IContractFetchListener
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Action()

Event Cancelled()

Event Cleared()

Event ContractsLoaded(ByVal pContracts As IContracts)

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

Private Const PropNameActionButtonCaption           As String = "ActionButtonCaption"
Private Const PropNameAllowMultipleSelection        As String = "AllowMultipleSelection"
Private Const PropNameBackcolor                     As String = "BackColor"
Private Const PropNameForecolor                     As String = "ForeColor"
Private Const PropNameIncludeHistoricalContracts    As String = "IncludeHistoricalContracts"
Private Const PropNameRowBackColorEven              As String = "RowBackColorEven"
Private Const PropNameRowBackColorOdd               As String = "RowBackColorOdd"
Private Const PropNameTextBackColor                 As String = "TextBackColor"
Private Const PropNameTextForeColor                 As String = "TextForeColor"

Private Const CancelCaption                         As String = "Cancel"
Private Const ClearCaption                          As String = "Clear"

'@================================================================================
' Member variables
'@================================================================================

Private mContractStorePrimary                       As IContractStore
Private mContractStoreSecondary                     As IContractStore

Private mContractsBuilder                           As IContractsBuilder
Private mSortKeys()                                 As ContractSortKeyIds

Private mLoadingContracts                           As Boolean

Private mAllowMultipleSelection                     As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mCookie                                     As Variant

Private mTheme                                      As ITheme

Private WithEvents mContractSelectorInitialisationTC    As TaskController
Attribute mContractSelectorInitialisationTC.VB_VarHelpID = -1

Private mCancelInitiatedByUser                      As Boolean

Private mSingleContracts                            As IContracts

Private mHistoricalContractsFound                   As Boolean

Private mContractSpec                               As IContractSpecifier

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

ReDim mSortKeys(8) As ContractSortKeyIds
mSortKeys(0) = ContractSortKeySecType
mSortKeys(1) = ContractSortKeySymbol
mSortKeys(2) = ContractSortKeyExchange
mSortKeys(3) = ContractSortKeyCurrency
mSortKeys(4) = ContractSortKeyExpiry
mSortKeys(5) = ContractSortKeyMultiplier
mSortKeys(6) = ContractSortKeyStrike
mSortKeys(7) = ContractSortKeyRight
mSortKeys(8) = ContractSortKeyLocalSymbol
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

AllowMultipleSelection = True
ActionButtonCaption = "Start"
BackColor = vbButtonFace
ForeColor = vbButtonText
IncludeHistoricalContracts = False
RowBackColorEven = CRowBackColorEven
RowBackColorOdd = CRowBackColorOdd
TextBackColor = vbWindowBackground
TextForeColor = vbWindowText
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

ActionButtonCaption = PropBag.ReadProperty(PropNameActionButtonCaption, "Start")
AllowMultipleSelection = CBool(PropBag.ReadProperty(PropNameAllowMultipleSelection, "True"))
BackColor = PropBag.ReadProperty(PropNameBackcolor, vbButtonFace)
ForeColor = PropBag.ReadProperty(PropNameForecolor, vbButtonText)
IncludeHistoricalContracts = CBool(PropBag.ReadProperty(PropNameIncludeHistoricalContracts, "False"))
RowBackColorEven = PropBag.ReadProperty(PropNameRowBackColorEven, CRowBackColorEven)
RowBackColorOdd = PropBag.ReadProperty(PropNameRowBackColorOdd, CRowBackColorOdd)
TextBackColor = PropBag.ReadProperty(PropNameTextBackColor, vbWindowBackground)
TextForeColor = PropBag.ReadProperty(PropNameTextForeColor, vbWindowText)
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
PropBag.WriteProperty PropNameActionButtonCaption, ActionButtonCaption, "Start"
PropBag.WriteProperty PropNameAllowMultipleSelection, AllowMultipleSelection, "True"
PropBag.WriteProperty PropNameBackcolor, BackColor, vbButtonFace
PropBag.WriteProperty PropNameForecolor, ForeColor, vbButtonText
PropBag.WriteProperty PropNameIncludeHistoricalContracts, IncludeHistoricalContracts, "False"
PropBag.WriteProperty PropNameRowBackColorEven, RowBackColorEven, CRowBackColorEven
PropBag.WriteProperty PropNameRowBackColorOdd, RowBackColorOdd, CRowBackColorOdd
PropBag.WriteProperty PropNameTextBackColor, TextBackColor, vbWindowBackground
PropBag.WriteProperty PropNameTextForeColor, TextForeColor, vbWindowText
End Sub

'@================================================================================
' IContractFetchListener Interface Members
'@================================================================================

Private Sub IContractFetchListener_FetchCancelled(ByVal pCookie As Variant)

End Sub

Private Sub IContractFetchListener_FetchCompleted(ByVal pCookie As Variant)

End Sub

Private Sub IContractFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)

End Sub

Private Sub IContractFetchListener_NotifyContract(ByVal pCookie As Variant, ByVal pContract As IContract)
Const ProcName As String = "IContractFetchListener_NotifyContract"
On Error GoTo Err

If Not IncludeHistoricalContracts And IsContractExpired(pContract) Then
    mHistoricalContractsFound = True
    Exit Sub
End If

If mContractsBuilder Is Nothing Then
    Set mContractsBuilder = New ContractsBuilder
    mContractsBuilder.Contracts.SortKeys = mSortKeys
    ActionButton.Enabled = False
    ClearButton.Visible = True
    ClearButton.Caption = CancelCaption
    ClearButton.Enabled = True
    ContractSelector1.Clear
    ContractSelector1.Visible = True
    ContractSpecBuilder1.Visible = False
End If

mContractsBuilder.Add pContract

If mContractsBuilder.Contracts.Count Mod 100 = 0 Then
    MessageLabel.Caption = getContractsCountMessage(mContractsBuilder.Contracts.Count)
    MessageLabel.Refresh
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ActionButton_Click()
Const ProcName As String = "ActionButton_Click"
On Error GoTo Err

If ContractSelector1.Visible Then
    RaiseEvent Action
    mCookie = Empty
Else
    Set mSingleContracts = Nothing
    mHistoricalContractsFound = False
    Set mContractSpec = ContractSpecBuilder1.ContractSpecifier
    mFutureWaiter.Add FetchContracts(mContractSpec, mContractStorePrimary, mContractStoreSecondary, Me)
    mLoadingContracts = True
    UserControl.MousePointer = MousePointerConstants.vbHourglass
    MessageLabel.Caption = "Searching..."
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ClearButton_Click()
Const ProcName As String = "ClearButton_Click"
On Error GoTo Err

If ClearButton.Caption = CancelCaption Then
    mCancelInitiatedByUser = True
    CancelSearch
Else
    Clear
    RaiseEvent Cleared
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ContractSelector1_SelectionChanged()
ActionButton.Enabled = True
End Sub

Private Sub ContractSelector1_SelectionCleared()
ActionButton.Enabled = False
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
' mContractSelectorInitialisationTC Event Handlers
'@================================================================================

Private Sub mContractSelectorInitialisationTC_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mContractSelectorInitialisationTC_Completed"
On Error GoTo Err

If ev.Cancelled Then
    Clear
    MessageLabel.Caption = "Cancelled"
    If mCancelInitiatedByUser Then
        RaiseEvent Cancelled
        mCancelInitiatedByUser = False
    End If
ElseIf ev.ErrorNumber <> 0 Then
    Clear
    
    Dim lEv As ErrorEventData
    Set lEv.Source = Me
    lEv.ErrorCode = ev.ErrorNumber
    lEv.ErrorMessage = ev.ErrorMessage
    lEv.ErrorSource = ev.ErrorSource
    RaiseEvent Error(lEv)
Else
    contractSelectorCompletion
    Dim lContractsSuppliedByCaller As Boolean
    lContractsSuppliedByCaller = CBool(ev.Cookie)
    If Not lContractsSuppliedByCaller Then RaiseEvent ContractsLoaded(ev.result)
End If

UserControl.MousePointer = MousePointerConstants.vbDefault


Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    UserControl.MousePointer = MousePointerConstants.vbDefault
    Clear
    If mCancelInitiatedByUser Then
        RaiseEvent Cancelled
        mCancelInitiatedByUser = False
    End If
ElseIf ev.Future.IsFaulted Then
    UserControl.MousePointer = MousePointerConstants.vbDefault
    Clear
    
    Dim lEv As ErrorEventData
    Set lEv.Source = Me
    lEv.ErrorCode = ev.Future.ErrorNumber
    lEv.ErrorMessage = ev.Future.ErrorMessage
    lEv.ErrorSource = ev.Future.ErrorSource
    RaiseEvent Error(lEv)
Else
    If mContractsBuilder Is Nothing Then
        If mHistoricalContractsFound Then
            MessageLabel.Caption = "No unexpired contracts"
        Else
            MessageLabel.Caption = "No contracts"
        End If
        UserControl.MousePointer = MousePointerConstants.vbDefault
        RaiseEvent NoContracts
    Else
        MessageLabel.Caption = getContractsCountMessage(mContractsBuilder.Contracts.Count)
        Set mContractSelectorInitialisationTC = handleContractsLoaded(mContractsBuilder.Contracts, False)
        If mContractSelectorInitialisationTC Is Nothing Then
            ClearButton.Caption = ClearCaption
            UserControl.MousePointer = MousePointerConstants.vbDefault
            RaiseEvent ContractsLoaded(mContractsBuilder.Contracts)
        End If
    End If
End If

Set mContractsBuilder = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let ActionButtonCaption( _
                ByVal Value As String)
ActionButton.Caption = Value
PropertyChanged PropNameActionButtonCaption
End Property

Public Property Get ActionButtonCaption() As String
ActionButtonCaption = ActionButton.Caption
End Property

Public Property Get AllowMultipleSelection() As Boolean
AllowMultipleSelection = mAllowMultipleSelection
End Property

Public Property Let AllowMultipleSelection(ByVal Value As Boolean)
mAllowMultipleSelection = Value
PropertyChanged PropNameAllowMultipleSelection
End Property

Public Property Let BackColor( _
                ByVal Value As OLE_COLOR)
UserControl.BackColor = Value
ContractSpecBuilder1.BackColor = Value
MessageLabel.BackColor = Value
PropertyChanged PropNameBackcolor
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpec
End Property

Public Property Get Cookie() As Variant
Attribute Cookie.VB_MemberFlags = "400"
Cookie = mCookie
End Property

Public Property Let ForeColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "foreColor"
On Error GoTo Err

ContractSpecBuilder1.ForeColor = Value
MessageLabel.ForeColor = Value
PropertyChanged PropNameForecolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
ForeColor = MessageLabel.ForeColor
End Property

Public Property Let IncludeHistoricalContracts( _
                ByVal Value As Boolean)
Const ProcName As String = "IncludeHistoricalContracts"
On Error GoTo Err

ContractSelector1.IncludeHistoricalContracts = Value

PropertyChanged PropNameIncludeHistoricalContracts

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

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

RowBackColorEven = ContractSelector1.RowBackColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorEven(ByVal Value As OLE_COLOR)
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

ContractSelector1.RowBackColorEven = Value
PropertyChanged PropNameRowBackColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

RowBackColorOdd = ContractSelector1.RowBackColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorOdd(ByVal Value As OLE_COLOR)
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

ContractSelector1.RowBackColorOdd = Value
PropertyChanged PropNameRowBackColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelectedContracts() As IContracts
Attribute SelectedContracts.VB_MemberFlags = "400"
Const ProcName As String = "SelectedContracts"
On Error GoTo Err

If Not mSingleContracts Is Nothing Then
    Set SelectedContracts = mSingleContracts
Else
    Set SelectedContracts = ContractSelector1.SelectedContracts
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TextBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "TextBackColor"
On Error GoTo Err

ContractSpecBuilder1.TextBackColor = Value
PropertyChanged PropNameTextBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextBackColor() As OLE_COLOR
TextBackColor = ContractSpecBuilder1.TextBackColor
End Property

Public Property Let TextForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "TextForeColor"
On Error GoTo Err

ContractSpecBuilder1.TextForeColor = Value
ContractSelector1.ForeColor = Value
PropertyChanged PropNameTextForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextForeColor() As OLE_COLOR
TextForeColor = ContractSpecBuilder1.TextForeColor
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
UserControl.ForeColor = mTheme.GridForeColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CancelSearch()
Const ProcName As String = "CancelSearch"
On Error GoTo Err

If Not mContractsBuilder Is Nothing Then
    mFutureWaiter.Cancel
ElseIf Not mContractSelectorInitialisationTC Is Nothing Then
    mContractSelectorInitialisationTC.CancelTask
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Set mSingleContracts = Nothing
ContractSpecBuilder1.Visible = True
ContractSelector1.Visible = False
ContractSelector1.Clear
ClearButton.Caption = ClearCaption
ClearButton.Visible = False
MessageLabel.Caption = ""

setActionButton

On Error Resume Next    ' SetFocus gives an error if the control is not visible
ContractSpecBuilder1.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pContractStorePrimary As IContractStore, _
                ByVal pContractStoreSecondary As IContractStore)
Set mContractStorePrimary = pContractStorePrimary
Set mContractStoreSecondary = pContractStoreSecondary
End Sub

Public Sub LoadContracts( _
                ByVal pContracts As IContracts, _
                Optional ByVal pCookie As Variant)
Const ProcName As String = "LoadContracts"
On Error GoTo Err

AssertArgument Not pContracts Is Nothing, "pContracts is Nothing"

If pContracts.Count = 1 Then Exit Sub   ' to prevent Action event firing

ActionButton.Enabled = True
ActionButton.Default = True

gSetVariant mCookie, pCookie

UserControl.MousePointer = MousePointerConstants.vbHourglass
ClearButton.Caption = CancelCaption
ClearButton.Visible = True
ClearButton.Enabled = True

ContractSelector1.Visible = True
ContractSpecBuilder1.Visible = False

Set mContractSelectorInitialisationTC = handleContractsLoaded(pContracts, True)
If mContractSelectorInitialisationTC Is Nothing Then
    ClearButton.Caption = ClearCaption
    UserControl.MousePointer = MousePointerConstants.vbDefault
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub contractSelectorCompletion()
Const ProcName As String = "contractSelectorCompletion"
On Error GoTo Err

mLoadingContracts = False
MessageLabel.Caption = getContractsCountMessage(ContractSelector1.Count)
ActionButton.Enabled = False
ClearButton.Caption = ClearCaption

On Error Resume Next    ' because SetFocus gives an error if called during a Form_Load event
ContractSelector1.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getContractsCountMessage(ByVal pCount As Long) As String
getContractsCountMessage = pCount & IIf(pCount = 1, " contract", " contracts")
End Function

Private Function handleContractsLoaded( _
                ByVal pContracts As IContracts, _
                ByVal pContractsSuppliedByCaller As Boolean) As TaskController
Const ProcName As String = "handleContractsLoaded"
On Error GoTo Err

If pContracts.Count = 1 Then
    Set mSingleContracts = pContracts
    RaiseEvent Action
    Clear
    mCookie = Empty
Else
    Set handleContractsLoaded = setupContractSelector(pContracts, mAllowMultipleSelection, pContractsSuppliedByCaller)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setActionButton()
If ContractSpecBuilder1.IsReady Then
    ActionButton.Enabled = True
    ActionButton.Default = True
Else
    ActionButton.Enabled = False
    ActionButton.Default = False
End If
End Sub

Private Function setupContractSelector( _
                ByVal pContracts As IContracts, _
                ByVal pAllowMultipleSelection As Boolean, _
                ByVal pContractsSuppliedByCaller As Boolean) As TaskController
Const ProcName As String = "setupContractSelector"
On Error GoTo Err

If pContracts.Count <= 20 Then
    ContractSelector1.Initialise pContracts, pAllowMultipleSelection
    contractSelectorCompletion
Else
    Set setupContractSelector = ContractSelector1.InitialiseAsync(pContracts, pAllowMultipleSelection, pContractsSuppliedByCaller)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
