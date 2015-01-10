VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#24.2#0"; "TWControls40.ocx"
Begin VB.UserControl ContractSearch 
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4800
   Begin TWControls40.TWButton ClearButton 
      Height          =   330
      Left            =   1680
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
   End
   Begin TWControls40.TWButton ActionButton 
      Height          =   330
      Left            =   3120
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
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
      Caption         =   "Command1"
   End
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   3863
      _ExtentY        =   5556
      ForeColor       =   -2147483640
   End
   Begin TradingUI27.ContractSelector ContractSelector1 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _ExtentX        =   0
      _ExtentY        =   0
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

Implements IThemeable

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

Private Const PropNameActionButtonCaption           As String = "ActionButtonCaption"
Private Const PropNameAllowMultipleSelection        As String = "AllowMultipleSelection"
Private Const PropNameBackcolor                     As String = "BackColor"
Private Const PropNameForecolor                     As String = "ForeColor"
Private Const PropNameIncludeHistoricalContracts    As String = "IncludeHistoricalContracts"
Private Const PropNameRowBackColorEven              As String = "RowBackColorEven"
Private Const PropNameRowBackColorOdd               As String = "RowBackColorOdd"
Private Const PropNameTextBackColor                 As String = "TextBackColor"
Private Const PropNameTextForeColor                 As String = "TextForeColor"

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

Private mCookie                                     As Variant

Private mTheme                                      As ITheme

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
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

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
PropertyChanged PropNameActionButtonCaption
End Property

Public Property Get ActionButtonCaption() As String
ActionButtonCaption = ActionButton.caption
End Property

Public Property Get AllowMultipleSelection() As Boolean
AllowMultipleSelection = mAllowMultipleSelection
End Property

Public Property Let AllowMultipleSelection(ByVal value As Boolean)
mAllowMultipleSelection = value
PropertyChanged PropNameAllowMultipleSelection
End Property

Public Property Let BackColor( _
                ByVal value As OLE_COLOR)
UserControl.BackColor = value
ContractSpecBuilder1.BackColor = value
MessageLabel.BackColor = value
PropertyChanged PropNameBackcolor
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = UserControl.BackColor
End Property

Public Property Get Cookie() As Variant
Attribute Cookie.VB_MemberFlags = "400"
Cookie = mCookie
End Property

Public Property Let ForeColor( _
                ByVal value As OLE_COLOR)
Const ProcName As String = "foreColor"
On Error GoTo Err

ContractSpecBuilder1.ForeColor = value
MessageLabel.ForeColor = value
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
                ByVal value As Boolean)
Const ProcName As String = "IncludeHistoricalContracts"
On Error GoTo Err

ContractSelector1.IncludeHistoricalContracts = value

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

Public Property Get RowBackColorEven() As OLE_COLOR
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

RowBackColorEven = ContractSelector1.RowBackColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorEven(ByVal value As OLE_COLOR)
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

ContractSelector1.RowBackColorEven = value
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

Public Property Let RowBackColorOdd(ByVal value As OLE_COLOR)
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

ContractSelector1.RowBackColorOdd = value
PropertyChanged PropNameRowBackColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelectedContracts() As IContracts
Attribute SelectedContracts.VB_MemberFlags = "400"
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

Public Property Let TextBackColor(ByVal value As OLE_COLOR)
Const ProcName As String = "TextBackColor"
On Error GoTo Err

ContractSpecBuilder1.TextBackColor = value
PropertyChanged PropNameTextBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextBackColor() As OLE_COLOR
TextBackColor = ContractSpecBuilder1.TextBackColor
End Property

Public Property Let TextForeColor(ByVal value As OLE_COLOR)
Const ProcName As String = "TextForeColor"
On Error GoTo Err

ContractSpecBuilder1.TextForeColor = value
ContractSelector1.ForeColor = value
PropertyChanged PropNameTextForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextForeColor() As OLE_COLOR
TextForeColor = ContractSpecBuilder1.TextForeColor
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
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

ActionButton.Enabled = True
ActionButton.Default = True

Set mContracts = pContracts
gSetVariant mCookie, pCookie

handleContractsLoaded

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
        mCookie = Empty
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
