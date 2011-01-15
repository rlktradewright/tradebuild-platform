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
   Begin TradeBuildUI26.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6509
   End
   Begin TradeBuildUI26.ContractSelector ContractSelector1 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5741
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

Event Error( _
                ByVal errorNumber, _
                ByVal errorMessage As String)

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

Private WithEvents mContractsLoadTC                 As TaskController
Attribute mContractsLoadTC.VB_VarHelpID = -1
Private mContracts                                  As Contracts

Private mLoadingContracts                           As Boolean

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
ContractSpecBuilder1.Visible = True
ContractSelector1.Visible = False
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

ActionButtonCaption = "Start"
IncludeHistoricalContracts = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

ActionButtonCaption = PropBag.ReadProperty("ActionButtonCaption ", "Start")
If Err.Number <> 0 Then
    ActionButtonCaption = "Start"
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
Dim failpoint As String
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
PropBag.WriteProperty "IncludeHistoricalContracts", IncludeHistoricalContracts, "False"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ActionButton_Click()
Const ProcName As String = "ActionButton_Click"
Dim failpoint As String
On Error GoTo Err

If ContractSelector1.Visible Then
    RaiseEvent Action
Else
    Set mContractsLoadTC = TradeBuildAPI.LoadContracts(ContractSpecBuilder1.contractSpecifier)
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
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mContractsLoadTC_Completed"
Dim failpoint As String
On Error GoTo Err

MessageLabel.caption = ""
UserControl.MousePointer = MousePointerConstants.vbDefault
ActionButton.Enabled = True
mLoadingContracts = False

If ContractSpecBuilder1.isReady Then
    ActionButton.Enabled = True
    ActionButton.Default = True
Else
    ActionButton.Enabled = False
    ActionButton.Default = False
End If

If ev.errorNumber <> 0 Then
    RaiseEvent Error(ev.errorNumber, ev.errorMessage)
Else
    Set mContracts = ev.result
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
End Property

Public Property Get ActionButtonCaption() As String
ActionButtonCaption = ActionButton.caption
End Property

Public Property Let IncludeHistoricalContracts( _
                ByVal value As Boolean)
Const ProcName As String = "IncludeHistoricalContracts"
Dim failpoint As String
On Error GoTo Err

ContractSelector1.IncludeHistoricalContracts = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeHistoricalContracts() As Boolean
Const ProcName As String = "IncludeHistoricalContracts"
Dim failpoint As String
On Error GoTo Err

IncludeHistoricalContracts = ContractSelector1.IncludeHistoricalContracts

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get SelectedContracts() As Contracts
Const ProcName As String = "SelectedContracts"
Dim failpoint As String
On Error GoTo Err

If mContracts.Count = 1 Then
    Set SelectedContracts = mContracts
Else
    Set SelectedContracts = ContractSelector1.SelectedContracts
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleContractsLoaded()
Const ProcName As String = "handleContractsLoaded"
Dim failpoint As String
On Error GoTo Err

If mContracts.Count = 0 Then
    MessageLabel.caption = "No contracts"
    RaiseEvent NoContracts
ElseIf mContracts.Count = 1 Then
    RaiseEvent Action
Else
    MessageLabel.caption = mContracts.Count & IIf(mContracts.Count = 1, " contract found", " contracts found")
    ContractSelector1.Initialise mContracts
    ContractSelector1.Visible = True
    ContractSpecBuilder1.Visible = False
    ContractSelector1.SetFocus
    ClearButton.Visible = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


