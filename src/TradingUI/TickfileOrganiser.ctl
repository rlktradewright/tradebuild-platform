VERSION 5.00
Begin VB.UserControl TickfileOrganiser 
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   ScaleHeight     =   2295
   ScaleWidth      =   8400
   Begin TradingUI27.TickfileChooser TickfileChooser1 
      Left            =   6720
      Top             =   1440
      _ExtentX        =   1296
      _ExtentY        =   873
   End
   Begin TradingUI27.TickfileListManager TickfileListManager1 
      Height          =   1935
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
   End
   Begin VB.CommandButton ClearButton 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton AddTickfilesButton 
      Caption         =   "Add &files..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton AddTickstreamsButton 
      Caption         =   "Add &streams..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "TickfileOrganiser"
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

Event TickfileCountChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickfileOrganiser"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileStore                              As ITickfileStore
Private mPrimaryContractStore                       As IContractStore
Private mSecondaryContractStore                     As IContractStore

Private mEnabled                                    As Boolean

Private mMinimumHeight                              As Long
Private mMinimumWidth                               As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

mMinimumHeight = TickfileListManager1.MinimumHeight + 105 + ClearButton.Height
mMinimumWidth = ClearButton.Width + 105 + AddTickfilesButton.Width + 105 + AddTickstreamsButton.Width + 315

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

If UserControl.Width < mMinimumWidth Then UserControl.Width = mMinimumWidth
If UserControl.Height < mMinimumHeight Then UserControl.Height = mMinimumHeight

TickfileListManager1.Width = UserControl.Width
TickfileListManager1.Height = UserControl.Height - AddTickfilesButton.Height - 105

ClearButton.Top = UserControl.Height - ClearButton.Height
AddTickfilesButton.Top = UserControl.Height - AddTickfilesButton.Height
AddTickstreamsButton.Top = UserControl.Height - AddTickstreamsButton.Height

AddTickstreamsButton.Left = UserControl.Width - 315 - AddTickstreamsButton.Width
AddTickfilesButton.Left = AddTickstreamsButton.Left - 105 - AddTickfilesButton.Width

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddTickfilesButton_Click()
Const ProcName As String = "AddTickfilesButton_Click"
On Error GoTo Err

Dim tickfileNames() As String
tickfileNames = TickfileChooser1.ChooseTickfiles

If TickfileChooser1.cancelled Then Exit Sub

TickfileListManager1.AddTickfileNames tickfileNames

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub AddTickstreamsButton_Click()
Const ProcName As String = "AddTickstreamsButton_Click"
On Error GoTo Err

Static sTickstreamSpecifier As fTickStreamSpecifier

If sTickstreamSpecifier Is Nothing Then
    Set sTickstreamSpecifier = New fTickStreamSpecifier
    sTickstreamSpecifier.Initialise mTickfileStore, mPrimaryContractStore, mSecondaryContractStore
End If

sTickstreamSpecifier.Show vbModal

If sTickstreamSpecifier.cancelled Then Exit Sub

TickfileListManager1.AddTickfileSpecifiers sTickstreamSpecifier.TickfileSpecifiers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ClearButton_Click()
Const ProcName As String = "ClearButton_Click"

On Error GoTo Err

TickfileListManager1.Clear

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickfileListManager1_TickfileCountChanged()
Const ProcName As String = "TickfileListManager1_TickfileCountChanged"
On Error GoTo Err

If TickfileListManager1.TickfileCount > 0 Then
    ClearButton.Enabled = True
Else
    ClearButton.Enabled = False
End If

RaiseEvent TickfileCountChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Enabled(ByVal value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

mEnabled = value
If mEnabled Then
    enableAddButtons
    ClearButton.Enabled = (TickfileListManager1.TickfileCount > 0)
Else
    AddTickfilesButton.Enabled = False
    AddTickstreamsButton.Enabled = False
    ClearButton.Enabled = False
End If
TickfileListManager1.Enabled = mEnabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
Enabled = mEnabled
End Property

Public Property Let ListIndex(ByVal value As Long)
Const ProcName As String = "ListIndex"
On Error GoTo Err

TickfileListManager1.ListIndex = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ListIndex() As Long
Const ProcName As String = "ListIndex"
On Error GoTo Err

ListIndex = TickfileListManager1.ListIndex

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MinimumHeight() As Long
MinimumHeight = mMinimumHeight
End Property

Public Property Get MinimumWidth() As Long
MinimumWidth = mMinimumWidth
End Property

Public Property Get TickfileCount() As Long
Const ProcName As String = "TickfileCount"
On Error GoTo Err

TickfileCount = TickfileListManager1.TickfileCount

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifiers
Const ProcName As String = "TickfileSpecifiers"
On Error GoTo Err

Set TickfileSpecifiers = TickfileListManager1.TickfileSpecifiers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

TickfileListManager1.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileStore = pTickfileStore
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore

TickfileListManager1.Initialise mTickfileStore
TickfileChooser1.Initialise mTickfileStore

If Enabled Then enableAddButtons

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub enableAddButtons()
Const ProcName As String = "enableAddButtons"
On Error GoTo Err

If TickfileListManager1.SupportsTickFiles Then AddTickfilesButton.Enabled = True
If TickfileListManager1.SupportsTickStreams Then AddTickstreamsButton.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


