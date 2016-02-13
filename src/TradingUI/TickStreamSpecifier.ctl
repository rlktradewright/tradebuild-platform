VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.UserControl TickStreamSpecifier 
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ScaleHeight     =   4200
   ScaleWidth      =   6720
   Begin VB.Frame Frame2 
      Caption         =   "Contract specification"
      Height          =   3615
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   120
         ScaleHeight     =   3255
         ScaleWidth      =   2535
         TabIndex        =   21
         Top             =   240
         Width           =   2535
         Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
            Height          =   3690
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   5556
            ForeColor       =   -2147483640
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data source"
      Height          =   735
      Left            =   2880
      TabIndex        =   17
      Top             =   2880
      Width           =   3735
      Begin VB.PictureBox DataSourcePicture 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3495
         TabIndex        =   18
         Top             =   240
         Width           =   3495
         Begin TWControls40.TWImageCombo FormatCombo 
            Height          =   270
            Left            =   720
            TabIndex        =   23
            Top             =   0
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   476
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "TickStreamSpecifier.ctx":0000
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Format"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dates/Times"
      Height          =   2775
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox DatesTimesPicture 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   3495
         TabIndex        =   10
         Top             =   240
         Width           =   3495
         Begin VB.TextBox ToDateText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   2160
            TabIndex        =   2
            Top             =   120
            Width           =   1260
         End
         Begin VB.TextBox FromDateText 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   480
            TabIndex        =   1
            Top             =   120
            Width           =   1260
         End
         Begin VB.CheckBox CompleteSessionCheck 
            Appearance      =   0  'Flat
            Caption         =   "Complete sessions"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   480
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.CheckBox UseExchangeTimezoneCheck 
            Appearance      =   0  'Flat
            Caption         =   "Use exchange timezone (otherwise local time)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   720
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.Frame SessionTimesFrame 
            Caption         =   "Session times"
            Height          =   1215
            Left            =   0
            TabIndex        =   11
            Top             =   1200
            Width           =   3495
            Begin VB.PictureBox SessionTimesPicture 
               BorderStyle     =   0  'None
               Height          =   930
               Left            =   120
               ScaleHeight     =   930
               ScaleWidth      =   3285
               TabIndex        =   12
               Top             =   240
               Width           =   3285
               Begin VB.OptionButton UseContractTimesOption 
                  Appearance      =   0  'Flat
                  Caption         =   "Use contract times"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   0
                  TabIndex        =   5
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1695
               End
               Begin VB.TextBox CustomToTimeText 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   8
                  Top             =   600
                  Width           =   660
               End
               Begin VB.TextBox CustomFromTimeText 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   7
                  Top             =   360
                  Width           =   660
               End
               Begin VB.OptionButton UseCustomTimesOption 
                  Appearance      =   0  'Flat
                  Caption         =   "Use custom times (must be in exchange timezone)"
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   615
                  Left            =   0
                  TabIndex        =   6
                  Top             =   240
                  Width           =   1815
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "To"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   14
                  Top             =   600
                  Width           =   495
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "From"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   13
                  Top             =   360
                  Width           =   495
               End
            End
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "To"
            Height          =   255
            Left            =   1800
            TabIndex        =   16
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "From"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin VB.Label ErrorLabel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   3720
      Width           =   6615
   End
End
Attribute VB_Name = "TickStreamSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event NotReady()
Event Ready()
Event TickStreamsSpecified(ByVal pTickfileSpecifiers As TickfileSpecifiers)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "TickStreamSpecifier"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileStore                      As ITickfileStore
Private mPrimaryContractStore               As IContractStore
Private mSecondaryContractStore             As IContractStore

Private mSupportedTickStreamFormats()       As TickfileFormatSpecifier
Private mContracts                          As IContracts
Attribute mContracts.VB_VarHelpID = -1

Private mSecType                            As SecurityTypes

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTheme                              As ITheme

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_EnterFocus()
Const ProcName As String = "UserControl_EnterFocus"
On Error GoTo Err

ContractSpecBuilder1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 4200
UserControl.Width = 6720
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

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CompleteSessionCheck_Click()
Const ProcName As String = "CompleteSessionCheck_Click"
On Error GoTo Err

If CompleteSessionCheck = vbChecked Then
    UseContractTimesOption.Enabled = True
    UseCustomTimesOption.Enabled = True
    UseExchangeTimezoneCheck.Enabled = False
Else
    UseContractTimesOption.Enabled = False
    UseCustomTimesOption.Enabled = False
    UseExchangeTimezoneCheck.Enabled = True
End If
adjustCustomTimeFieldAttributes
checkReady True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ContractSpecBuilder1_NotReady()
Const ProcName As String = "ContractSpecBuilder1_NotReady"
On Error GoTo Err

checkReady True
RaiseEvent NotReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ContractSpecBuilder1_Ready()
Const ProcName As String = "ContractSpecBuilder1_Ready"
On Error GoTo Err

checkReady True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CustomFromTimeText_Change()
Const ProcName As String = "CustomFromTimeText_Change"
On Error GoTo Err

checkReady False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub CustomFromTimeText_GotFocus()
Const ProcName As String = "CustomFromTimeText_GotFocus"
On Error GoTo Err

CustomFromTimeText.SelStart = 0
CustomFromTimeText.SelLength = Len(CustomFromTimeText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CustomFromTimeText_Validate(Cancel As Boolean)
Const ProcName As String = "CustomFromTimeText_Validate"
On Error GoTo Err

checkReady True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub CustomToTimeText_Change()
Const ProcName As String = "CustomToTimeText_Change"
On Error GoTo Err

checkReady False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub CustomToTimeText_GotFocus()
Const ProcName As String = "CustomToTimeText_GotFocus"
On Error GoTo Err

CustomToTimeText.SelStart = 0
CustomToTimeText.SelLength = Len(CustomToTimeText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub CustomToTimeText_Validate(Cancel As Boolean)
Const ProcName As String = "CustomToTimeText_Validate"
On Error GoTo Err

checkReady True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FormatCombo_GotFocus()
Const ProcName As String = "FormatCombo_GotFocus"
On Error GoTo Err

FormatCombo.SelStart = 1
FormatCombo.SelLength = Len(FormatCombo.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FromDateText_Change()
Const ProcName As String = "FromDateText_Change"
On Error GoTo Err

checkReady False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FromDateText_GotFocus()
Const ProcName As String = "FromDateText_GotFocus"
On Error GoTo Err

FromDateText.SelStart = 0
FromDateText.SelLength = Len(FromDateText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FromDateText_Validate(Cancel As Boolean)
Const ProcName As String = "FromDateText_Validate"
On Error GoTo Err

checkReady True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ToDateText_Change()
Const ProcName As String = "ToDateText_Change"
On Error GoTo Err

checkReady False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ToDateText_GotFocus()
Const ProcName As String = "ToDateText_GotFocus"
On Error GoTo Err

ToDateText.SelStart = 0
ToDateText.SelLength = Len(ToDateText.Text)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ToDateText_Validate(Cancel As Boolean)
Const ProcName As String = "ToDateText_Validate"
On Error GoTo Err

checkReady True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UseContractTimesOption_Click()
Const ProcName As String = "UseContractTimesOption_Click"
On Error GoTo Err

adjustCustomTimeFieldAttributes
checkReady True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UseCustomTimesOption_Click()
Const ProcName As String = "UseCustomTimesOption_Click"
On Error GoTo Err

adjustCustomTimeFieldAttributes
checkReady True

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

Screen.MousePointer = vbDefault

If ev.Future.IsFaulted <> 0 Then
    ErrorLabel.Caption = ev.Future.ErrorMessage
ElseIf ev.Future.IsCancelled <> 0 Then
    ErrorLabel.Caption = "Contracts fetch Cancelled"
Else
    Set mContracts = ev.Future.value
    processContracts
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls
ErrorLabel.ForeColor = vbRed
ErrorLabel.FontBold = True

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
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileStore = pTickfileStore
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
getSupportedTickstreamFormats

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Load()
Const ProcName As String = "Load"
On Error GoTo Err

ErrorLabel.Caption = ""

Screen.MousePointer = vbHourglass

Dim contractSpec As IContractSpecifier
Set contractSpec = ContractSpecBuilder1.ContractSpecifier
mSecType = contractSpec.secType

Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add FetchContracts(contractSpec, mPrimaryContractStore, mSecondaryContractStore)

Exit Sub

Err:
Screen.MousePointer = vbDefault
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    ErrorLabel.Caption = Err.Description
    Exit Sub
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub adjustCustomTimeFieldAttributes()
Const ProcName As String = "adjustCustomTimeFieldAttributes"
On Error GoTo Err

If UseCustomTimesOption Then
    enableCustomTimeFields
Else
    disableCustomTimeFields
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function checkOk(ByRef pMessage As String) As Boolean
Const ProcName As String = "checkOk"
On Error GoTo Err

If FormatCombo.ComboItems.Count = 0 Then
    pMessage = "No formats available"
    Exit Function
End If

If Not ContractSpecBuilder1.IsReady Then
    pMessage = "Contract specifier invalid"
    Exit Function
End If

If Not IsDate(FromDateText.Text) Then
    pMessage = "'From' is not a valid datetime"
    Exit Function
End If

If ToDateText.Text <> "" And Not IsDate(ToDateText.Text) Then
    pMessage = "'To' is not a valid datetime"
    Exit Function
End If

If IsDate(ToDateText.Text) Then
    If CDate(FromDateText.Text) > CDate(ToDateText.Text) Then
        pMessage = "'From' cannot be later than 'To'"
        Exit Function
    End If
End If

If UseCustomTimesOption Then
    If Not IsDate(CustomFromTimeText) Then
        pMessage = "Custom 'From' is not a valid time"
        Exit Function
    End If
    If Not IsDate(CustomToTimeText) Then
        pMessage = "Custom 'To' is not a valid time"
        Exit Function
    End If
    If CDbl(CDate(CustomFromTimeText)) >= 1# Then
        pMessage = "Custom 'From' must be a time only"
        Exit Function
    End If
    If CDbl(CDate(CustomToTimeText)) >= 1# Then
        pMessage = "Custom 'To' must be a time only"
        Exit Function
    End If
End If

checkOk = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub checkReady(ByVal pShowErrors As Boolean)
Const ProcName As String = "checkReady"
On Error GoTo Err

Dim lMsg As String
If checkOk(lMsg) Then
    ErrorLabel.Caption = ""
    RaiseEvent Ready
Else
    If pShowErrors Then ErrorLabel.Caption = lMsg
    RaiseEvent NotReady
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub disableCustomTimeFields()
Const ProcName As String = "disableCustomTimeFields"
On Error GoTo Err

CustomFromTimeText.Enabled = False
CustomToTimeText.Enabled = False

If mTheme Is Nothing Then
    CustomFromTimeText.BackColor = vbButtonFace
    CustomToTimeText.BackColor = vbButtonFace
Else
    ' leave the themed colours unchanged
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableCustomTimeFields()
Const ProcName As String = "enableCustomTimeFields"
On Error GoTo Err

CustomFromTimeText.Enabled = True
CustomToTimeText.Enabled = True

If mTheme Is Nothing Then
    CustomFromTimeText.BackColor = vbWindowBackground
    CustomToTimeText.BackColor = vbWindowBackground
Else
    ' leave the themed colours unchanged
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getSupportedTickstreamFormats()
Const ProcName As String = "getSupportedTickstreamFormats"
On Error GoTo Err

Dim tff() As TickfileFormatSpecifier
tff = mTickfileStore.SupportedFormats

ReDim mSupportedTickStreamFormats(9) As TickfileFormatSpecifier

Dim j As Long
j = -1

Dim i As Long
For i = 0 To UBound(tff)
    If tff(i).FormatType = TickfileModeStreamBased Then
        j = j + 1
        If j > UBound(mSupportedTickStreamFormats) Then
            ReDim Preserve mSupportedTickStreamFormats(UBound(mSupportedTickStreamFormats) + 9) As TickfileFormatSpecifier
        End If
        mSupportedTickStreamFormats(j) = tff(i)
        FormatCombo.ComboItems.Add , , mSupportedTickStreamFormats(j).Name
    End If
Next

Set FormatCombo.SelectedItem = FormatCombo.ComboItems(1)

If j = -1 Then
    Erase mSupportedTickStreamFormats
Else
    ReDim Preserve mSupportedTickStreamFormats(j) As TickfileFormatSpecifier
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContracts()
On Error GoTo Err

If mContracts.Count = 0 Then
    ErrorLabel.Caption = "No contracts meet this specification"
    Exit Sub
End If

If mSecType <> SecurityTypes.SecTypeFuture And _
    mSecType <> SecurityTypes.SecTypeOption And _
    mSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    If mContracts.Count > 1 Then
        ' don't see how this can happen, but just in case!
        ErrorLabel.Caption = "More than one contract meets this specification"
        Exit Sub
    End If
End If
    
Dim TickfileFormatID As String
Dim k As Long
For k = 0 To UBound(mSupportedTickStreamFormats)
    If mSupportedTickStreamFormats(k).Name = FormatCombo.Text Then
        TickfileFormatID = mSupportedTickStreamFormats(k).FormalID
        Exit For
    End If
Next

Dim lTickfileSpecifiers As TickfileSpecifiers
Set lTickfileSpecifiers = GenerateTickfileSpecifiers( _
                                                mContracts, _
                                                TickfileFormatID, _
                                                CDate(FromDateText), _
                                                CDate(IIf(ToDateText <> "", ToDateText, 0)), _
                                                CompleteSessionCheck = vbChecked, _
                                                UseExchangeTimezoneCheck = vbChecked, _
                                                CDate(IIf(CustomFromTimeText <> "", CustomFromTimeText, 0)), _
                                                CDate(IIf(CustomToTimeText <> "", CustomToTimeText, 0)))

RaiseEvent TickStreamsSpecified(lTickfileSpecifiers)
Exit Sub

Err:
ErrorLabel.Caption = Err.Description

End Sub





