VERSION 5.00
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
      TabIndex        =   21
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
         TabIndex        =   22
         Top             =   240
         Width           =   2535
         Begin TradeBuildUI26.ContractSpecBuilder ContractSpecBuilder1 
            Height          =   3330
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   6509
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data source"
      Height          =   735
      Left            =   2880
      TabIndex        =   18
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3495
         TabIndex        =   19
         Top             =   240
         Width           =   3495
         Begin VB.ComboBox FormatCombo 
            Height          =   315
            ItemData        =   "TickStreamSpecifier.ctx":0000
            Left            =   720
            List            =   "TickStreamSpecifier.ctx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Format"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dates/Times"
      Height          =   2895
      Left            =   2880
      TabIndex        =   10
      Top             =   720
      Width           =   3735
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   3495
         TabIndex        =   11
         Top             =   240
         Width           =   3495
         Begin VB.TextBox ToDateText 
            Height          =   285
            Left            =   2160
            TabIndex        =   3
            Top             =   120
            Width           =   1260
         End
         Begin VB.TextBox FromDateText 
            Height          =   285
            Left            =   480
            TabIndex        =   2
            Top             =   120
            Width           =   1260
         End
         Begin VB.CheckBox CompleteSessionCheck 
            Caption         =   "Complete sessions"
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   480
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.CheckBox UseExchangeTimezoneCheck 
            Caption         =   "Use exchange timezone (otherwise local time)"
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   720
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.Frame SessionTimesFrame 
            Caption         =   "Session times"
            Height          =   1095
            Left            =   0
            TabIndex        =   12
            Top             =   1200
            Width           =   3495
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   810
               Left            =   120
               ScaleHeight     =   810
               ScaleWidth      =   3285
               TabIndex        =   13
               Top             =   240
               Width           =   3285
               Begin VB.OptionButton UseContractTimesOption 
                  Caption         =   "Use contract times"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   0
                  TabIndex        =   6
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1695
               End
               Begin VB.TextBox CustomToTimeText 
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   9
                  Top             =   240
                  Width           =   660
               End
               Begin VB.TextBox CustomFromTimeText 
                  BackColor       =   &H8000000F&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   8
                  Top             =   0
                  Width           =   660
               End
               Begin VB.OptionButton UseCustomTimesOption 
                  Caption         =   "Use custom times (must be in exchange timezone)"
                  Enabled         =   0   'False
                  Height          =   495
                  Left            =   0
                  TabIndex        =   7
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "To"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   15
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "From"
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   14
                  Top             =   0
                  Width           =   495
               End
            End
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "To"
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "From"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   855
         End
      End
   End
   Begin VB.Label ErrorLabel 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   23
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

Private mSupportedTickStreamFormats()       As TickfileFormatSpecifier
Private WithEvents mContractsLoadTC         As TaskController
Attribute mContractsLoadTC.VB_VarHelpID = -1
Private mContracts                          As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mSecType                            As SecurityTypes

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
Dim failpoint As String
On Error GoTo Err

getSupportedTickstreamFormats

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CompleteSessionCheck_Click()
Const ProcName As String = "CompleteSessionCheck_Click"
Dim failpoint As String
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
checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ContractSpecBuilder1_NotReady()
Const ProcName As String = "ContractSpecBuilder1_NotReady"
Dim failpoint As String
On Error GoTo Err

RaiseEvent NotReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ContractSpecBuilder1_Ready()
Const ProcName As String = "ContractSpecBuilder1_Ready"
Dim failpoint As String
On Error GoTo Err

checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CustomFromTimeText_Change()
Const ProcName As String = "CustomFromTimeText_Change"
Dim failpoint As String
On Error GoTo Err

checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CustomToTimeText_Change()
Const ProcName As String = "CustomToTimeText_Change"
Dim failpoint As String
On Error GoTo Err

checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub FromDateText_Change()
Const ProcName As String = "FromDateText_Change"
Dim failpoint As String
On Error GoTo Err

checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ToDateText_Change()
Const ProcName As String = "ToDateText_Change"
Dim failpoint As String
On Error GoTo Err

checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UseContractTimesOption_Click()
Const ProcName As String = "UseContractTimesOption_Click"
Dim failpoint As String
On Error GoTo Err

adjustCustomTimeFieldAttributes
checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UseCustomTimesOption_Click()
Const ProcName As String = "UseCustomTimesOption_Click"
Dim failpoint As String
On Error GoTo Err

adjustCustomTimeFieldAttributes
checkReady

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mContractsLoadTC Event Handlers
'@================================================================================

Private Sub mContractsLoadTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mContractsLoadTC_Completed"
Dim failpoint As String
On Error GoTo Err

Screen.MousePointer = vbDefault
If ev.errorNumber <> 0 Then
    ErrorLabel.caption = ev.errorMessage
Else
    Set mContracts = ev.result
    processContracts
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mContractsLoadTC_Notification(ev As TWUtilities30.TaskNotificationEventData)
Const ProcName As String = "mContractsLoadTC_Notification"
Dim failpoint As String
On Error GoTo Err

ErrorLabel.caption = ev.eventMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub load()
Dim contractSpec As contractSpecifier

Const ProcName As String = "load"
Dim failpoint As String
On Error GoTo Err

ErrorLabel.caption = ""

Screen.MousePointer = vbHourglass

Set contractSpec = ContractSpecBuilder1.contractSpecifier
mSecType = contractSpec.secType

Set mContractsLoadTC = TradeBuildAPI.LoadContracts(contractSpec)

Exit Sub

Err:
Screen.MousePointer = vbDefault
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    ErrorLabel.caption = Err.Description
    Exit Sub
End If
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub adjustCustomTimeFieldAttributes()
Const ProcName As String = "adjustCustomTimeFieldAttributes"
Dim failpoint As String
On Error GoTo Err

If UseCustomTimesOption Then
    enableCustomTimeFields
Else
    disableCustomTimeFields
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function checkOk() As Boolean

Const ProcName As String = "checkOk"
Dim failpoint As String
On Error GoTo Err

If FormatCombo.ListCount = 0 Then Exit Function

If Not ContractSpecBuilder1.isReady Then Exit Function

If Not IsDate(FromDateText.Text) Then Exit Function
If CompleteSessionCheck.value = vbUnchecked And Not IsDate(ToDateText.Text) Then Exit Function
If CompleteSessionCheck.value = vbChecked And _
    ToDateText.Text <> "" And _
    Not IsDate(ToDateText.Text) Then Exit Function
If IsDate(ToDateText.Text) Then
    If CDate(FromDateText.Text) > CDate(ToDateText.Text) Then Exit Function
End If

If UseCustomTimesOption Then
    If Not IsDate(CustomFromTimeText) Then Exit Function
    If Not IsDate(CustomToTimeText) Then Exit Function
    If CDbl(CDate(CustomFromTimeText)) >= 1# Then Exit Function
    If CDbl(CDate(CustomToTimeText)) >= 1# Then Exit Function
End If

checkOk = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Function

Private Sub checkReady()
Const ProcName As String = "checkReady"
Dim failpoint As String
On Error GoTo Err

If checkOk Then
    RaiseEvent Ready
Else
    RaiseEvent NotReady
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub disableCustomTimeFields()
Const ProcName As String = "disableCustomTimeFields"
Dim failpoint As String
On Error GoTo Err

CustomFromTimeText.Enabled = False
CustomFromTimeText.backColor = vbButtonFace
CustomToTimeText.Enabled = False
CustomToTimeText.backColor = vbButtonFace

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub enableCustomTimeFields()
Const ProcName As String = "enableCustomTimeFields"
Dim failpoint As String
On Error GoTo Err

CustomFromTimeText.Enabled = True
CustomFromTimeText.backColor = vbWindowBackground
CustomToTimeText.Enabled = True
CustomToTimeText.backColor = vbWindowBackground

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub getSupportedTickstreamFormats()
Dim tff() As TickfileFormatSpecifier
Dim i As Long
Dim j As Long

On Error GoTo Err

tff = TradeBuildAPI.SupportedInputTickfileFormats

ReDim mSupportedTickStreamFormats(9) As TickfileFormatSpecifier
j = -1

For i = 0 To UBound(tff)
    If tff(i).FormatType = StreamBased Then
        j = j + 1
        If j > UBound(mSupportedTickStreamFormats) Then
            ReDim Preserve mSupportedTickStreamFormats(UBound(mSupportedTickStreamFormats) + 9) As TickfileFormatSpecifier
        End If
        mSupportedTickStreamFormats(j) = tff(i)
        FormatCombo.addItem mSupportedTickStreamFormats(j).name
    End If
Next

FormatCombo.ListIndex = 0

If j = -1 Then
    Erase mSupportedTickStreamFormats
Else
    ReDim Preserve mSupportedTickStreamFormats(j) As TickfileFormatSpecifier
End If

Exit Sub

Err:

End Sub

Private Sub processContracts()
Dim lTickfileSpecifiers As TickfileSpecifiers
Dim k As Long
Dim TickfileFormatID As String

On Error GoTo Err

If mContracts.Count = 0 Then
    ErrorLabel.caption = "No contracts meet this specification"
    Exit Sub
End If

If mSecType <> SecurityTypes.SecTypeFuture And _
    mSecType <> SecurityTypes.SecTypeOption And _
    mSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    If mContracts.Count > 1 Then
        ' don't see how this can happen, but just in case!
        ErrorLabel.caption = "More than one contract meets this specification"
        Exit Sub
    End If
End If
    
For k = 0 To UBound(mSupportedTickStreamFormats)
    If mSupportedTickStreamFormats(k).name = FormatCombo.Text Then
        TickfileFormatID = mSupportedTickStreamFormats(k).FormalID
        Exit For
    End If
Next

Set lTickfileSpecifiers = TradeBuildAPI.GenerateTickfileSpecifiers( _
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
ErrorLabel.caption = Err.Description

End Sub





