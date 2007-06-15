VERSION 5.00
Begin VB.Form fTickfileSpecifier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Specify tickfile"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame SessionTimesFrame 
      Caption         =   "Session times"
      Height          =   855
      Left            =   3000
      TabIndex        =   21
      Top             =   2520
      Width           =   3735
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   3495
         TabIndex        =   22
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton UseContractTimesOption 
            Caption         =   "Use contract times"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.TextBox CustomToTimeText 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   660
         End
         Begin VB.TextBox CustomFromTimeText 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   7
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton UseCustomTimesOption 
            Caption         =   "Use custom times"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "To"
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "From"
            Height          =   255
            Left            =   2040
            TabIndex        =   23
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dates/Times"
      Height          =   1455
      Left            =   3000
      TabIndex        =   16
      Top             =   960
      Width           =   3735
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   3495
         TabIndex        =   17
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox UseExchangeTimezoneCheck 
            Caption         =   "Use exchange timezone"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox CompleteSessionCheck 
            Caption         =   "Complete sessions"
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   840
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.TextBox FromText 
            Height          =   285
            Left            =   480
            TabIndex        =   3
            Top             =   480
            Width           =   1260
         End
         Begin VB.TextBox ToText 
            Height          =   285
            Left            =   2160
            TabIndex        =   4
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label8 
            Caption         =   "From"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "To"
            Height          =   255
            Left            =   1800
            TabIndex        =   18
            Top             =   480
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data source"
      Height          =   735
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   3495
         TabIndex        =   14
         Top             =   240
         Width           =   3495
         Begin VB.ComboBox FormatCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0000
            Left            =   720
            List            =   "fTickfileSpecifier.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Format"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contract specification"
      Height          =   3255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   2535
         TabIndex        =   12
         Top             =   240
         Width           =   2535
         Begin TradeBuildUI26.ContractSpecBuilder ContractSpecBuilder1 
            Height          =   2895
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   5106
         End
      End
   End
   Begin VB.Label ErrorLabel 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   7455
   End
End
Attribute VB_Name = "fTickfileSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Event TickfilesSpecified(ByRef pTickfileSpecifier() As TickfileSPecifier)

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSupportedTickfileFormats() As TickfileFormatSpecifier
Private WithEvents mContracts       As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mSecType                    As SecurityTypes

'@================================================================================
' Form Event Handlers
'@================================================================================

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub CompleteSessionCheck_Click()
If CompleteSessionCheck = vbChecked Then
    UseContractTimesOption.Enabled = True
    UseCustomTimesOption.Enabled = True
Else
    UseContractTimesOption.Enabled = False
    UseCustomTimesOption.Enabled = False
End If
adjustCustomTimeFieldAttributes
checkOk
End Sub

Private Sub ContractSpecBuilder1_NotReady()
OkButton.Enabled = False
End Sub

Private Sub ContractSpecBuilder1_Ready()
checkOk
End Sub

Private Sub CustomFromTimeText_Change()
checkOk
End Sub

Private Sub CustomToTimeText_Change()
checkOk
End Sub

Private Sub FromText_Change()
checkOk
End Sub

Private Sub OkButton_Click()
Dim contractSpec As contractSpecifier
Dim lContractsBuilder As ContractsBuilder

On Error GoTo Err

ErrorLabel.caption = ""

Screen.MousePointer = vbHourglass

Set contractSpec = ContractSpecBuilder1.contractSpecifier
mSecType = contractSpec.sectype

Set lContractsBuilder = CreateContractsBuilder(contractSpec)
Set mContracts = lContractsBuilder.Contracts
TradeBuildAPI.loadContracts lContractsBuilder

Exit Sub

Err:

Screen.MousePointer = vbDefault
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    ErrorLabel.caption = Err.Description
Else
    Err.Raise Err.Number
End If
    
End Sub

Private Sub ToText_Change()
checkOk
End Sub

Private Sub UseContractTimesOption_Click()
adjustCustomTimeFieldAttributes
checkOk
End Sub

Private Sub UseCustomTimesOption_Click()
adjustCustomTimeFieldAttributes
checkOk
End Sub

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
Screen.MousePointer = vbDefault
ErrorLabel.caption = "Invalid contract specification:" & reason
End Sub

Private Sub mContracts_NoMoreContractDetails()
Dim lTickfileSpecifiers() As TickfileSPecifier
Dim i As Long
Dim j As Long
Dim k As Long
Dim lContract As Contract
Dim lSessionBuilder As SessionBuilder
Dim lSession As session
Dim currContract As Contract
Dim sessionStartTime As Date
Dim sessionEndTime As Date
Dim fromSessionStart As Date
Dim fromSessionEnd As Date
Dim toSessionStart As Date
Dim toSessionEnd As Date
Dim endTime As Date
Dim TickfileFormatID As String

Screen.MousePointer = vbDefault
If mContracts.count = 0 Then
    ErrorLabel.caption = "No contracts meet this specification"
    Exit Sub
End If

If mSecType <> SecurityTypes.SecTypeFuture And _
    mSecType <> SecurityTypes.SecTypeOption And _
    mSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    If mContracts.count > 1 Then
        ' don't see how this can happen, but just in case!
        ErrorLabel.caption = "More than one contract meets this specification"
        Exit Sub
    End If
End If
    
For k = 0 To UBound(mSupportedTickfileFormats)
    If mSupportedTickfileFormats(k).name = FormatCombo.Text Then
        TickfileFormatID = mSupportedTickfileFormats(k).FormalID
        Exit For
    End If
Next

' get the most recent contract (though they should all have the same
'info regarding session times)
Set lContract = mContracts(mContracts.count)

Set lSessionBuilder = New SessionBuilder
Set lSession = lSessionBuilder.session

If UseCustomTimesOption Then
    lSessionBuilder.sessionStartTime = CDate(CustomFromTimeText)
    lSessionBuilder.sessionEndTime = CDate(CustomToTimeText)
Else
    lSessionBuilder.sessionStartTime = lContract.sessionStartTime
    lSessionBuilder.sessionEndTime = lContract.sessionEndTime
End If

sessionStartTime = lContract.sessionStartTime
sessionEndTime = lContract.sessionEndTime

' datetime=sessionstart(fromdatetime)
lSession.SessionTimes CDate(FromText), _
                    fromSessionStart, _
                    fromSessionEnd

If ToText <> "" Then
    lSession.SessionTimes CDate(ToText), toSessionStart, toSessionEnd
    If CompleteSessionCheck.value = vbChecked Then
        endTime = toSessionEnd
    Else
        endTime = CDate(ToText)
    End If
Else
    toSessionStart = fromSessionStart
    toSessionEnd = fromSessionEnd
    endTime = toSessionEnd
End If

' find contract for datetime
Dim aContract As Contract

If mSecType <> SecurityTypes.SecTypeFuture And _
    mSecType <> SecurityTypes.SecTypeOption And _
    mSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    Set currContract = mContracts(1)
Else
    For i = 1 To mContracts.count
        Set aContract = mContracts(i)
        If DateValue(fromSessionStart) <= _
            (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
        Then
            Set currContract = aContract
            Exit For
        End If
    Next
    
    If currContract Is Nothing Then
        ErrorLabel.caption = "No contract for this from date"
        Exit Sub
    End If
End If

If UseCustomTimesOption Then
    Set currContract = editContractSessionTimes(currContract, sessionStartTime, sessionEndTime)
End If

ReDim lTickfileSpecifiers(1000) As TickfileSPecifier

Dim currSessionStart As Date
Dim thisSessionStart As Date
Dim thisSessionEnd As Date

'currSessionStart = skipWeekends(fromSessionStart, sessionStartTime, sessionEndTime)
currSessionStart = fromSessionStart
j = 0
If CompleteSessionCheck.value = vbChecked Then
    Do While currSessionStart < endTime
        If j > UBound(lTickfileSpecifiers) Then
            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TickfileSPecifier
        End If
        Set lTickfileSpecifiers(j) = New TickfileSPecifier
        lTickfileSpecifiers(j).Contract = currContract
        lTickfileSpecifiers(j).TickfileFormatID = TickfileFormatID
        
        If UseCustomTimesOption Then
            lSession.SessionTimes currSessionStart, thisSessionStart, thisSessionEnd
            lTickfileSpecifiers(j).FromDate = thisSessionStart
            lTickfileSpecifiers(j).ToDate = thisSessionEnd
            lTickfileSpecifiers(j).EntireSession = False
            lTickfileSpecifiers(j).FileName = FormatDateTime(lTickfileSpecifiers(j).FromDate, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(lTickfileSpecifiers(j).ToDate, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.specifier.ToString, vbCrLf, "; ")
        Else
            lTickfileSpecifiers(j).FromDate = currSessionStart
            lTickfileSpecifiers(j).EntireSession = True
            lTickfileSpecifiers(j).FileName = "Session " & _
                                            FormatDateTime(DateValue(currSessionStart), vbShortDate) & _
                                            " " & _
                                            Replace(currContract.specifier.ToString, vbCrLf, "; ")
            
        End If
        
'        currSessionStart = skipWeekends(currSessionStart + 1, _
'                                        sessionStartTime, _
'                                        sessionEndTime)
        currSessionStart = currSessionStart + 1
        
        If mSecType = SecurityTypes.SecTypeFuture Or _
            mSecType = SecurityTypes.SecTypeOption Or _
            mSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.expiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To mContracts.count
                    Set aContract = mContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        Set currContract = aContract
                        If UseCustomTimesOption Then
                            Set currContract = editContractSessionTimes(currContract, _
                                                                        sessionStartTime, _
                                                                        sessionEndTime)
                        End If
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    ErrorLabel.caption = "No contract from " & currSessionStart
                    Exit Sub
                End If
            End If
        End If
        
        j = j + 1
    Loop
    If j = 0 Then
        ErrorLabel.caption = "No trading sessions in specified date range"
        Exit Sub
    End If
    ReDim Preserve lTickfileSpecifiers(j - 1) As TickfileSPecifier
Else
    Set lTickfileSpecifiers(0) = New TickfileSPecifier
    lTickfileSpecifiers(0).Contract = currContract
    lTickfileSpecifiers(0).TickfileFormatID = TickfileFormatID

    lTickfileSpecifiers(0).FromDate = CDate(FromText)
    currSessionStart = currSessionStart + 1
    Do While currSessionStart < endTime
        
        If mSecType = SecurityTypes.SecTypeFuture Or _
            mSecType = SecurityTypes.SecTypeOption Or _
            mSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.expiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To mContracts.count
                    Set aContract = mContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        lTickfileSpecifiers(j).ToDate = currSessionStart
                        lTickfileSpecifiers(j).FileName = FormatDateTime(lTickfileSpecifiers(j).FromDate, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(lTickfileSpecifiers(j).ToDate, vbGeneralDate) & " " & _
                                                    Replace(currContract.specifier.ToString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        j = j + 1
                        If j > UBound(lTickfileSpecifiers) Then
                            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TickfileSPecifier
                        End If
                        
                        Set lTickfileSpecifiers(j) = New TickfileSPecifier
                        lTickfileSpecifiers(j).Contract = currContract
                        lTickfileSpecifiers(j).TickfileFormatID = TickfileFormatID
                    
                        lTickfileSpecifiers(j).FromDate = currSessionStart
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    ErrorLabel.caption = "No contract from " & currSessionStart
                    Exit Sub
                End If
            End If
        End If
        
        currSessionStart = currSessionStart + 1
        
    Loop
        
    lTickfileSpecifiers(j).ToDate = endTime
    lTickfileSpecifiers(j).FileName = FormatDateTime(lTickfileSpecifiers(j).FromDate, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(lTickfileSpecifiers(j).ToDate, vbGeneralDate) & " " & _
                                Replace(currContract.specifier.ToString, vbCrLf, "; ")

    ReDim Preserve lTickfileSpecifiers(j) As TickfileSPecifier
End If

RaiseEvent TickfilesSpecified(lTickfileSpecifiers)

Unload Me

End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let SupportedTickfileFormats( _
                            ByRef value() As TickfileFormatSpecifier)
Dim i As Long

mSupportedTickfileFormats = value

For i = 0 To UBound(mSupportedTickfileFormats)
    FormatCombo.addItem mSupportedTickfileFormats(i).name
Next

FormatCombo.ListIndex = 0
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub adjustCustomTimeFieldAttributes()
If UseCustomTimesOption Then
    enableCustomTimeFields
Else
    disableCustomTimeFields
End If
End Sub

Private Sub checkOk()

OkButton.Enabled = False

If Not ContractSpecBuilder1.ready Then Exit Sub

If Not IsDate(FromText.Text) Then Exit Sub
If CompleteSessionCheck.value = vbUnchecked And Not IsDate(ToText.Text) Then Exit Sub
If CompleteSessionCheck.value = vbChecked And _
    ToText.Text <> "" And _
    Not IsDate(ToText.Text) Then Exit Sub
If IsDate(ToText.Text) Then
    If CDate(FromText.Text) > CDate(ToText.Text) Then Exit Sub
End If

If UseCustomTimesOption Then
    If Not IsDate(CustomFromTimeText) Then Exit Sub
    If Not IsDate(CustomToTimeText) Then Exit Sub
    If CDbl(CDate(CustomFromTimeText)) >= 1# Then Exit Sub
    If CDbl(CDate(CustomToTimeText)) >= 1# Then Exit Sub
End If

OkButton.Enabled = True

End Sub

Private Sub disableCustomTimeFields()
CustomFromTimeText.Enabled = False
CustomFromTimeText.backColor = vbButtonFace
CustomToTimeText.Enabled = False
CustomToTimeText.backColor = vbButtonFace
End Sub

Private Function editContractSessionTimes( _
                ByVal pContract As Contract, _
                ByVal sessionStartTime As Date, _
                ByVal sessionEndTime As Date) As Contract
Dim lContractBuilder As ContractBuilder

Set lContractBuilder = CreateContractBuilder(pContract.specifier)
lContractBuilder.buildFrom pContract
lContractBuilder.sessionEndTime = sessionEndTime
lContractBuilder.sessionStartTime = sessionStartTime
Set editContractSessionTimes = lContractBuilder.Contract
End Function

Private Sub enableCustomTimeFields()
CustomFromTimeText.Enabled = True
CustomFromTimeText.backColor = vbWindowBackground
CustomToTimeText.Enabled = True
CustomToTimeText.backColor = vbWindowBackground
End Sub

'Private Function skipWeekends( _
'                ByVal timestamp As Date, _
'                ByVal sessionStartTime As Date, _
'                ByVal sessionEndTime As Date) As Date
'Dim tradesOvernight As Boolean
'Dim dayOfWeek As Integer
'
'If sessionEndTime < sessionStartTime Then
'    tradesOvernight = True
'End If
'
'dayOfWeek = DatePart("w", timestamp, vbMonday)
'Do While (tradesOvernight And (dayOfWeek = 5 Or dayOfWeek = 6)) Or _
'    (Not tradesOvernight And (dayOfWeek = 6 Or dayOfWeek = 7))
'    timestamp = timestamp + 1
'    dayOfWeek = DatePart("w", timestamp, vbMonday)
'Loop
'skipWeekends = timestamp
'End Function



