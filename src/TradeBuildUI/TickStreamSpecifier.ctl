VERSION 5.00
Begin VB.UserControl TickStreamSpecifier 
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   ScaleHeight     =   3720
   ScaleWidth      =   6720
   Begin VB.Frame Frame2 
      Caption         =   "Contract specification"
      Height          =   3255
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   2535
         TabIndex        =   22
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
      Height          =   2535
      Left            =   2880
      TabIndex        =   10
      Top             =   720
      Width           =   3735
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2175
         ScaleWidth      =   3495
         TabIndex        =   11
         Top             =   240
         Width           =   3495
         Begin VB.TextBox ToDateText 
            Height          =   285
            Left            =   2160
            TabIndex        =   3
            Top             =   0
            Width           =   1260
         End
         Begin VB.TextBox FromDateText 
            Height          =   285
            Left            =   480
            TabIndex        =   2
            Top             =   0
            Width           =   1260
         End
         Begin VB.CheckBox CompleteSessionCheck 
            Caption         =   "Complete sessions"
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   360
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.CheckBox UseExchangeTimezoneCheck 
            Caption         =   "Use exchange timezone (otherwise local time)"
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.Frame SessionTimesFrame 
            Caption         =   "Session times"
            Height          =   1095
            Left            =   0
            TabIndex        =   12
            Top             =   1080
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
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "From"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Label ErrorLabel 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   3360
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
Event ready()
Event TickStreamsSpecified(ByRef pTickfileSpecifier() As TickfileSpecifier)

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

Private mSupportedTickStreamFormats() As TickfileFormatSpecifier
Private WithEvents mContracts       As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mSecType                    As SecurityTypes

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
getSupportedTickstreamFormats
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CompleteSessionCheck_Click()
If CompleteSessionCheck = vbChecked Then
    UseContractTimesOption.Enabled = True
    UseCustomTimesOption.Enabled = True
Else
    UseContractTimesOption.Enabled = False
    UseCustomTimesOption.Enabled = False
End If
adjustCustomTimeFieldAttributes
checkReady
End Sub

Private Sub ContractSpecBuilder1_NotReady()
RaiseEvent NotReady
End Sub

Private Sub ContractSpecBuilder1_Ready()
checkReady
End Sub

Private Sub CustomFromTimeText_Change()
checkReady
End Sub

Private Sub CustomToTimeText_Change()
checkReady
End Sub

Private Sub FromDateText_Change()
checkReady
End Sub

Private Sub ToDateText_Change()
checkReady
End Sub

Private Sub UseContractTimesOption_Click()
adjustCustomTimeFieldAttributes
checkReady
End Sub

Private Sub UseCustomTimesOption_Click()
adjustCustomTimeFieldAttributes
checkReady
End Sub

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
Screen.MousePointer = vbDefault
ErrorLabel.caption = "Invalid contract specification:" & reason
End Sub

Private Sub mContracts_NoMoreContractDetails()
Dim lTickfileSpecifiers() As TickfileSpecifier
Dim i As Long
Dim j As Long
Dim k As Long
Dim lContract As Contract
Dim lSessionBuilder As SessionBuilder
Dim lSession As session
Dim currContract As Contract
Dim customSessionStartTime As Date
Dim customSessionEndTime As Date
Dim fromSessionStart As Date
Dim fromSessionEnd As Date
Dim toSessionStart As Date
Dim toSessionEnd As Date
Dim TickfileFormatID As String

' the from and to dates converted (if necessary) to the contract's timezone
Dim fromDate As Date
Dim toDate As Date

' the from and to dates, session-oriented if required, converted to UTC
Dim fromDateUTC As Date
Dim toDateUTC As Date

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
    
For k = 0 To UBound(mSupportedTickStreamFormats)
    If mSupportedTickStreamFormats(k).name = FormatCombo.Text Then
        TickfileFormatID = mSupportedTickStreamFormats(k).FormalID
        Exit For
    End If
Next

' get the most recent contract (though they should all have the same
' info regarding session times)
Set lContract = mContracts(mContracts.count)

If UseExchangeTimezoneCheck = vbChecked Then
    fromDate = CDate(FromDateText)
    If ToDateText <> "" Then toDate = CDate(ToDateText)
Else
    fromDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(CDate(FromDateText)), lContract.TimeZone)
    If ToDateText <> "" Then toDate = ConvertDateUTCToTZ(ConvertDateLocalToUTC(CDate(ToDateText)), lContract.TimeZone)
End If

' determine start and end times ---------------------------------------------------

Set lSessionBuilder = New SessionBuilder
Set lSession = lSessionBuilder.session

' note that the custom times are in the contract's timezone
If CustomFromTimeText <> "" Then customSessionStartTime = CDate(CustomFromTimeText)
If CustomToTimeText <> "" Then customSessionEndTime = CDate(CustomToTimeText)

If UseCustomTimesOption Then
    lSessionBuilder.sessionStartTime = customSessionStartTime
    lSessionBuilder.sessionEndTime = customSessionEndTime
Else
    lSessionBuilder.sessionStartTime = lContract.sessionStartTime
    lSessionBuilder.sessionEndTime = lContract.sessionEndTime
End If

' set the session start and end times for the starting date (in the contract's
' local timezone)...
lSession.SessionTimes fromDate, _
                    fromSessionStart, _
                    fromSessionEnd
' ... and convert to UTC
fromSessionStart = ConvertDateTZToUTC(fromSessionStart, lContract.TimeZone)
fromSessionEnd = ConvertDateTZToUTC(fromSessionEnd, lContract.TimeZone)
fromDateUTC = fromSessionStart

If toDate <> 0 Then
    ' set the session start and end times for the starting date (in the contract's
    ' local timezone)...
    lSession.SessionTimes toDate, toSessionStart, toSessionEnd
    ' ... and convert to UTC
    toSessionStart = ConvertDateTZToUTC(toSessionStart, lContract.TimeZone)
    toSessionEnd = ConvertDateTZToUTC(toSessionEnd, lContract.TimeZone)
    
    If CompleteSessionCheck.value = vbChecked Then
        toDateUTC = toSessionEnd
    Else
        toDateUTC = ConvertDateTZToUTC(toDate, lContract.TimeZone)
    End If
Else
    toSessionStart = fromSessionStart
    toSessionEnd = fromSessionEnd
    toDateUTC = toSessionEnd
End If

' find contract for start date ------------------------------------------------------
Dim aContract As Contract

If mSecType <> SecurityTypes.SecTypeFuture And _
    mSecType <> SecurityTypes.SecTypeOption And _
    mSecType <> SecurityTypes.SecTypeFuturesOption _
Then
    Set currContract = mContracts(1)
Else
    For i = 1 To mContracts.count
        Set aContract = mContracts(i)
        If DateValue(fromDateUTC) <= _
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
    Set currContract = editContractSessionTimes(currContract, customSessionStartTime, customSessionEndTime)
End If

ReDim lTickfileSpecifiers(1000) As TickfileSpecifier

Dim currSessionStartUTC As Date
Dim thisSessionStart As Date
Dim thisSessionEnd As Date

currSessionStartUTC = fromDateUTC
j = 0
If CompleteSessionCheck.value = vbChecked Then
    Do While currSessionStartUTC < toDateUTC
        If j > UBound(lTickfileSpecifiers) Then
            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TickfileSpecifier
        End If
        Set lTickfileSpecifiers(j) = New TickfileSpecifier
        lTickfileSpecifiers(j).Contract = currContract
        lTickfileSpecifiers(j).TickfileFormatID = TickfileFormatID
        
        If UseCustomTimesOption Then
            lSession.SessionTimes ConvertDateUTCToTZ(currSessionStartUTC, currContract.TimeZone), _
                                    thisSessionStart, _
                                    thisSessionEnd
            lTickfileSpecifiers(j).fromDate = ConvertDateTZToUTC(thisSessionStart, currContract.TimeZone)
            lTickfileSpecifiers(j).toDate = ConvertDateTZToUTC(thisSessionEnd, currContract.TimeZone)
            lTickfileSpecifiers(j).EntireSession = False
            lTickfileSpecifiers(j).FileName = FormatDateTime(lTickfileSpecifiers(j).fromDate, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(lTickfileSpecifiers(j).toDate, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.specifier.ToString, vbCrLf, "; ")
        Else
            lTickfileSpecifiers(j).fromDate = currSessionStartUTC
            lTickfileSpecifiers(j).EntireSession = True
            lTickfileSpecifiers(j).FileName = "Session " & _
                                            FormatDateTime(DateValue(currSessionStartUTC), vbShortDate) & _
                                            " " & _
                                            Replace(currContract.specifier.ToString, vbCrLf, "; ")
            
        End If
        
        currSessionStartUTC = currSessionStartUTC + 1
        
        If mSecType = SecurityTypes.SecTypeFuture Or _
            mSecType = SecurityTypes.SecTypeOption Or _
            mSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStartUTC) > _
                (currContract.expiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To mContracts.count
                    Set aContract = mContracts(i)
                    If DateValue(currSessionStartUTC) <= _
                        (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        Set currContract = aContract
                        If UseCustomTimesOption Then
                            Set currContract = editContractSessionTimes(currContract, _
                                                                        customSessionStartTime, _
                                                                        customSessionEndTime)
                        End If
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    ErrorLabel.caption = "No contract from " & currSessionStartUTC
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
    ReDim Preserve lTickfileSpecifiers(j - 1) As TickfileSpecifier
Else
    Set lTickfileSpecifiers(0) = New TickfileSpecifier
    lTickfileSpecifiers(0).Contract = currContract
    lTickfileSpecifiers(0).TickfileFormatID = TickfileFormatID

    lTickfileSpecifiers(0).fromDate = fromDateUTC
    currSessionStartUTC = currSessionStartUTC + 1
    Do While currSessionStartUTC < toDateUTC
        
        If mSecType = SecurityTypes.SecTypeFuture Or _
            mSecType = SecurityTypes.SecTypeOption Or _
            mSecType = SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStartUTC) > _
                (currContract.expiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To mContracts.count
                    Set aContract = mContracts(i)
                    If DateValue(currSessionStartUTC) <= _
                        (aContract.expiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        lTickfileSpecifiers(j).toDate = currSessionStartUTC
                        lTickfileSpecifiers(j).FileName = FormatDateTime(lTickfileSpecifiers(j).fromDate, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(lTickfileSpecifiers(j).toDate, vbGeneralDate) & " " & _
                                                    Replace(currContract.specifier.ToString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        j = j + 1
                        If j > UBound(lTickfileSpecifiers) Then
                            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TickfileSpecifier
                        End If
                        
                        Set lTickfileSpecifiers(j) = New TickfileSpecifier
                        lTickfileSpecifiers(j).Contract = currContract
                        lTickfileSpecifiers(j).TickfileFormatID = TickfileFormatID
                    
                        lTickfileSpecifiers(j).fromDate = currSessionStartUTC
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    ErrorLabel.caption = "No contract from " & currSessionStartUTC
                    Exit Sub
                End If
            End If
        End If
        
        currSessionStartUTC = currSessionStartUTC + 1
        
    Loop
        
    lTickfileSpecifiers(j).toDate = toDateUTC
    lTickfileSpecifiers(j).FileName = FormatDateTime(lTickfileSpecifiers(j).fromDate, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(lTickfileSpecifiers(j).toDate, vbGeneralDate) & " " & _
                                Replace(currContract.specifier.ToString, vbCrLf, "; ")

    ReDim Preserve lTickfileSpecifiers(j) As TickfileSpecifier
End If

RaiseEvent TickStreamsSpecified(lTickfileSpecifiers)

End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub load()
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

Private Function checkOk() As Boolean

If FormatCombo.ListCount = 0 Then Exit Function

If Not ContractSpecBuilder1.ready Then Exit Function

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

End Function

Private Sub checkReady()
If checkOk Then
    RaiseEvent ready
Else
    RaiseEvent NotReady
End If
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

Set lContractBuilder = CreateContractBuilderFromContract(pContract)
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





