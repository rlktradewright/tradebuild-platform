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
      TabIndex        =   35
      Top             =   2160
      Width           =   3735
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   120
         ScaleHeight     =   570
         ScaleWidth      =   3495
         TabIndex        =   36
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton UseContractTimesOption 
            Caption         =   "Use contract times"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.TextBox CustomToTimeText 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   14
            Top             =   240
            Width           =   660
         End
         Begin VB.TextBox CustomFromTimeText 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   13
            Top             =   0
            Width           =   660
         End
         Begin VB.OptionButton UseCustomTimesOption 
            Caption         =   "Use custom times"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "To"
            Height          =   255
            Left            =   2040
            TabIndex        =   38
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "From"
            Height          =   255
            Left            =   2040
            TabIndex        =   37
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
      TabIndex        =   15
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dates/Times"
      Height          =   975
      Left            =   3000
      TabIndex        =   29
      Top             =   1200
      Width           =   3735
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   3495
         TabIndex        =   30
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox CompleteSessionCheck 
            Height          =   255
            Left            =   1560
            TabIndex        =   11
            Top             =   360
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox FromText 
            Height          =   285
            Left            =   480
            TabIndex        =   9
            Top             =   0
            Width           =   1260
         End
         Begin VB.TextBox ToText 
            Height          =   285
            Left            =   2160
            TabIndex        =   10
            Top             =   0
            Width           =   1260
         End
         Begin VB.Label Label9 
            Caption         =   "Complete session(s)"
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "From"
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "To"
            Height          =   255
            Left            =   1800
            TabIndex        =   31
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data source"
      Height          =   1095
      Left            =   3000
      TabIndex        =   26
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   3495
         TabIndex        =   27
         Top             =   240
         Width           =   3495
         Begin VB.ComboBox FormatCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0000
            Left            =   720
            List            =   "fTickfileSpecifier.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Format"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   0
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contract specification"
      Height          =   3255
      Left            =   120
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   240
         Width           =   2535
         Begin VB.TextBox ShortNameText 
            Height          =   285
            Left            =   1200
            TabIndex        =   0
            Top             =   0
            Width           =   1335
         End
         Begin VB.ComboBox RightCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0004
            Left            =   1200
            List            =   "fTickfileSpecifier.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2520
            Width           =   855
         End
         Begin VB.ComboBox TypeCombo 
            Height          =   315
            ItemData        =   "fTickfileSpecifier.frx":0008
            Left            =   1200
            List            =   "fTickfileSpecifier.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox SymbolText 
            Height          =   285
            Left            =   1200
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox ExpiryText 
            Height          =   285
            Left            =   1200
            TabIndex        =   3
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox ExchangeText 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox StrikePriceText 
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox CurrencyText 
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Short name"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Right"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Strike price"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Symbol"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Type"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Expiry"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Exchange"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "Currency"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   855
         End
      End
   End
   Begin VB.Label ErrorLabel 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   34
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

Event TickfilesSpecified(ByRef pTickfileSpecifier() As TradeBuild26.TickfileSpecifier)

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

Private mTradeBuildAPIRef           As WeakReference

Private mSupportedTickfileFormats() As TradeBuild26.TickfileFormatSpecifier
Private WithEvents mContracts As TradeBuild26.Contracts
Attribute mContracts.VB_VarHelpID = -1

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub Form_Load()
gAddItemToCombo TypeCombo, gSecTypeToString(SecurityTypes.SecTypeStock), SecurityTypes.SecTypeStock
gAddItemToCombo TypeCombo, gSecTypeToString(SecurityTypes.SecTypeFuture), SecurityTypes.SecTypeFuture
gAddItemToCombo TypeCombo, gSecTypeToString(SecurityTypes.SecTypeOption), SecurityTypes.SecTypeOption
gAddItemToCombo TypeCombo, gSecTypeToString(SecurityTypes.SecTypeFuturesOption), SecurityTypes.SecTypeFuturesOption
gAddItemToCombo TypeCombo, gSecTypeToString(SecurityTypes.SecTypeCash), SecurityTypes.SecTypeCash
gAddItemToCombo TypeCombo, gSecTypeToString(SecurityTypes.SecTypeIndex), SecurityTypes.SecTypeIndex

gAddItemToCombo RightCombo, gOptionRightToString(OptionRights.OptCall), OptionRights.OptCall
gAddItemToCombo RightCombo, gOptionRightToString(OptionRights.OptPut), OptionRights.OptPut

End Sub

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
    UseContractTimesOption.enabled = True
    UseCustomTimesOption.enabled = True
Else
    UseContractTimesOption.enabled = False
    UseCustomTimesOption.enabled = False
End If
adjustCustomTimeFieldAttributes
checkOk
End Sub

Private Sub CurrencyText_Change()
checkOk
End Sub

Private Sub CustomFromTimeText_Change()
checkOk
End Sub

Private Sub CustomToTimeText_Change()
checkOk
End Sub

Private Sub ExchangeText_Change()
checkOk
End Sub

Private Sub ExpiryText_Change()
checkOk
End Sub

Private Sub FromText_Change()
checkOk
End Sub

Private Sub OkButton_Click()
Dim contractSpec As TradeBuild26.contractSpecifier

ErrorLabel.Caption = ""

Screen.MousePointer = vbHourglass

Set contractSpec = New TradeBuild26.contractSpecifier
With contractSpec
    .symbol = SymbolText.Text
    .localSymbol = ShortNameText.Text
    .sectype = getSecType
    .expiry = IIf(.sectype = SecurityTypes.SecTypeFuture Or _
                    .sectype = SecurityTypes.SecTypeFuturesOption Or _
                    .sectype = SecurityTypes.SecTypeOption, _
                    ExpiryText.Text, _
                    "")
    .exchange = ExchangeText.Text
    .currencyCode = CurrencyText.Text
    If .sectype = SecurityTypes.SecTypeFuturesOption Or _
        .sectype = SecurityTypes.SecTypeOption _
    Then
        .strike = IIf(StrikePriceText.Text = "", 0, StrikePriceText.Text)
        If RightCombo.Text <> "" Then
            .right = RightCombo.itemData(RightCombo.ListIndex)
        End If
    End If
End With

Set mContracts = tb.NewContracts(contractSpec)
mContracts.Load

End Sub

Private Sub RightCombo_Click()
checkOk
End Sub

Private Sub shortnametext_Change()
checkOk
End Sub

Private Sub StrikePriceText_Change()
checkOk
End Sub

Private Sub SymbolText_Change()
checkOk
End Sub

Private Sub ToText_Change()
checkOk
End Sub

Private Sub TypeCombo_Click()

Select Case getSecType
Case SecurityTypes.SecTypeNone
    ExpiryText.enabled = True
    StrikePriceText.enabled = True
    RightCombo.enabled = True
Case SecurityTypes.SecTypeFuture
    ExpiryText.enabled = True
    StrikePriceText.enabled = False
    RightCombo.enabled = False
Case SecurityTypes.SecTypeStock
    ExpiryText.enabled = False
    StrikePriceText.enabled = False
    RightCombo.enabled = False
Case SecurityTypes.SecTypeOption
    ExpiryText.enabled = True
    StrikePriceText.enabled = True
    RightCombo.enabled = True
Case SecurityTypes.SecTypeFuturesOption
    ExpiryText.enabled = True
    StrikePriceText.enabled = True
    RightCombo.enabled = True
Case SecurityTypes.SecTypeCash
    ExpiryText.enabled = False
    StrikePriceText.enabled = False
    RightCombo.enabled = False
Case SecurityTypes.SecTypeIndex
    ExpiryText.enabled = False
    StrikePriceText.enabled = False
    RightCombo.enabled = False
Case SecurityTypes.SecTypeBag
    ExpiryText.enabled = False
    StrikePriceText.enabled = False
    RightCombo.enabled = False
End Select

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
ErrorLabel.Caption = "Invalid contract specification:" & reason
End Sub

Private Sub mContracts_NoMoreContractDetails()
Dim lTickfileSpecifiers() As TradeBuild26.TickfileSpecifier
Dim i As Long
Dim j As Long
Dim k As Long
Dim lContract As TradeBuild26.Contract
Dim lSessionBuilder As SessionBuilder
Dim lSession As session
Dim currContract As TradeBuild26.Contract
Dim sectype As TradeBuild26.SecurityTypes
Dim sessionStartTime As Date
Dim sessionEndTime As Date
Dim fromSessionStart As Date
Dim fromSessionEnd As Date
Dim toSessionStart As Date
Dim toSessionEnd As Date
Dim endTime As Date
Dim tickfileFormatID As String

Screen.MousePointer = vbDefault
If mContracts.Count = 0 Then
    ErrorLabel.Caption = "No contracts meet this specification"
    Exit Sub
End If

sectype = getSecType
If sectype <> TradeBuild26.SecurityTypes.SecTypeFuture And _
    sectype <> TradeBuild26.SecurityTypes.SecTypeOption And _
    sectype <> TradeBuild26.SecurityTypes.SecTypeFuturesOption _
Then
    If mContracts.Count > 1 Then
        ' don't see how this can happen, but just in case!
        ErrorLabel.Caption = "More than one contract meets this specification"
        Exit Sub
    End If
End If
    
For k = 0 To UBound(mSupportedTickfileFormats)
    If mSupportedTickfileFormats(k).name = FormatCombo.Text Then
        tickfileFormatID = mSupportedTickfileFormats(k).FormalID
        Exit For
    End If
Next

' get the most recent contract (though they should all have the same
'info regarding session times)
Set lContract = mContracts(mContracts.Count)

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
Dim aContract As TradeBuild26.Contract

If sectype <> TradeBuild26.SecurityTypes.SecTypeFuture And _
    sectype <> TradeBuild26.SecurityTypes.SecTypeOption And _
    sectype <> TradeBuild26.SecurityTypes.SecTypeFuturesOption _
Then
    Set currContract = mContracts(1)
Else
    For i = 1 To mContracts.Count
        Set aContract = mContracts(i)
        If DateValue(fromSessionStart) <= _
            (aContract.ExpiryDate - aContract.daysBeforeExpiryToSwitch) _
        Then
            Set currContract = aContract
            Exit For
        End If
    Next
    
    If currContract Is Nothing Then
        ErrorLabel.Caption = "No contract for this from date"
        Exit Sub
    End If
End If

If UseCustomTimesOption Then
    currContract.sessionStartTime = sessionStartTime
    currContract.sessionEndTime = sessionEndTime
End If

ReDim lTickfileSpecifiers(1000) As TradeBuild26.TickfileSpecifier

Dim currSessionStart As Date
Dim thisSessionStart As Date
Dim thisSessionEnd As Date

'currSessionStart = skipWeekends(fromSessionStart, sessionStartTime, sessionEndTime)
currSessionStart = fromSessionStart
j = 0
If CompleteSessionCheck.value = vbChecked Then
    Do While currSessionStart < endTime
        If j > UBound(lTickfileSpecifiers) Then
            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TradeBuild26.TickfileSpecifier
        End If
        Set lTickfileSpecifiers(j).Contract = currContract
        lTickfileSpecifiers(j).tickfileFormatID = tickfileFormatID
        
        If UseCustomTimesOption Then
            lSession.SessionTimes currSessionStart, thisSessionStart, thisSessionEnd
            lTickfileSpecifiers(j).From = thisSessionStart
            lTickfileSpecifiers(j).To = thisSessionEnd
            lTickfileSpecifiers(j).EntireSession = False
            lTickfileSpecifiers(j).filename = FormatDateTime(lTickfileSpecifiers(j).From, vbGeneralDate) & _
                                        "-" & _
                                        FormatDateTime(lTickfileSpecifiers(j).To, vbGeneralDate) & _
                                        " " & _
                                        Replace(currContract.specifier.ToString, vbCrLf, "; ")
        Else
            lTickfileSpecifiers(j).From = currSessionStart
            lTickfileSpecifiers(j).EntireSession = True
            lTickfileSpecifiers(j).filename = "Session " & _
                                            FormatDateTime(DateValue(currSessionStart), vbShortDate) & _
                                            " " & _
                                            Replace(currContract.specifier.ToString, vbCrLf, "; ")
            
        End If
        
'        currSessionStart = skipWeekends(currSessionStart + 1, _
'                                        sessionStartTime, _
'                                        sessionEndTime)
        currSessionStart = currSessionStart + 1
        
        If sectype = TradeBuild26.SecurityTypes.SecTypeFuture Or _
            sectype = TradeBuild26.SecurityTypes.SecTypeOption Or _
            sectype = TradeBuild26.SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.ExpiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To mContracts.Count
                    Set aContract = mContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        Set currContract = aContract
                        If UseCustomTimesOption Then
                            currContract.sessionStartTime = sessionStartTime
                            currContract.sessionEndTime = sessionEndTime
                        End If
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    ErrorLabel.Caption = "No contract from " & currSessionStart
                    Exit Sub
                End If
            End If
        End If
        
        j = j + 1
    Loop
    If j = 0 Then
        ErrorLabel.Caption = "No trading sessions in specified date range"
        Exit Sub
    End If
    ReDim Preserve lTickfileSpecifiers(j - 1) As TradeBuild26.TickfileSpecifier
Else
    Set lTickfileSpecifiers(0).Contract = currContract
    lTickfileSpecifiers(j).tickfileFormatID = tickfileFormatID

    lTickfileSpecifiers(0).From = CDate(FromText)
    currSessionStart = currSessionStart + 1
    Do While currSessionStart < endTime
        
        If sectype = TradeBuild26.SecurityTypes.SecTypeFuture Or _
            sectype = TradeBuild26.SecurityTypes.SecTypeOption Or _
            sectype = TradeBuild26.SecurityTypes.SecTypeFuturesOption _
        Then
            If DateValue(currSessionStart) > _
                (currContract.ExpiryDate - currContract.daysBeforeExpiryToSwitch) _
            Then
                For i = i + 1 To mContracts.Count
                    Set aContract = mContracts(i)
                    If DateValue(currSessionStart) <= _
                        (aContract.ExpiryDate - aContract.daysBeforeExpiryToSwitch) _
                    Then
                        lTickfileSpecifiers(j).To = currSessionStart
                        lTickfileSpecifiers(j).filename = FormatDateTime(lTickfileSpecifiers(j).From, vbGeneralDate) & _
                                                    "-" & _
                                                    FormatDateTime(lTickfileSpecifiers(j).To, vbGeneralDate) & " " & _
                                                    Replace(currContract.specifier.ToString, vbCrLf, "; ")
                        
                        Set currContract = aContract
                        
                        j = j + 1
                        If j > UBound(lTickfileSpecifiers) Then
                            ReDim Preserve lTickfileSpecifiers(UBound(lTickfileSpecifiers) + 1000) As TradeBuild26.TickfileSpecifier
                        End If
                        
                        Set lTickfileSpecifiers(j).Contract = currContract
                        lTickfileSpecifiers(j).tickfileFormatID = tickfileFormatID
                    
                        lTickfileSpecifiers(j).From = currSessionStart
                        Exit For
                    End If
                Next
                If currContract Is Nothing Then
                    ErrorLabel.Caption = "No contract from " & currSessionStart
                    Exit Sub
                End If
            End If
        End If
        
        currSessionStart = currSessionStart + 1
        
    Loop
        
    lTickfileSpecifiers(j).To = endTime
    lTickfileSpecifiers(j).filename = FormatDateTime(lTickfileSpecifiers(j).From, vbGeneralDate) & _
                                "-" & _
                                FormatDateTime(lTickfileSpecifiers(j).To, vbGeneralDate) & " " & _
                                Replace(currContract.specifier.ToString, vbCrLf, "; ")

    ReDim Preserve lTickfileSpecifiers(j) As TradeBuild26.TickfileSpecifier
End If

RaiseEvent TickfilesSpecified(lTickfileSpecifiers)

Unload Me

End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let SupportedTickfileFormats( _
                            ByRef value() As TradeBuild26.TickfileFormatSpecifier)
Dim i As Long

mSupportedTickfileFormats = value

For i = 0 To UBound(mSupportedTickfileFormats)
    FormatCombo.AddItem mSupportedTickfileFormats(i).name
Next

FormatCombo.ListIndex = 0
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = CreateWeakReference(value)
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
Dim sectype As TradeBuild26.SecurityTypes

OkButton.enabled = False

sectype = getSecType

If ShortNameText.Text = "" Then
    If SymbolText = "" Then Exit Sub
    If sectype = SecurityTypes.SecTypeNone Then Exit Sub
    If sectype = SecurityTypes.SecTypeBag Then Exit Sub
    If sectype = SecurityTypes.SecTypeOption Or _
        sectype = SecurityTypes.SecTypeFuturesOption _
    Then
        If StrikePriceText = "" Or RightCombo.Text = "" Then Exit Sub
    End If
End If
    
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

OkButton.enabled = True

End Sub

Private Sub disableCustomTimeFields()
CustomFromTimeText.enabled = False
CustomFromTimeText.BackColor = vbButtonFace
CustomToTimeText.enabled = False
CustomToTimeText.BackColor = vbButtonFace
End Sub

Private Sub enableCustomTimeFields()
CustomFromTimeText.enabled = True
CustomFromTimeText.BackColor = vbWindowBackground
CustomToTimeText.enabled = True
CustomToTimeText.BackColor = vbWindowBackground
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

Private Function getSecType() As TradeBuild26.SecurityTypes
If TypeCombo.ListIndex = -1 Then
    getSecType = SecTypeNone
Else
    getSecType = TypeCombo.itemData(TypeCombo.ListIndex)
End If
End Function

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function


