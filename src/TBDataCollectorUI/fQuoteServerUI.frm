VERSION 5.00
Begin VB.Form fDataCollectorUI 
   Caption         =   "TradeBuild Quote Server"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox LogText 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   1200
      Width           =   5055
   End
   Begin VB.TextBox ConnectionStatusText 
      BackColor       =   &H8000000F&
      Height          =   255
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   15
      Left            =   4080
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   15
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   14
      Left            =   4080
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   14
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   13
      Left            =   4080
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   13
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   12
      Left            =   4080
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   12
      Left            =   3360
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   11
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   11
      Left            =   2280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   10
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   10
      Left            =   2280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   9
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   9
      Left            =   2280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   8
      Left            =   3000
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   8
      Left            =   2280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   7
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   7
      Left            =   1200
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   6
      Left            =   1200
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   5
      Left            =   1200
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   4
      Left            =   1200
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   840
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   840
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox DataLightText 
      BackColor       =   &H8000000F&
      Height          =   255
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox ShortNameText 
      Height          =   255
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "Stop"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Conn"
      Height          =   255
      Left            =   4440
      TabIndex        =   34
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "fDataCollectorUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements DataSignalListener
Implements WriterListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const MaxTickerListenerIndex As Long = 15

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mDataCollector As TBDataCollector
Attribute mDataCollector.VB_VarHelpID = -1
Private mStop As Boolean
Private mFailpoint As Long

Private mTickerListenerIndex As Long

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

'================================================================================
' DataSignalListener Members
'================================================================================

Private Sub DataSignalListener_signalOff(ev As DataSignalEvent)
Dim listener As TickerListener
Set listener = ev.Source
If listener.Index = -1 Then Exit Sub
DataLightText(listener.Index).BackColor = vbButtonFace
End Sub

Private Sub DataSignalListener_signalOn(ev As DataSignalEvent)
Dim listener As TickerListener
Set listener = ev.Source
If listener.Index = -1 Then Exit Sub
DataLightText(listener.Index).BackColor = vbGreen
ConnectionStatusText.BackColor = vbGreen
End Sub

'================================================================================
' WriterListener Interface Members
'================================================================================

Private Sub WriterListener_notify(ev As WriterEvent)
Dim tf As Timeframe
Dim tk As ticker

Select Case ev.Action
Case WriterNotifications.WriterNotReady
    If TypeOf ev.Source Is Timeframe Then
        Set tf = ev.Source
        logMessage tf.barLength & _
                        "-" & TimePeriodUnitsToString(tf.barUnit) & _
                        " bar writer not ready for " & _
                        tf.contract.specifier.localSymbol
    Else
        Set tk = ev.Source
        logMessage "Tickfile writer not ready for " & _
                        tk.contract.specifier.localSymbol
    End If
Case WriterNotifications.WriterReady
    If TypeOf ev.Source Is Timeframe Then
        Set tf = ev.Source
        logMessage tf.barLength & _
                        "-" & TimePeriodUnitsToString(tf.barUnit) & _
                        " bar writer ready for " & _
                        tf.contract.specifier.localSymbol
    Else
        Set tk = ev.Source
        logMessage "Tickfile writer ready for " & _
                        tk.contract.specifier.localSymbol
    End If
Case WriterNotifications.WriterFileCreated
    If TypeOf ev.Source Is Timeframe Then
        Set tf = ev.Source
        logMessage "Writing " & tf.barLength & _
                    "-" & TimePeriodUnitsToString(tf.barUnit) & _
                    " bars for " & _
                    tf.contract.specifier.localSymbol & _
                    " to " & ev.FileName
    Else
        Set tk = ev.Source
        logMessage "Writing tickdata for " & _
                    tk.contract.specifier.localSymbol & _
                    " to " & ev.FileName
    End If
End Select
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub StopButton_Click()
mStop = True
mDataCollector.stopCollection
StopButton.Enabled = False
ConnectionStatusText.BackColor = vbButtonFace
logMessage "Data collection stopped by user"
End Sub

'================================================================================
' mDataCollector Event Handlers
'================================================================================

Private Sub mDataCollector_connected()
ConnectionStatusText.BackColor = vbGreen
logMessage "Connected ok to realtime data source"
StopButton.Enabled = True
End Sub

Private Sub mDataCollector_connectFailed(ByVal description As String)
ConnectionStatusText.BackColor = vbRed
logMessage "Connect failed: " & description
StopButton.Enabled = False
End Sub

Private Sub mDataCollector_ConnectionClosed()
ConnectionStatusText.BackColor = vbRed
logMessage "Connection to realtime data source closed"
StopButton.Enabled = False
End Sub

Private Sub mDataCollector_Reconnecting()
logMessage "Reconnecting to realtime data source"
StopButton.Enabled = True
End Sub

Private Sub mDataCollector_connectionToTWSClosed( _
                ByVal reconnecting As Boolean)
ConnectionStatusText.BackColor = vbRed
logMessage "Connection to TWS closed"
If reconnecting Then
    logMessage "Attempting to reconnect"
Else
    clearTickers
    StopButton.Enabled = False
End If
End Sub

Private Sub mDataCollector_errorMessage( _
                ByVal errorCode As ApiNotifyCodes, _
                ByVal errorMsg As String)

logMessage "Error " & errorCode & ": " & errorMsg
End Sub

Private Sub mDataCollector_Info(ev As InfoEvent)
logMessage ev.Data
End Sub

Private Sub mDataCollector_NotifyMessage( _
                ByVal eventCode As TradeBuild25.ApiNotifyCodes, _
                ByVal eventMsg As String)
logMessage "Notification " & eventCode & ": " & eventMsg
End Sub

Private Sub mDataCollector_ServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal serviceProviderName As String, _
                ByVal message As String)
logMessage "Service provider error (" & serviceProviderName & "): " & errorCode & ": " & message
End Sub

Private Sub mDataCollector_TickerAdded(ByVal ticker As ticker)
ticker.addTickfileWriterListener Me
End Sub

Private Sub mDataCollector_TickerListenerAdded( _
                ByVal listener As TickerListener)
If mTickerListenerIndex > MaxTickerListenerIndex Then
    listener.Index = -1
    logMessage "Can't display ticker for " & listener.contract.specifier.localSymbol
    Exit Sub
End If

listener.Index = mTickerListenerIndex
mTickerListenerIndex = mTickerListenerIndex + 1
listener.addDataSignalListener Me
ShortNameText(listener.Index) = listener.contract.specifier.localSymbol

End Sub

Private Sub mDataCollector_TimeframeAdded(ByVal tf As Timeframe)
tf.addBarWriterListener Me
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let dataCollector(ByVal value As TBDataCollector)
Set mDataCollector = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearTickers()
Dim i As Long

If mTickerListenerIndex = 0 Then Exit Sub
    
For i = 0 To mTickerListenerIndex - 1
    ShortNameText(i).text = ""
    DataLightText(i).BackColor = vbButtonFace
Next
mTickerListenerIndex = 0
End Sub

Private Sub logMessage(ByVal text As String)
If Len(LogText.text) > (32000 - Len(text)) Then
    LogText.text = Right$(LogText.text, 32000 - Len(text)) & vbCrLf & Format(Now, "hh:mm:ss") & "  " & text
Else
    LogText.text = LogText.text & vbCrLf & Format(Now, "hh:mm:ss") & "  " & text
End If
LogText.SelStart = Len(LogText.text)
End Sub


