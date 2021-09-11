VERSION 5.00
Object = "{CABB46DA-3D1D-40AA-A327-E1B9FA2B5DB5}#5.0#0"; "SimplyVBUnitUI.ocx"
Begin VB.Form fMain 
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8685
   Begin SimplyVBUnitUI.SimplyVBUnitCtl SimplyVBUnitCtl1 
      Height          =   6135
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10821
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' This program uses the excellent SimplyVBUnit VB6 unit testing framework by
' Kelly Ethridge (see https://github.com/kellyethridge/SimplyVBUnit).



' Namespaces Available:
'       Assert.*            ie. Assert.AreEqual Expected, Actual
'
' Public Functions Availabe:
'       AddTest <TestObject>
'       AddListener <ITestListener Object>
'       AddFilter <ITestFilter Object>
'       RemoveFilter <ITestFilter Object>
'       WriteLine "Message"
'
' Adding a testcase:
'   Use AddTest <object>
'
' Steps to create a TestCase:
'
'   1. Add a new class
'   2. Name it as desired
'   3. (Optionally) Add a Setup/Teardown method to be run before and after every test.
'   4. (Optionally) Add a TestFixtureSetup/TestFixtureTeardown method to be run at the
'      before the first test and after the last test.
'   5. Add public Subs of the tests you want run. No parameters.
'
'      Public Sub MyTest()
'          Assert.AreEqual a, b
'      End Sub
'

Implements ILogListener

Private mUnhandledExceptionHandler As UnhandledExcptnHndlr
Private mLogFormatter As ILogFormatter


Private Sub Form_Load()
    Set mUnhandledExceptionHandler = New UnhandledExcptnHndlr
    
    ApplicationGroupName = "TradeWright"
    ApplicationName = "TradeBuild Unit Tester"
    SetupDefaultLogging Command
    
    GetLogger("").AddLogListener Me
    Set mLogFormatter = CreateBasicLogFormatter()
        
    AddListener New SimpleListener

    ' Add tests here
    '
    AddTest New TestTickfileManager
    AddTest New TestTickfileReader
    AddTest New TestTickfileReaderDB
    AddTest New TestTickfileRdrDbAsync
    AddTest New TestTickfileWriter
    
    AddTest New TestBarUtils
    AddTest New TestSessionUtils
    
    AddTest New TestTickfileListGenerator
    
    AddTest New TestOrderUtils
        
    
End Sub

Private Sub Form_Initialize()
    SimplyVBUnitCtl1.Init App.EXEName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Dim ar() As String
ar = Split(mLogFormatter.FormatRecord(Logrec), vbCrLf)
Dim i As Long
For i = 0 To UBound(ar)
    WriteLine ar(i)
Next
End Sub
