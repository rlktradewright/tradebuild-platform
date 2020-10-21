VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
' DebugWindowListener
'

Implements ITestListener

Private Const TO_FILE As Boolean = False

Private mAssertCount As Long
Private mTime As Currency

Private mFile As Long


Private Sub WriteLine(Optional ByVal s As String)
    If TO_FILE Then
        Print #mFile, s
    Else
        LogMessage s
    End If
End Sub

Private Sub ITestListener_Hear(ByVal Message As String, Args() As Variant)

End Sub

Private Sub ITestListener_RunCompleted(ByVal Results As ITestResult)

    WriteLine String$(100, "-")
    WriteLine "Tests Run    : " & Results.TestsRun
    WriteLine "Failures     : " & Results.FailureCount
    WriteLine "Total Asserts: " & Results.AssertCount
    WriteLine "Time         : " & Results.Time & " ms"

    If TO_FILE Then
        Close mFile
    End If
End Sub

Private Sub ITestListener_RunStarted(ByVal Test As ITest)
    If TO_FILE Then
        mFile = FreeFile
        Open "results.txt" For Output As mFile
    End If
    
    mAssertCount = 0
    WriteLine "Starting Tests"
    WriteLine
    WriteLine "Test Name" & Space$(28) & "Message"
    WriteLine String$(100, "-")
End Sub

Private Sub ITestListener_SetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_SetupStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TeardownStarted(ByVal Test As SimplyVBUnitLib.TestMethod, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestCaseCompleted(ByVal Result As TestCaseResult)

End Sub

Private Sub ITestListener_TestCaseStarted(ByVal TestCase As TestCase)

End Sub

Private Sub ITestListener_TestCompleted(ByVal Result As SimplyVBUnitLib.TestResult)
    If Result.IsSuccess Then
        WriteLine "---> Succeeded: "
    ElseIf Result.IsFailure Then
        WriteLine "---> Failed   : "
        WriteLine "     Cause    : " & Result.ToString
        WriteLine
    ElseIf Result.IsError Then
        WriteLine "---> Error    : "
        WriteLine "     Cause    : " & Result.ToString
        WriteLine "     Source   : " & Replace$(Result.Source, vbCrLf, vbCrLf & String$(16, " "))
        WriteLine
    End If
End Sub

Private Sub ITestListener_TestFixtureSetupCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureSetupStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownCompleted(ByVal Result As SimplyVBUnitLib.ContextResult, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestFixtureTeardownStarted(ByVal Test As ITest, ByVal HasContext As Boolean)

End Sub

Private Sub ITestListener_TestStarted(ByVal Test As TestMethod)
    WriteLine "Test Started  : " & Test.FullName
End Sub

Private Sub ITestListener_TestSuiteCompleted(ByVal Result As SimplyVBUnitLib.TestSuiteResult)
    mAssertCount = mAssertCount + Result.AssertCount
    mTime = mTime + Result.Time
End Sub

Private Sub ITestListener_TestSuiteStarted(ByVal Suite As TestSuite)

End Sub

