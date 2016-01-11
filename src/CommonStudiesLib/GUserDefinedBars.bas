Attribute VB_Name = "GUserDefinedBars"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "GUserDefinedBars"

Public Const UserDefinedBarsInputOpenInterest As String = "Open interest"
Public Const UserDefinedBarsInputOpenInterestUcase As String = "OPEN INTEREST"

Public Const UserDefinedBarsInputValue  As String = "Value"
Public Const UserDefinedBarsInputValueUCase  As String = "VALUE"

Public Const UserDefinedBarsInputTotalVolume As String = "Total volume"
Public Const UserDefinedBarsInputTotalVolumeUcase As String = "TOTAL VOLUME"

Public Const UserDefinedBarsInputTickVolume As String = "Tick volume"
Public Const UserDefinedBarsInputTickVolumeUcase As String = "TICK VOLUME"

Public Const UserDefinedBarsInputBarNumber  As String = "Bar number"
Public Const UserDefinedBarsInputBarNumberUCase  As String = "BAR NUMBER"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================


Private mStudyDefinition As StudyDefinition

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

'@================================================================================
' Procedures
'@================================================================================


Public Property Let defaultParameters(ByVal Value As Parameters)
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Assert False, "Study has no parameters", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get defaultParameters() As Parameters
Const ProcName As String = "defaultParameters"
On Error GoTo Err

Set defaultParameters = New Parameters

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyDefinition() As StudyDefinition
Const ProcName As String = "StudyDefinition"
On Error GoTo Err

If mStudyDefinition Is Nothing Then
    Set mStudyDefinition = gCreateBarStudyDefinition( _
                                UserDefinedBarsStudyName, _
                                UserDefinedBarsStudyShortName, _
                                "User-defined bars " & _
                                "divide value movement into periods (bars) of duration " & _
                                "determined by the program that supplies the values. " & _
                                "For each period the open, high, low and close values " & _
                                "are determined.", _
                                UserDefinedBarsInputValue, _
                                UserDefinedBarsInputTotalVolume, _
                                UserDefinedBarsInputTickVolume, _
                                UserDefinedBarsInputOpenInterest, _
                                UserDefinedBarsInputBarNumber)
    
End If

Set StudyDefinition = mStudyDefinition.Clone

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Helper Function
'@================================================================================














