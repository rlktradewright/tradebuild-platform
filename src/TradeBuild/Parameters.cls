VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IParameters

'================================================================================
' Events
'================================================================================

Event Dirty()

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mParameters As Parameters2.parameters
Attribute mParameters.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mParameters = New Parameters2.parameters
End Sub

'================================================================================
' IParameters Interface Members
'================================================================================

Private Sub IParameters_addParameter( _
                ByVal param As TradeBuildSP.IParameter)
addParameter param
End Sub

Private Function IParameters_Clone() As TradeBuildSP.IParameters
Set IParameters_Clone = clone
End Function

Private Property Get IParameters_Count() As Long
IParameters_Count = Count
End Property

Private Function IParameters_Equals( _
                params As TradeBuildSP.IParameters) As Boolean
IParameters_Equals = Equals(params)
End Function

Private Function IParameters_getFirstParameter() As TradeBuildSP.IParameter
Set IParameters_getFirstParameter = getFirstParameter
End Function

Private Function IParameters_getNextParameter() As TradeBuildSP.IParameter
Set IParameters_getNextParameter = getNextParameter
End Function

Private Function IParameters_getParameterValue( _
                ByVal name As String, _
                Optional ByVal defaultValue As String) As String
IParameters_getParameterValue = getParameterValue(name, defaultValue)
End Function

Private Property Get IParameters_isDirty() As Boolean
IParameters_isDirty = isDirty
End Property

Private Function IParameters_setParameterValue( _
                ByVal name As String, _
                ByVal value As String) As TradeBuildSP.IParameter
Set IParameters_setParameterValue = setParameterValue(name, value)
End Function

Private Sub IParameters_sortByName()
sortByName
End Sub

'================================================================================
' mParameters Event Handlers
'================================================================================

Private Sub mParameters_Dirty()
RaiseEvent Dirty
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mParameters.Count
End Property

Public Property Get datasource() As IUnknown
Set datasource = mParameters
End Property

Public Property Get isDirty() As Boolean
isDirty = mParameters.isDirty
End Property

Friend Property Let parameters(ByVal value As Parameters2.parameters)
Set mParameters = value
End Property

Friend Property Get parameters() As Parameters2.parameters
Set parameters = mParameters
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addParameter(ByVal param As TradeBuild.parameter)
mParameters.addParameter param.parameter
End Sub

Public Function clone() As TradeBuild.parameters
Set clone = New TradeBuild.parameters
clone.parameters = Me.parameters.clone
End Function

Public Function Equals(params As TradeBuild.parameters) As Boolean
If params Is Nothing Then Exit Function
Equals = mParameters.Equals(params.parameters)
End Function

Public Function getFirstParameter() As TradeBuild.parameter
Dim param As Parameters2.parameter
Set param = mParameters.getFirstParameter

If Not param Is Nothing Then
    Set getFirstParameter = New TradeBuild.parameter
    getFirstParameter.parameter = param
End If
End Function


Public Function getNextParameter() As TradeBuild.parameter
Dim param As Parameters2.parameter
Set param = mParameters.getNextParameter

If Not param Is Nothing Then
    Set getNextParameter = New TradeBuild.parameter
    getNextParameter.parameter = param
End If
End Function

Public Function getParameterValue( _
                ByVal name As String, _
                Optional ByVal defaultValue As String) As String
getParameterValue = mParameters.getParameterValue(name, defaultValue)
End Function

Public Function setParameterValue(ByVal name As String, ByVal value As String) As TradeBuild.parameter
Set setParameterValue = New TradeBuild.parameter
setParameterValue.parameter = mParameters.setParameterValue(name, value)
End Function

Public Sub sortByName()
mParameters.sortByName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mParameters.NewEnum
End Function

'================================================================================
' Helper Functions
'================================================================================
