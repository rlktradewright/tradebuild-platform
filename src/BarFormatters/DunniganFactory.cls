VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DunniganFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements BarFormatterFactory

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "DunniganFactory"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' BarFormatterFactory Interface Members
'@================================================================================

Private Property Let BarFormatterFactory_ConfigurationSection( _
                ByVal RHS As TWUtilities30.ConfigurationSection)
'Nothing we need to record in the config
End Property

Private Function BarFormatterFactory_CreateBarFormatter( _
                ByVal pBars As TimeframeUtils26.bars) As ChartUtils26.BarFormatter
Dim bf As New DunniganBarFormatter
bf.bars = pBars
Set BarFormatterFactory_CreateBarFormatter = bf
End Function

Private Sub BarFormatterFactory_LoadFromConfig( _
                ByVal config As TWUtilities30.ConfigurationSection)
'Nothing to load from the config
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


