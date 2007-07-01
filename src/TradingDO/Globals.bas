Attribute VB_Name = "Globals"
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

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

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "Globals"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function gGenerateErrorMessage( _
                ByVal pConnection As ADODB.Connection) As String
Dim Err As ADODB.Error
Dim errMsg As String

For Each Err In pConnection.Errors
    errMsg = "--------------------" & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
            "    Source: " & Err.Source & vbCrLf & _
            "    SQL state: " & Err.SQLState & vbCrLf & _
            "    Native error: " & Err.NativeError & vbCrLf
Next
pConnection.Errors.Clear
gGenerateErrorMessage = errMsg
End Function

Public Function gCategoryFromString(ByVal value As String) As InstrumentCategories
Select Case UCase$(value)
Case "STOCK", "STK"
    gCategoryFromString = InstrumentCategoryStock
Case "FUTURE", "FUT"
    gCategoryFromString = InstrumentCategoryFuture
Case "OPTION", "OPT"
    gCategoryFromString = InstrumentCategoryOption
Case "FUTURES OPTION", "FOP"
    gCategoryFromString = InstrumentCategoryFuturesOption
Case "CASH"
    gCategoryFromString = InstrumentCategoryCash
'Case "BAG"
'    gCategoryFromString = InstrumentCategoryBag
Case "INDEX", "IND"
    gCategoryFromString = InstrumentCategoryIndex
End Select
End Function

Public Function gCategoryToString(ByVal value As InstrumentCategories) As String
Select Case value
Case InstrumentCategoryStock
    gCategoryToString = "STK"
Case InstrumentCategoryFuture
    gCategoryToString = "FUT"
Case InstrumentCategoryOption
    gCategoryToString = "OPT"
Case InstrumentCategoryFuturesOption
    gCategoryToString = "FOP"
Case InstrumentCategoryCash
    gCategoryToString = "CASH"
'Case InstrumentCategoryBag
'    gCategoryToString = "BAG"
Case InstrumentCategoryIndex
    gCategoryToString = "IND"
End Select
End Function

Public Function gDatabaseTypeFromString( _
                ByVal value As String) As DatabaseTypes
Select Case UCase$(value)
Case "SQLSERVER7", "SQL SERVER 7"
    gDatabaseTypeFromString = DbSQLServer7
Case "SQLSERVER2000", "SQL SERVER 2000"
    gDatabaseTypeFromString = DbSQLServer2000
Case "SQLSERVER2005", "SQL SERVER 2005"
    gDatabaseTypeFromString = DbSQLServer2005
Case "MYSQL5", "MYSQL 5", "MYSQL"
    gDatabaseTypeFromString = DbMySQL5
End Select
End Function

Public Function gDatabaseTypeToString( _
                ByVal value As DatabaseTypes) As String
Select Case value
Case DbSQLServer7
    gDatabaseTypeToString = "SQL Server 7"
Case DbSQLServer2000
    gDatabaseTypeToString = "SQL Server 2000"
Case DbSQLServer2005
    gDatabaseTypeToString = "SQL Server 2005"
Case DbMySQL5
    gDatabaseTypeToString = "MySQL 5"
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================


