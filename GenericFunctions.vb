Attribute VB_Name = "GenericFunctions"
Option Compare Database
Option Explicit
'======================================================
' General, non-TN-specific functions
'
' andrew.d.bentley@gmail.com
'======================================================

Public Function GetFileNameFromPath(sFilepath)
    'requires MS Scripting RunTime
    Dim fso As New FileSystemObject
    GetFileNameFromPath = fso.GetFileName(sFilepath)
End Function

'---------------------------------------------------------------------------------------
' Procedure : CloseAllOpenTables
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Close all the currently open Tables in the database
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: None required
'
' Usage:
' ~~~~~~
' ? CloseAllOpenTables
'   Returns -> True     => Closed all Tables successfully
'              False    => A problem occurred
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2015-02-17              Initial Release
' 2         2023-02-09              Update variable naming, Error handler, copyright
'---------------------------------------------------------------------------------------
Function CloseAllOpenTables() As Boolean
    On Error GoTo Error_Handler
    Dim oTbls                 As Object
    Dim oTbl                  As Access.AccessObject
    DebugOP "CloseAllOpenTables()"
    Set oTbls = CurrentData.AllTables

    For Each oTbl In oTbls    'Loop all the tables
        If oTbl.IsLoaded = True Then 'check if it is open
            DoCmd.Close acTable, oTbl.Name, acSaveNo
            DebugOP oTbl.Name
        End If
    Next oTbl
    
    CloseAllOpenTables = True

Error_Handler_Exit:
    On Error Resume Next
    Set oTbl = Nothing
    Set oTbls = Nothing
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: CloseAllOpenTables" & vbCrLf & _
           "Error Number: " & Err.NUMBER & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function




