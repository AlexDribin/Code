Attribute VB_Name = "StripTablesMod"
Option Compare Database
Option Explicit
'=================================================================
'1) ensure all tables are relinked
'2) strip tables of Clan data not in approved list
'
'andrew.d.bentley@gmail.com
'=================================================================

Public Sub StripTablesMacro()
    Dim output As New Form_DebugOutput
    output.Visible = True
    
    Call UpdateLinkTables
    'Debug.Print "=== Links Updated ==="
    DebugOP ("=== Links Updated ===")
    Call StripTables
    'Debug.Print "=== Tables Stripped ==="
    DebugOP ("=== Tables stripped ===")
    
End Sub

Public Function PointDirectory()
'=================================================================
'points CurDir$ to the path that this database is in.
'
'andrew.d.bentley@gmail.com
'=================================================================
On Error GoTo Err_PointDirectory

    ChDir CurrentProject.Path
    
Exit_PointDirectory:
     Exit Function
Err_PointDirectory:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_PointDirectory
End Function


Public Sub UpdateLinkTables()
'=================================================================
'Ensure that tables are relinked to the back end in the current folder
'andrew.d.bentley@gmail.com
'=================================================================
On Error GoTo Err_UpdateLinkTables

    Dim strPathTVbe As String
    Dim objDatabase As Object
    Dim tblDef As TableDef
    '
    strPathTVbe = CurrentProject.Path & _
            "\tvdatapr.accdb"
    '
    'Set the object to current database
    Set objDatabase = CurrentDb
    '
    'Go through each table in the database and check if it is linked table
    'if yes, then update new database path
    For Each tblDef In objDatabase.TableDefs
        If tblDef.SourceTableName <> "" Then
            'If it is hidden table then make it a visible table
            tblDef.Properties("Attributes").Value = 0
            'Change the database path to new path
            'if database requires password then you can un-comment password section in below code
            tblDef.Connect = ";DATABASE=" & strPathTVbe '& "; PWD=1234"
            'Refresh the table
            tblDef.RefreshLink
            DebugOP "REFRESHING LINK: " & tblDef.Name & " - " & tblDef.Connect
        End If
    Next
    '
    'Close the objects
    Set tblDef = Nothing
    Set objDatabase = Nothing

Exit_UpdateLinkTables:
     Exit Sub
Err_UpdateLinkTables:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_UpdateLinkTables
End Sub

Public Sub StripTables()
'=================================================================
'Will strip sensitive tables of Tribe data leaving a core data set
'andrew.d.bentley@gmail.com
'=================================================================
On Error GoTo Err_StripTables
    
    Dim rs As DAO.Recordset
    Dim sRSQL As String
    sRSQL = "SELECT * FROM StripTables " & _
            "WHERE AT_Active = True " & _
            "ORDER BY AT_TableName;"
    Dim sSQL As String
    Set rs = CurrentDb.OpenRecordset(sRSQL)
    
    'Check to see if the recordset actually contains rows
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
            'Perform an edit
            Debug.Print rs!AT_TableName & " - " & rs!AT_UnitField
            
'            sSQL = "DELETE * FROM " & _
'                    rs!AT_TableName & _
'                    " WHERE IsClanInApprovedList(" & _
'                    rs!AT_UnitField & _
'                    ") = FALSE;"
            sSQL = "DELETE * FROM " & _
                    rs!AT_TableName & _
                    " AS T WHERE Not Exists " & _
                    "(SELECT * FROM StripClans WHERE StripClans.StrippingClanNo = " & _
                    "GetClanFromUnit(T." & _
                    rs!AT_UnitField & _
                    "));"
            
            
            CurrentDb.Execute sSQL, dbFailOnError
    
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
    
    MsgBox "Finished Stripping Tables."
    
    rs.Close 'Close the recordset
    Set rs = Nothing 'Clean up

Exit_StripTables:
     Exit Sub
Err_StripTables:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_StripTables
End Sub

Public Function GetClanFromUnit(vUnit As Variant) As String
'=================================================================
'Will extract the Clan no from a unit or retain the Clan if in ### form
'andrew.d.bentley@gmail.com
'=================================================================
On Error GoTo Err_GetClanFromUnit

    If Len(Nz(vUnit, "")) = 0 Then
        GetClanFromUnit = ""
    ElseIf Len(Nz(vUnit, "")) = 3 Then
        GetClanFromUnit = vUnit
    Else
        GetClanFromUnit = Mid(vUnit, 2, 3)
    End If

Exit_GetClanFromUnit:
     Exit Function
Err_GetClanFromUnit:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_GetClanFromUnit
End Function

Public Sub DebugLinkedTables()
'=================================================================
'Lists tables and their connections
'andrew.d.bentley@gmail.com
'=================================================================
On Error GoTo Err_DebugLinkedTables

    Dim strPath As String
    Dim objDatabase As Object
    Dim tblDef As TableDef
    '
    strPath = CurrentProject.Path & _
            "\tvdatapr.accdb"
    '
    'Set the object to current database
    Set objDatabase = CurrentDb
    '
    'Go through each table in the database and check if it is linked table
    'if yes, then update new database path
    For Each tblDef In objDatabase.TableDefs
        If tblDef.SourceTableName <> "" Then
            'If it is hidden table then make it a visible table
            tblDef.Properties("Attributes").Value = 0
            'Change the database path to new path
            'if database requires password then you can un-comment password section in below code
'            tblDef.Connect = ";DATABASE=" & strPath '& "; PWD=1234"
'            'Refresh the table
'            tblDef.RefreshLink
        Debug.Print tblDef.Name
        Debug.Print tblDef.Connect
        
        End If
    Next
    '
    'Close the objects
    Set tblDef = Nothing
    Set objDatabase = Nothing
    
Exit_DebugLinkedTables:
     Exit Sub
Err_DebugLinkedTables:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_DebugLinkedTables
End Sub

Public Sub BreakLinksToExternalTables()
On Error GoTo Err_BreakLinksToExternalTables
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim index As Long
    Dim C As Long
    Set dbs = CurrentDb
    
    C = dbs.TableDefs.count
    
    For index = (C - 1) To 0 Step -1
        Set tdf = dbs.TableDefs(index)
        'Debug.Print tdf.Connect & " - " & tdf.Name
        If Left(tdf.Connect, 10) = ";DATABASE=" Then
        DoCmd.DeleteObject acTable, tdf.Name
        DebugOP "DELETING LINK TO:" & tdf.Connect & " - " & tdf.Name
        End If
    Next index
    
Exit_BreakLinksToExternalTables:
     Exit Sub
Err_BreakLinksToExternalTables:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_BreakLinksToExternalTables
End Sub


Sub CreateLinksToAllTablesInExternalDB(sExtDbPath As String)
On Error GoTo Err_CreateLinksToAllTablesInExternalDB
    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    
    Set db = OpenDatabase(sExtDbPath)
    
    For Each tdf In db.TableDefs 'Loop through all the table in the external database
        If Left(tdf.Name, 4) <> "MSys" Then 'Exclude System Tables
            On Error Resume Next
            Access.DoCmd.TransferDatabase acLink, "Microsoft Access", sExtDbPath, _
                                          acTable, tdf.Name, tdf.Name
            'DebugOP "Linked - " & tdf.Name
            DebugOP "ESTABLISHING LINK TO: " & tdf.Name
        End If
    Next tdf
    db.Close
   
    Set db = Nothing

Exit_CreateLinksToAllTablesInExternalDB:
     Exit Sub
Err_CreateLinksToAllTablesInExternalDB:
     MsgBox Err.NUMBER & Err.Description
     Resume Exit_CreateLinksToAllTablesInExternalDB
End Sub


