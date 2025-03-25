Attribute VB_Name = "AA_Startup"
Option Compare Database
Option Explicit
'=================================================================
'run from AUTOEXEC macro
'
'andrew.d.bentley@gmail.com
'=================================================================

Public Function StartUpProcedures()
On Error GoTo Err_StartUpProcedures
    
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Do you want to open TribeVibes?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "TribeVibes StartUp"
    Help = "DEMO.HLP"
    Ctxt = 1000
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    
    If Response = vbYes Then    ' User chose Yes.
        DebugOP ("==START==StartUpProcedures")
        Call StartUpProcedures1
        Call StartUpProcedures2
       
        DebugOP ("==DONE==StartUpProcedures")
    End If
        
Exit_StartUpProcedures:
    Exit Function
Err_StartUpProcedures:
    MsgBox Err.NUMBER & Err.Description
    Resume Exit_StartUpProcedures
    
End Function

Public Function CurrentDirectory()
    CurrentDirectory = CurDir$
End Function


Public Sub StartUpProcedures1()
On Error GoTo Err_StartUpProcedures1
    
    DoCmd.OpenForm "DebugOutput"
    
    Randomize
    DebugOP "Randomize"
    
    DebugOP ("Start-DoCmd.OpenForm 'TRIBEVIBES'")
    DoCmd.OpenForm "TRIBEVIBES"
    DoCmd.Maximize
    
    DebugOP ("Start-PointDirectory")
    PointDirectory
    
    If FormIsOpen("TRIBEVIBES") Then
        Forms.TRIBEVIBES.Refresh

    End If
    
    DebugOP ("Start-UpdateLinkTables")
    UpdateLinkTables


Exit_StartUpProcedures1:
    Exit Sub
Err_StartUpProcedures1:
    MsgBox Err.NUMBER & Err.Description
    Resume Exit_StartUpProcedures1
    
End Sub

Public Sub StartUpProcedures2()
On Error GoTo Err_StartUpProcedures2
    DoCmd.OpenForm "DebugOutput"
    
    DebugOP ("Start-diceroll")
    diceroll
    
    DebugOP ("Start-FIX-TABLE")
    FIX_TABLE
    
    DebugOP ("Start-Fix_General_Info")
    Fix_General_Info
    
    DebugOP ("Start-Fix_Goods")
    Fix_Goods
    
    DebugOP ("Start-FIX_VALID_GOODS")
    FIX_VALID_GOODS
    
    DebugOP ("Start-Tribe_Checking")
    Call Tribe_Checking("Update_All", "", "", "")
    
    DebugOP ("Start-POPULATE_CAPACITIES")
    POPULATE_CAPACITIES
    
    DebugOP ("Start-POPULATE_WEIGHTS")
    POPULATE_WEIGHTS
    
    DebugOP ("Start-Importing_Clan_Spreadsheets")
    Importing_Clan_Spreadsheets
    
    DebugOP ("Start-Importing_City_Spreadsheets")
    Importing_City_Spreadsheets
    
    UpdateUnitMorale
    
    

Exit_StartUpProcedures2:
    Exit Sub
Err_StartUpProcedures2:
    MsgBox Err.NUMBER & Err.Description
    Resume Exit_StartUpProcedures2
    
End Sub


Public Sub UpdateUnitMorale()
'=================================================================
'Updates the unit morale in table TRIBES_General_info to match parent tribe
'
'andrew.d.bentley@gmail.com
'=================================================================

On Error GoTo Err_UpdateUnitMorale

    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "UPDATE_UnitMorale", dbFailOnError
    DebugOP "Morale in " & db.RecordsAffected & " units updated."
    
Exit_UpdateUnitMorale:
    Exit Sub
Err_UpdateUnitMorale:
    MsgBox Err.NUMBER & Err.Description
    Resume Exit_UpdateUnitMorale
End Sub


