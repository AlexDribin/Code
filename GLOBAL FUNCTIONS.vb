Attribute VB_Name = "GLOBAL FUNCTIONS"
'*                           VERSION 3.1.4                                         *'
'first call is for Jeff's PC at Home
'Declare Function DROLL Lib "e:\office\access\tribes\tvutil.dll" (ByVal roll_type%, ByVal level%, ByVal dice_sides%, ByVal reset_roll%, ByVal TRIBE%, ByVal PRESET%, ByVal MODIFY%) As Long

'second call is for Jeff's PC at Work
'Declare Function DROLL Lib "c:\Users\jf70\OneDrive - VicGov\Documents\My Documents\Tribes\Office\access\tribes\tvutil.dll" (ByVal roll_type%, ByVal level%, ByVal dice_sides%, ByVal reset_roll%, ByVal TRIBE%, ByVal PRESET%, ByVal MODIFY%) As Long

'first call is for Mum's PC
'Declare Function DROLL Lib "F:\Office\Access\Tribes\tvutil.dll" (ByVal roll_type%, ByVal level%, ByVal dice_sides%, ByVal reset_roll%, ByVal TRIBE%, ByVal PRESET%, ByVal MODIFY%) As Long

'third call is for Peter's PC's
'Declare Function DROLL Lib "c:\office\access\tribes\tvutil.dll" (ByVal roll_type%, ByVal level%, ByVal dice_sides%, ByVal reset_roll%, ByVal TRIBE%, ByVal PRESET%, ByVal MODIFY%) As Long

'Fourth call is for David's laptop
'Declare Function DROLL Lib "C:\Users\Dell\Documents\TNDB\tvutil.dll" (ByVal roll_type%, ByVal level%, ByVal dice_sides%, ByVal reset_roll%, ByVal TRIBE%, ByVal PRESET%, ByVal MODIFY%) As Long

' It is important that the default directory in access is set to the same directory as tribes is in. Otherwise, errors will happen.

Option Compare Database   'Use database order for string comparisons
Option Explicit


Function calc_cost()
Dim CostClan As String
Dim OriginalClan As String
Dim OriginalTribe As String
Dim qdfCurrent As QueryDef
Dim GM_NAME As String
Dim TGroup(10) As String
Dim TValue(10) As Single
Dim TCount As Integer
Dim TCost As Single
Dim TDiscount As Integer

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

If GMTABLE![cALc_costs_PROCESSED] = "Y" Then
   Msg = "Calc Costs Function has already been processed!!!"
   MsgBox (Msg)
   Exit Function
End If

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GM_NAME = GMTABLE![Name]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst

COST_CLAN = "EMPTY"
OriginalClan = TRIBEINFO![CLAN]

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW TRIBES_general_info SET TRIBES_general_info.[AMT RECEIVED] = 0, TRIBES_general_info.COST = 0;")
qdfCurrent.Execute

' populate array
Set COSTSTABLE = TVDBGM.OpenRecordset("GM_Costs_Table")
COSTSTABLE.MoveFirst
TCount = 0
Do While Not COSTSTABLE.EOF
       TGroup(TCount) = COSTSTABLE![Group]
       TValue(TCount) = COSTSTABLE![Cost]
       COSTSTABLE.MoveNext
       TCount = TCount + 1
Loop

TCost = 0

Do Until TRIBEINFO.EOF
   If TRIBEINFO![CLAN] = TRIBEINFO![TRIBE] Then
      CostClan = TRIBEINFO![COST CLAN]
      'loop array to grab general cost
      TCount = 0
      Do
            If TGroup(TCount) = "COST CLAN" Then
               Exit Do
            End If
            TCount = TCount + 1
            If TCount > 10 Then
                Exit Do
            End If
      Loop
      TCost = TCost + TValue(TCount)
      TCount = 0
      Do
            If TGroup(TCount) = TRIBEINFO![Village] Then
                Exit Do
            End If
            TCount = TCount + 1
            If TCount > 10 Then
                Exit Do
            End If
      Loop
      TCost = TCost + TValue(TCount)
   ElseIf TRIBEINFO![CLAN] = TRIBEINFO![CLAN] Then
       TCount = 0
      Do
            If TGroup(TCount) = TRIBEINFO![Village] Then
                Exit Do
            End If
             TCount = TCount + 1
             If TCount > 10 Then
                 Exit Do
             End If
       Loop
       TCost = TCost + TValue(TCount)
   Else
      ' NOT A VALID GROUP
      
 End If

 TRIBEINFO.MoveNext
 If TRIBEINFO.EOF Then
      Exit Do
   End If
 
 If TRIBEINFO![CLAN] = OriginalClan Then
     ' KEEP GOING
 Else
     OriginalClan = TRIBEINFO![CLAN]
     OriginalTribe = TRIBEINFO![TRIBE]
     TRIBEINFO.MoveFirst
     TRIBEINFO.Seek "=", CostClan, CostClan
     ' if discount > 0 then apply discount to TCost
     If IsNull(TRIBEINFO![Discount]) Then
        TDiscount = 0
     Else
        TDiscount = (TCost / 100) * TRIBEINFO![Discount]
     End If
     TCost = TCost - TDiscount
     TRIBEINFO.Edit
     TRIBEINFO![Cost] = TRIBEINFO![Cost] + TCost
     TRIBEINFO.UPDATE
     TRIBEINFO.MoveFirst
     TRIBEINFO.Seek "=", OriginalClan, OriginalTribe
     TCost = 0
     TCount = 0
  End If

Loop

TRIBEINFO.MoveFirst

Do Until TRIBEINFO.EOF
   If TRIBEINFO![Cost] > 0 Then
      TRIBEINFO.Edit
      TRIBEINFO![CREDIT] = TRIBEINFO![CREDIT] - TRIBEINFO![Cost]
      TRIBEINFO.UPDATE
   End If
   TRIBEINFO.MoveNext

   If TRIBEINFO.EOF Then
      Exit Do
   End If

Loop

TRIBEINFO.Close

DoCmd.Hourglass False

End Function

Function CHECK_MORALE(CLAN, TRIBE)

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst
TRIBEINFO.Seek "=", CLAN, TRIBE
TRIBEINFO.Edit

DICE1 = DICE_ROLL(CLAN, TRIBE)
DICE2 = DICE_ROLL(CLAN, TRIBE)
    
If (TRIBEINFO![MORALE] >= 0.9) And (TRIBEINFO![MORALE] <= 1.1) Then
   TRIBEINFO![MORALE] = TRIBEINFO![MORALE] + 0.01
       
ElseIf (TRIBEINFO![MORALE] >= 1.11) And (TRIBEINFO![MORALE] <= 1.15) Then
   If DICE1 <= 75 Then
      TRIBEINFO![MORALE] = TRIBEINFO![MORALE] + 0.01
   End If

ElseIf (TRIBEINFO![MORALE] >= 1.16) And (TRIBEINFO![MORALE] <= 1.2) Then
   If DICE1 <= 50 Then
      TRIBEINFO![MORALE] = TRIBEINFO![MORALE] + 0.01
   End If

ElseIf (TRIBEINFO![MORALE] >= 1.21) Then
   If DICE1 <= 25 Then
      TRIBEINFO![MORALE] = TRIBEINFO![MORALE] + 0.01
   End If

End If
    
TRIBEINFO.UPDATE
TRIBEINFO.Close

End Function

Function DELETE_ATTACHED_TABLES()
On Error GoTo ERR_DEL_ATT_TAB

'======================================
' Commented out as Split database can reattach ALL tables
' andrew.d.bentley@gmail.com
'======================================
'TRIBE_STATUS = "Delete Attached Tables"
'DebugOP "f - DELETE_ATTACHED_TABLES()"
'
'   DoCmd.DeleteObject A_TABLE, "CLAN_STATS"
'   DoCmd.DeleteObject A_TABLE, "COMPLETED_RESEARCH"
'   DoCmd.DeleteObject A_TABLE, "DICE_ROLLS"
'   DoCmd.DeleteObject A_TABLE, "GLOBAL"
'   DoCmd.DeleteObject A_TABLE, "GOODS_STATS"
'   DoCmd.DeleteObject A_TABLE, "GOODS_TRIBES_PROCESSING"
'   DoCmd.DeleteObject A_TABLE, "GAMES_WEATHER"
'   DoCmd.DeleteObject A_TABLE, "gm_costs_table"
'   DoCmd.DeleteObject A_TABLE, "HERD_SWAPS"
'   DoCmd.DeleteObject A_TABLE, "HEX_MAP"
'   DoCmd.DeleteObject A_TABLE, "HEX_MAP_CITY"
'   DoCmd.DeleteObject A_TABLE, "HEX_MAP_CONST"
'   DoCmd.DeleteObject A_TABLE, "HEX_MAP_MINERALS"
'   DoCmd.DeleteObject A_TABLE, "HEX_MAP_POLITICS"
'   DoCmd.DeleteObject A_TABLE, "HEXMAP_FARMING"
'   DoCmd.DeleteObject A_TABLE, "HEXMAP_PERMANENT_FARMING"
'   DoCmd.DeleteObject A_TABLE, "MASSTRANSFERS"
'   DoCmd.DeleteObject A_TABLE, "MODIFIERS"
'   DoCmd.DeleteObject A_TABLE, "PACIFICATION_TABLE"
'   DoCmd.DeleteObject A_TABLE, "PERMANENT_MESSAGES_table"
'   DoCmd.DeleteObject A_TABLE, "POPULATION_INCREASE"
'   DoCmd.DeleteObject A_TABLE, "Process_Scout_Movement"
'   DoCmd.DeleteObject A_TABLE, "Process_Research"
'   DoCmd.DeleteObject A_TABLE, "Process_Skills"
'   DoCmd.DeleteObject A_TABLE, "Process_Tribe_Movement"
'   DoCmd.DeleteObject A_TABLE, "Process_Tribes_Activity"
'   DoCmd.DeleteObject A_TABLE, "Process_Tribes_Activity_Copy"
'   DoCmd.DeleteObject A_TABLE, "Process_Tribes_Item_Allocation"
'   DoCmd.DeleteObject A_TABLE, "Process_Tribes_Item_Allocation_Copy"
'   DoCmd.DeleteObject A_TABLE, "Process_Tribes_Transfers"
'   DoCmd.DeleteObject A_TABLE, "PROVS_AVAILABILITY"
'   DoCmd.DeleteObject A_TABLE, "RESEARCH_ATTEMPTS"
'   DoCmd.DeleteObject A_TABLE, "Scout_Movement"
'   DoCmd.DeleteObject A_TABLE, "SEEKING_RETURNS_TABLE"
'   DoCmd.DeleteObject A_TABLE, "SHIP_DAMAGE"
'   DoCmd.DeleteObject A_TABLE, "SKILL_ATTEMPTS"
'   DoCmd.DeleteObject A_TABLE, "SKILLS"
'   DoCmd.DeleteObject A_TABLE, "SKILLS_STATS"
'   DoCmd.DeleteObject A_TABLE, "SPECIAL_TRANSFER_ROUTES"
'   DoCmd.DeleteObject A_TABLE, "TEMP_TRADING_POST"
'   DoCmd.DeleteObject A_TABLE, "TERRAIN_COMBAT"
'   DoCmd.DeleteObject A_TABLE, "TRADING_POST_GOODS"
'   DoCmd.DeleteObject A_TABLE, "TRIBE_CHECKING"
'   DoCmd.DeleteObject A_TABLE, "TRIBE_RESEARCH"
'   DoCmd.DeleteObject A_TABLE, "TRIBEs_PROCESSING"
'   DoCmd.DeleteObject A_TABLE, "TRIBES_BOOKS"
'   DoCmd.DeleteObject A_TABLE, "TRIBES_general_info"
'   DoCmd.DeleteObject A_TABLE, "TRIBES_GOODS"
'   DoCmd.DeleteObject A_TABLE, "TRIBES_SPECIALISTS"
'   DoCmd.DeleteObject A_TABLE, "TRIBES_TURNS_ACTIVITY"
'   DoCmd.DeleteObject A_TABLE, "TURN_INFO_REQD_NEXT_TURN"
'   DoCmd.DeleteObject A_TABLE, "TURNS_trading_post_activity"
'   DoCmd.DeleteObject A_TABLE, "TURNS_ACTIVITIES"
'   DoCmd.DeleteObject A_TABLE, "UNDER_CONSTRUCTION"
'   DoCmd.DeleteObject A_TABLE, "VALID_GOODS"
'   DoCmd.DeleteObject A_TABLE, "WEAPON_ARMOUR"
'   DoCmd.DeleteObject A_TABLE, "WEATHER"
'   DoCmd.DeleteObject A_TABLE, "WEATHER_COMBAT"

ERR_DEL_ATT_TAB_CLOSE:
   Exit Function

ERR_DEL_ATT_TAB:
   Resume Next

End Function

Function CLEAN_UP_and_RESET()
Dim skilltab As Recordset        ' PROCESS_SKILLS
Dim researchtab As Recordset     ' PROCESS_RESEARCH
Dim QUERY As String
Dim strSQL As String
Dim qdfCurrent As QueryDef

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Forms![GLOBAL INFO]![Status] = "Start Clean Up"
Forms![GLOBAL INFO].Repaint

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM TURNS_ACTIVITIES;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Goods_Tribes_Processing;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Tribes_Processing;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Implement_Usage;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TRIBES_GOODS_USAGE;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Tribe_Activity_Required_By_Later_Activities;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW SKILLS SET SKILLS.ATTEMPTED = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW SKILLS SET SKILLS.SUCCESSFUL = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW TRIBE_RESEARCH SET TRIBE_RESEARCH.[RESEARCH ATTEMPTED] = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW TRIBE_RESEARCH SET TRIBE_RESEARCH.[RESEARCH ATTAINED] = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW TRIBES_general_info SET TRIBES_general_info.COST = 0;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW HEX_MAP_POLITICS SET HEX_MAP_POLITICS.[POP_INCREASED] = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW COMPLETED_RESEARCH SET COMPLETED_RESEARCH.[COMPLETED_THIS_TURN] = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "Delete SKILLS.[SKILL LEVEL] FROM SKILLS WHERE (((SKILLS.[SKILL LEVEL])=0));")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM TRIBE_CHECKING;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Building_Usage;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Clan_Stats;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Goods_Stats;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Dice_Rolls;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Skills_Stats;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM TURNS_Trading_Post_Activity;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM process_tribes_activity_copy;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "INSERT INTO Process_Tribes_Activity_Copy SELECT Process_Tribes_Activity.* FROM Process_Tribes_Activity;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW Process_Tribes_Activity_Copy SET Process_Tribes_Activity_Copy.PROCESSED = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM process_tribes_activity;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM process_tribes_item_allocation_copy;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "INSERT INTO process_tribes_item_allocation_copy SELECT process_tribes_item_allocation.* FROM process_tribes_item_allocation;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW process_tribes_item_allocation_copy SET process_tribes_item_allocation_copy.PROCESSED = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE DISTINCTROW Tribes_Specialists SET Tribes_Specialists.SPECIALISTS_USED = 0;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM process_tribes_item_allocation;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM process_tribes_transfers;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM process_tribe_movement;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM scout_movement_copy;")
qdfCurrent.Execute

strSQL = "Insert Into process_tribe_movement(TRIBE, FOLLOW_TRIBE,MOVEMENT_1) Values('9999','','EMPTY')"
CurrentDb.Execute strSQL

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM MASSTRANSFERS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Movement_Trace;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Process_Skills;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM Process_Research;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Scouting_Results;")
qdfCurrent.Execute

Set SkillsTab = TVDBGM.OpenRecordset("Process_Skills")
SkillsTab.index = "PRIMARYKEY"
SkillsTab.AddNew
SkillsTab![TRIBE] = "9999"
SkillsTab![Order] = "1"
SkillsTab![TOPIC] = "EMPTY"
SkillsTab![PROCESSED] = "Y"
SkillsTab.UPDATE
SkillsTab.Close

Set researchtab = TVDBGM.OpenRecordset("Process_Research")
researchtab.AddNew
researchtab![TRIBE] = "9999"
researchtab![TOPIC] = "EMPTY"
researchtab![PROCESSED] = "Y"
researchtab.UPDATE
researchtab.Close

' delete existing clan data where tribe name = abandoned
' Loop through TRIBES_general_info
Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst

Do While Not TRIBEINFO.EOF

   If TRIBEINFO![TRIBE NAME] = "ABANDONED" Then
      ' TRIBES_GOODS
      QUERY_STRING = "DELETE * FROM TRIBES_GOODS"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRIBES_GOODS.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' TRADING_POST_GOODS
      QUERY_STRING = "DELETE * FROM TRADING_POST_GOODS"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRADING_POST_GOODS.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' SKILLS
      QUERY_STRING = "DELETE * FROM SKILLS"
      QUERY_STRING = QUERY_STRING & " WHERE (((SKILLS.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' COMPLETED_RESEARCH
      QUERY_STRING = "DELETE * FROM COMPLETED_RESEARCH"
      QUERY_STRING = QUERY_STRING & " WHERE (((COMPLETED_RESEARCH.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' HERD_SWAPS
      QUERY_STRING = "DELETE * FROM HERD_SWAPS"
      QUERY_STRING = QUERY_STRING & " WHERE (((HERD_SWAPS.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' HEX_MAP_FARMING
      QUERY_STRING = "DELETE * FROM HEXMAP_FARMING"
      QUERY_STRING = QUERY_STRING & " WHERE (((HEXMAP_FARMING.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' HEXMAP_Permanent_FARMING
      QUERY_STRING = "DELETE * FROM HEXMAP_Permanent_FARMING"
      QUERY_STRING = QUERY_STRING & " WHERE (((HEXMAP_Permanent_FARMING.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' MODIFIERS
      QUERY_STRING = "DELETE * FROM MODIFIERS"
      QUERY_STRING = QUERY_STRING & " WHERE (((MODIFIERS.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' Pacification_Table
      QUERY_STRING = "DELETE * FROM Pacification_Table"
      QUERY_STRING = QUERY_STRING & " WHERE (((Pacification_Table.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' Permanent_Messages_Table
      QUERY_STRING = "DELETE * FROM Permanent_Messages_Table"
      QUERY_STRING = QUERY_STRING & " WHERE (((Permanent_Messages_Table.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' POPULATION_INCREASE
      QUERY_STRING = "DELETE * FROM POPULATION_INCREASE"
      QUERY_STRING = QUERY_STRING & " WHERE (((POPULATION_INCREASE.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' Provs_Availability
      QUERY_STRING = "DELETE * FROM Provs_Availability"
      QUERY_STRING = QUERY_STRING & " WHERE (((Provs_Availability.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' Research_Attempts
      QUERY_STRING = "DELETE * FROM Research_Attempts"
      QUERY_STRING = QUERY_STRING & " WHERE (((Research_Attempts.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' Skill_Attempts
      QUERY_STRING = "DELETE * FROM Skill_Attempts"
      QUERY_STRING = QUERY_STRING & " WHERE (((Skill_Attempts.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' TRIBE_CHECKING
      QUERY_STRING = "DELETE * FROM TRIBE_CHECKING"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRIBE_CHECKING.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' TRIBE_RESEARCH
      QUERY_STRING = "DELETE * FROM TRIBE_RESEARCH"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRIBE_RESEARCH.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' TRIBES_BOOKS
      QUERY_STRING = "DELETE * FROM TRIBES_BOOKS"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRIBES_BOOKS.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
      
      ' TRIBES_SPECIALISTS
      QUERY_STRING = "DELETE * FROM TRIBES_SPECIALISTS"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRIBES_SPECIALISTS.CLAN)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.CLAN '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
    
      ' HEX_MAP_CONST
      QUERY_STRING = "DELETE * FROM HEX_MAP_CONST"
      QUERY_STRING = QUERY_STRING & " WHERE (((HEX_MAP_CONST.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
   
      ' HEX_MAP_POLITICS
      QUERY_STRING = "DELETE * FROM HEX_MAP_POLITICS"
      QUERY_STRING = QUERY_STRING & " WHERE (((HEX_MAP_POLITICS.PL_TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
   
       ' UNDER_CONSTRUCTION
      QUERY_STRING = "DELETE * FROM UNDER_CONSTRUCTION"
      QUERY_STRING = QUERY_STRING & " WHERE (((UNDER_CONSTRUCTION.TRIBE)='"
      QUERY_STRING = QUERY_STRING & " TRIBEINFO.TRIBE '));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
   
        ' clean up
      QUERY_STRING = "DELETE * FROM PROCESS_TRIBES_ACTIVITY"
      QUERY_STRING = QUERY_STRING & " WHERE (((PROCESS_TRIBES_ACTIVITY.TRIBE)='ISNULL'"
      QUERY_STRING = QUERY_STRING & "));"
      Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
  
    End If
   
   TRIBEINFO.MoveNext

   If TRIBEINFO.EOF Then
      Exit Do
   End If
Loop

QUERY_STRING = "DELETE * FROM TRIBES_general_info"
QUERY_STRING = QUERY_STRING & " WHERE (((TRIBES_general_info.[TRIBE NAME])='ABANDONED'));"
Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute

' clean up of abandoned clans complete

QUERY_STRING = "INSERT INTO TRIBE_CHECKING ( CLAN, TRIBE, [CURRENT HEX], WARRIORS, ACTIVES,"
QUERY_STRING = QUERY_STRING & " INACTIVES, SLAVE ) SELECT CLAN, TRIBE, [CURRENT HEX],"
QUERY_STRING = QUERY_STRING & " WARRIORS, ACTIVES, INACTIVES, SLAVE"
QUERY_STRING = QUERY_STRING & " FROM TRIBES_GENERAL_INFO;"
Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute

Call Reset_Implements_and_Goods_Usage_Tables

Forms![GLOBAL INFO]![Status] = "End Clean Up"
Forms![GLOBAL INFO].Repaint

DoCmd.Hourglass False

End Function

Function Determine_Hex_Map_Across_Letter(HEX_MAP)
' ASCII OF A = 65
Determine_Hex_Map_Across_Letter = ((Asc(Mid(HEX_MAP, 2, 1))) - 65) * 30

End Function

Function Determine_Hex_Map_Down_Letter(HEX_MAP)
' ASCII OF A = 65
Determine_Hex_Map_Down_Letter = ((Asc(Mid(HEX_MAP, 1, 1))) - 65) * 21

End Function

Function Determine_New_Hex_Map_Across_Letter(MAP_ACROSS)

   If MAP_ACROSS < 0 Then
      Determine_New_Hex_Map_Across_Letter = "Z"
   ElseIf MAP_ACROSS < 30 Then
      Determine_New_Hex_Map_Across_Letter = "A"
   ElseIf MAP_ACROSS < 60 Then
      Determine_New_Hex_Map_Across_Letter = "B"
   ElseIf MAP_ACROSS < 90 Then
      Determine_New_Hex_Map_Across_Letter = "C"
   ElseIf MAP_ACROSS < 120 Then
      Determine_New_Hex_Map_Across_Letter = "D"
   ElseIf MAP_ACROSS < 150 Then
      Determine_New_Hex_Map_Across_Letter = "E"
   ElseIf MAP_ACROSS < 180 Then
      Determine_New_Hex_Map_Across_Letter = "F"
   ElseIf MAP_ACROSS < 210 Then
      Determine_New_Hex_Map_Across_Letter = "G"
   ElseIf MAP_ACROSS < 240 Then
      Determine_New_Hex_Map_Across_Letter = "H"
   ElseIf MAP_ACROSS < 270 Then
      Determine_New_Hex_Map_Across_Letter = "I"
   ElseIf MAP_ACROSS < 300 Then
      Determine_New_Hex_Map_Across_Letter = "J"
   ElseIf MAP_ACROSS < 330 Then
      Determine_New_Hex_Map_Across_Letter = "K"
   ElseIf MAP_ACROSS < 360 Then
      Determine_New_Hex_Map_Across_Letter = "L"
   ElseIf MAP_ACROSS < 390 Then
      Determine_New_Hex_Map_Across_Letter = "M"
   ElseIf MAP_ACROSS < 420 Then
      Determine_New_Hex_Map_Across_Letter = "N"
   ElseIf MAP_ACROSS < 450 Then
      Determine_New_Hex_Map_Across_Letter = "O"
   ElseIf MAP_ACROSS < 480 Then
      Determine_New_Hex_Map_Across_Letter = "P"
    ElseIf MAP_ACROSS < 510 Then
      Determine_New_Hex_Map_Across_Letter = "Q"
   ElseIf MAP_ACROSS < 540 Then
      Determine_New_Hex_Map_Across_Letter = "R"
   ElseIf MAP_ACROSS < 570 Then
      Determine_New_Hex_Map_Across_Letter = "S"
   ElseIf MAP_ACROSS < 600 Then
      Determine_New_Hex_Map_Across_Letter = "T"
   ElseIf MAP_ACROSS < 630 Then
      Determine_New_Hex_Map_Across_Letter = "U"
   ElseIf MAP_ACROSS < 660 Then
      Determine_New_Hex_Map_Across_Letter = "V"
   ElseIf MAP_ACROSS < 690 Then
      Determine_New_Hex_Map_Across_Letter = "W"
   ElseIf MAP_ACROSS < 720 Then
      Determine_New_Hex_Map_Across_Letter = "X"
   ElseIf MAP_ACROSS < 750 Then
      Determine_New_Hex_Map_Across_Letter = "Y"
   ElseIf MAP_ACROSS < 780 Then
      Determine_New_Hex_Map_Across_Letter = "Z"
  End If


End Function

Function Determine_New_Hex_Map_Down_Letter(MAP_DOWN)
   
   If MAP_DOWN < 0 Then
      Determine_New_Hex_Map_Down_Letter = "Z"
   ElseIf MAP_DOWN < 21 Then
      Determine_New_Hex_Map_Down_Letter = "A"
   ElseIf MAP_DOWN < 42 Then
      Determine_New_Hex_Map_Down_Letter = "B"
   ElseIf MAP_DOWN < 63 Then
      Determine_New_Hex_Map_Down_Letter = "C"
   ElseIf MAP_DOWN < 84 Then
      Determine_New_Hex_Map_Down_Letter = "D"
   ElseIf MAP_DOWN < 105 Then
      Determine_New_Hex_Map_Down_Letter = "E"
   ElseIf MAP_DOWN < 126 Then
      Determine_New_Hex_Map_Down_Letter = "F"
   ElseIf MAP_DOWN < 147 Then
      Determine_New_Hex_Map_Down_Letter = "G"
   ElseIf MAP_DOWN < 168 Then
      Determine_New_Hex_Map_Down_Letter = "H"
   ElseIf MAP_DOWN < 189 Then
      Determine_New_Hex_Map_Down_Letter = "I"
   ElseIf MAP_DOWN < 210 Then
      Determine_New_Hex_Map_Down_Letter = "J"
   ElseIf MAP_DOWN < 231 Then
      Determine_New_Hex_Map_Down_Letter = "K"
   ElseIf MAP_DOWN < 252 Then
      Determine_New_Hex_Map_Down_Letter = "L"
   ElseIf MAP_DOWN < 273 Then
      Determine_New_Hex_Map_Down_Letter = "M"
   ElseIf MAP_DOWN < 294 Then
      Determine_New_Hex_Map_Down_Letter = "N"
   ElseIf MAP_DOWN < 315 Then
      Determine_New_Hex_Map_Down_Letter = "O"
   ElseIf MAP_DOWN < 336 Then
      Determine_New_Hex_Map_Down_Letter = "P"
   ElseIf MAP_DOWN < 357 Then
      Determine_New_Hex_Map_Down_Letter = "Q"
   ElseIf MAP_DOWN < 378 Then
      Determine_New_Hex_Map_Down_Letter = "R"
   ElseIf MAP_DOWN < 399 Then
      Determine_New_Hex_Map_Down_Letter = "S"
   ElseIf MAP_DOWN < 420 Then
      Determine_New_Hex_Map_Down_Letter = "T"
   ElseIf MAP_DOWN < 441 Then
      Determine_New_Hex_Map_Down_Letter = "U"
   ElseIf MAP_DOWN < 462 Then
      Determine_New_Hex_Map_Down_Letter = "V"
   ElseIf MAP_DOWN < 483 Then
      Determine_New_Hex_Map_Down_Letter = "W"
   ElseIf MAP_DOWN < 504 Then
      Determine_New_Hex_Map_Down_Letter = "X"
   ElseIf MAP_DOWN < 525 Then
      Determine_New_Hex_Map_Down_Letter = "Y"
   ElseIf MAP_DOWN < 546 Then
      Determine_New_Hex_Map_Down_Letter = "Z"
   End If

End Function

Function DICE_ROLL(CLAN, TRIBE)

Dim Total_Provs As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
If CLAN = "AAA" Then
  Total_People = 0
Else
  Tribe_Checking_People = 0
  Tribe_Checking_Provs = 0
  Call Tribe_Checking("Get_People", CLAN, TRIBE, "")
  Total_People = Tribe_Checking_People
  Call Tribe_Checking("Get_Provs", CLAN, TRIBE, "")
  Total_Provs = Tribe_Checking_Provs

  Total_People = (Total_People / (Total_Provs + 1))
End If

CURRENT_TIME = Time$
HOURS = Mid(CURRENT_TIME, 4, 2)
SECONDS = Right(CURRENT_TIME, 2)
RANDOM_TIME = (HOURS * 60) + SECONDS + Total_People
Randomize (RANDOM_TIME)

DICE = Int(100 * Rnd + 1)

DICE_ROLL = DICE

End Function

Function diceroll()
' Only called when opening the database

roll1 = DROLL(6, sklevel, 100, 0, DICE_TRIBE, 0, 0)

' Due to problems with linking, automating a relink
' Check GM

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

Call SET_GM(GMTABLE![Name], "No")


End Function

Function do_herd_swap()

    CLAN1 = InputBox("Who is the first clan?", "HERD_SWAP", "0")
    TRIBE1 = InputBox("Who is the first tribe?", "HERD_SWAP", "0")
    CLAN2 = InputBox("Who is the second clan?", "HERD_SWAP", "0")
    TRIBE2 = InputBox("Who is the second tribe?", "HERD_SWAP", "0")
    
    ANIMAL = InputBox("Were GOATS swapped?", "HERD_SWAP", "N")
    
    If ANIMAL = "Y" Then
       Call HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, "GOAT", "Y")
    End If

    ANIMAL = InputBox("Were SHEEP swapped?", "HERD_SWAP", "N")
    
    If ANIMAL = "Y" Then
       Call HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, "SHEEP", "Y")
    End If

    ANIMAL = InputBox("Were CATTLE swapped?", "HERD_SWAP", "N")
    
    If ANIMAL = "Y" Then
       Call HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, "CATTLE", "Y")
    End If

    ANIMAL = InputBox("Were HORSES swapped?", "HERD_SWAP", "N")
    
    If ANIMAL = "Y" Then
       Call HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, "HORSE", "Y")
    End If

    ANIMAL = InputBox("Were DOGS swapped?", "HERD_SWAP", "N")
    
    If ANIMAL = "Y" Then
       Call HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, "DOG", "Y")
    End If

    ANIMAL = InputBox("Were Elephants swapped?", "HERD_SWAP", "N")
    
    If ANIMAL = "Y" Then
       Call HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, "ELEPHANT", "Y")
    End If


End Function



Function GET_SEASON(TURN)
Dim SEASON As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set Globaltable = TVDBGM.OpenRecordset("GLOBAL")
Globaltable.index = "PRIMARYKEY"
Globaltable.MoveFirst

If Left(Globaltable![CURRENT TURN], 2) = "01" Then
   GET_SEASON = "Spring"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "02" Then
   GET_SEASON = "Spring"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "03" Then
   GET_SEASON = "Spring"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "04" Then
   GET_SEASON = "Summer"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "05" Then
   GET_SEASON = "Summer"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "06" Then
   GET_SEASON = "Summer"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "07" Then
   GET_SEASON = "Autumn"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "08" Then
   GET_SEASON = "Autumn"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "09" Then
   GET_SEASON = "Autumn"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "10" Then
   GET_SEASON = "Winter"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "11" Then
   GET_SEASON = "Winter"
ElseIf Left(Globaltable![CURRENT TURN], 2) = "12" Then
   GET_SEASON = "Winter"
End If


End Function

Function global_info()
Dim OriginalClan As String
Dim OriginalTribe As String

DoCmd.Hourglass True

DoCmd.Close A_FORM, "GLOBAL INFO"
DoCmd.OpenForm "GLOBAL INFO"

ShowUpdateTribesVersion = 0
'---update previous hex with string from previous hex (AB 20230729)
    DebugOP "UPDATE Previous_Hex with CURRENT HEX"
    Dim sSQL As String
    sSQL = "UPDATE TRIBES_general_info SET TRIBES_general_info.Previous_Hex = [CURRENT HEX];"
    
    CurrentDb.Execute sSQL, dbFailOnError
'---

Forms![GLOBAL INFO]![Status] = "Start Global Processing"
Forms![GLOBAL INFO].Repaint
DebugOP "Start Global Processing"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GMTABLE.Edit
GMTABLE![cALc_costs_PROCESSED] = "N"
GMTABLE![FINAL_ACTIVITIES_PROCESSED] = "N"
GMTABLE.UPDATE

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst

Set FarmingTable = TVDBGM.OpenRecordset("HEXMAP_FARMING")
FarmingTable.index = "PRIMARYKEY"

Set GAMES_WEATHER = TVDBGM.OpenRecordset("GAMES_WEATHER")
GAMES_WEATHER.index = "PRIMARYKEY"

Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst

If Left(globalinfo![CURRENT TURN], 2) = "01" Then
    DebugOP "Clean up Farming and Weather"

   Forms![GLOBAL INFO]![Status] = "Clean up Farming and Weather"
   Forms![GLOBAL INFO].Repaint
   
   Do Until FarmingTable.EOF
       
      FarmingTable.Delete
      FarmingTable.MoveNext
   Loop
   Do Until GAMES_WEATHER.EOF
      GAMES_WEATHER.Delete
      GAMES_WEATHER.MoveNext
   Loop
End If

TRIBEINFO.MoveFirst

RECORD_DELETED = "NO"

Forms![GLOBAL INFO]![Status] = "Clean up absorbed tribes and update weather"
Forms![GLOBAL INFO].Repaint

DebugOP "Clean up absorbed tribes and update weather"

Do Until TRIBEINFO.EOF
   TRIBEINFO.Edit
   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", TRIBEINFO![CURRENT HEX]
   If HEXTABLE.NoMatch Then
      ' we have a problem with the hex information for the tribe
      Msg = TRIBEINFO![TRIBE] & " has a blank Current Hex which you need to fix"
      Response = MsgBox(Msg, True)
      Msg = "I will exit now so you can fix the data and restart"
      Response = MsgBox(Msg, True)
      GoTo Time_To_Leave
   ElseIf IsNull(HEXTABLE![WEATHER_ZONE]) Then
      ' we have a problem with the weather zone for the hexe
      Msg = "Hex " & TRIBEINFO![CURRENT HEX] & " has a null/blank weather zone and you will need to populate it to have everything process properly"
      Response = MsgBox(Msg, True)
       Msg = "I will exit now so you can fix the data and restart"
      Response = MsgBox(Msg, True)
      GoTo Time_To_Leave
   ElseIf TRIBEINFO![ABSORBED] = "Y" Then
      TRIBEINFO.Delete
      RECORD_DELETED = "YES"
   ElseIf Not IsNull(TRIBEINFO![TRIBE]) Then
      GAMES_WEATHER.Seek "=", HEXTABLE![WEATHER_ZONE], globalinfo![CURRENT TURN]
      If GAMES_WEATHER.NoMatch Then
         GAMES_WEATHER.AddNew
         GAMES_WEATHER![WEATHER_ZONE] = HEXTABLE![WEATHER_ZONE]
         GAMES_WEATHER![TURN] = globalinfo![CURRENT TURN]
         GAMES_WEATHER.UPDATE
         GAMES_WEATHER.Seek "=", HEXTABLE![WEATHER_ZONE], globalinfo![CURRENT TURN]
      End If
      
      GAMES_WEATHER.Edit
      
      If HEXTABLE![WEATHER_ZONE] = "GREEN" Then
         GAMES_WEATHER![WEATHER] = globalinfo![Zone1]
      ElseIf HEXTABLE![WEATHER_ZONE] = "RED" Then
         GAMES_WEATHER![WEATHER] = globalinfo![Zone2]
      ElseIf HEXTABLE![WEATHER_ZONE] = "ORANGE" Then
         GAMES_WEATHER![WEATHER] = globalinfo![Zone3]
      ElseIf HEXTABLE![WEATHER_ZONE] = "YELLOW" Then
         GAMES_WEATHER![WEATHER] = globalinfo![Zone4]
      ElseIf HEXTABLE![WEATHER_ZONE] = "BLUE" Then
         GAMES_WEATHER![WEATHER] = globalinfo![Zone5]
      ElseIf HEXTABLE![WEATHER_ZONE] = "BROWN" Then
         GAMES_WEATHER![WEATHER] = globalinfo![Zone6]
      End If
      GAMES_WEATHER.UPDATE
   End If

TRIBEINFO.MoveNext

Loop
' Conversion of HEX_MAP_CONST disabled after conversion is done
' Call ConvertBuildingsTable

Forms![GLOBAL INFO]![Status] = "Set up Herd Swaps"
Forms![GLOBAL INFO].Repaint
DebugOP "Set up Herd Swaps"


Set HSTABLE = TVDBGM.OpenRecordset("HERD_SWAPS")
HSTABLE.index = "PRIMARYKEY"
HSTABLE.AddNew
HSTABLE![TRIBE] = "zzzz"
HSTABLE![tribe swapped with] = "zzzz"
HSTABLE![ANIMAL] = "zzzz"
HSTABLE![turns to go] = 0
HSTABLE.UPDATE
HSTABLE.MoveFirst

Do Until HSTABLE.EOF
   If HSTABLE![turns to go] = 0 Then
      HSTABLE.Delete
   End If
   HSTABLE.MoveNext
   If HSTABLE.EOF Then
      Exit Do
   End If
Loop

' if Trading_Post not empty then set the relevant bits.
' loop through clans setting switch.
Forms![GLOBAL INFO]![Status] = "Clean up printing switches"
Forms![GLOBAL INFO].Repaint
DebugOP "Clean up printing switches"

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Printing_Switchs;")
qdfCurrent.Execute

TRIBEINFO.MoveFirst
Set Printing_Switch_TABLE = TVDB.OpenRecordset("Printing_Switchs")
Printing_Switch_TABLE.index = "PRIMARYKEY"
CLANNUMBER = TRIBEINFO![CLAN]
PREV_CLANNUMBER = "NONE"

Do Until TRIBEINFO.EOF
   If Not (CLANNUMBER = PREV_CLANNUMBER) Then
      PREV_CLANNUMBER = CLANNUMBER
  
      Printing_Switch_TABLE.Seek "=", CLANNUMBER
      
      If Printing_Switch_TABLE.NoMatch Then
         Printing_Switch_TABLE.AddNew
         Printing_Switch_TABLE![CLAN] = CLANNUMBER
         Printing_Switch_TABLE![CITY] = Forms![GLOBAL INFO]![Trading_Post]
         Printing_Switch_TABLE.UPDATE
      End If
   End If
   TRIBEINFO.MoveNext
   If TRIBEINFO.EOF Then
      Exit Do
   End If
   CLANNUMBER = TRIBEINFO![CLAN]
  
Loop

globalinfo.Close
TRIBEINFO.Close
FarmingTable.Close
HEXTABLE.Close
GAMES_WEATHER.Close

Forms![GLOBAL INFO]![Status] = "Call Update Trading Post Goods Table"
Forms![GLOBAL INFO].Repaint
DebugOP "Call Update Trading Post Goods Table"

Call UPDATE_TRADING_POST_GOODS_TABLE

Forms![GLOBAL INFO]![Status] = "Clean up and reset"
Forms![GLOBAL INFO].Repaint
DebugOP "Clean up and reset"

Call CLEAN_UP_and_RESET

Forms![GLOBAL INFO]![Status] = "Check GL Level against completed research"
Forms![GLOBAL INFO].Repaint
DebugOP "Check GL Level against completed research"

Call Check_GL_Level_against_Completed_Research

Forms![GLOBAL INFO]![Status] = "Ensure TP's are setup"
Forms![GLOBAL INFO].Repaint
DebugOP "Ensure TP's are setup"

Call Ensure_TPs_Are_Setup

Forms![GLOBAL INFO]![Status] = "Update Tribe Info"
Forms![GLOBAL INFO].Repaint
DebugOP "Update Tribe Info"

'DoCmd.OpenQuery "UPDATE TRIBE INFO"
CurrentDb.Execute "UPDATE TRIBE INFO", dbFailOnError

Forms![GLOBAL INFO]![Status] = "End Global Processing"
Forms![GLOBAL INFO].Repaint
DebugOP "End Global Processing"


Time_To_Leave:
DoCmd.Hourglass False

End Function

Function HERD_SWAPS(CLAN1, TRIBE1, CLAN2, TRIBE2, ANIMAL, SWAP)
Dim SWAPINEFFECT As Long

SWAPINEFFECT = 0

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set HSTABLE = TVDBGM.OpenRecordset("HERD_SWAPS")
If SWAP = "Y" Then
   HSTABLE.index = "PRIMARYKEY"
   HSTABLE.MoveFirst
   HSTABLE.Seek "=", TRIBE1, TRIBE2, ANIMAL

   If HSTABLE.NoMatch Then
      HSTABLE.AddNew
      HSTABLE![TRIBE] = TRIBE1
      HSTABLE![tribe swapped with] = TRIBE2
      HSTABLE![ANIMAL] = ANIMAL
      HSTABLE![turns to go] = 3
      HSTABLE.UPDATE
   ElseIf HSTABLE![turns to go] = 0 Then
      HSTABLE.Edit
      HSTABLE![turns to go] = 3
      HSTABLE.UPDATE
   End If

   HSTABLE.Seek "=", TRIBE2, TRIBE1, ANIMAL
  
   If HSTABLE.NoMatch Then
      HSTABLE.AddNew
      HSTABLE![TRIBE] = TRIBE2
      HSTABLE![tribe swapped with] = TRIBE1
      HSTABLE![ANIMAL] = ANIMAL
      HSTABLE![turns to go] = 3
      HSTABLE.UPDATE
   ElseIf HSTABLE![turns to go] = 0 Then
      HSTABLE.Edit
      HSTABLE![turns to go] = 3
      HSTABLE.UPDATE
   End If

   ' UPDATE TURNS ACTIVITIES WITH THE HERD SWAP DETAILS

   LineI = 1
  
   OutLine = "Herd Swap Current with Tribe " & TRIBE2
   Call WRITE_TURN_ACTIVITY(CLAN1, TRIBE1, "HERD SWAPS", LineI, OutLine, "No")
   OutLine = "Herd Swap Current with Tribe " & TRIBE1
   Call WRITE_TURN_ACTIVITY(CLAN2, TRIBE2, "HERD SWAPS", LineI, OutLine, "No")

Else
   HSTABLE.index = "TRIBE"
   HSTABLE.MoveFirst
   HSTABLE.Seek "=", TRIBE1, ANIMAL

   If Not HSTABLE.NoMatch Then
      Do While HSTABLE![TRIBE] = TRIBE1 And HSTABLE![ANIMAL] = ANIMAL
         If HSTABLE![ANIMAL] = ANIMAL Then
            If HSTABLE![turns to go] = 0 Then
               HSTABLE.MoveNext
            Else
               HSTABLE.Edit
               HSTABLE![turns to go] = HSTABLE![turns to go] - 1
               HSTABLE.UPDATE
               SWAPINEFFECT = SWAPINEFFECT + 1
               HSTABLE.MoveNext
            End If
         Else
            HSTABLE.MoveNext
         End If
      Loop
   End If
End If

HERD_SWAPS = SWAPINEFFECT

End Function

Function NEW_BANDIT()

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
TCLANNUMBER = InputBox("What is the new CLAN number?", "NEWCLAN", "BS00")

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.AddNew
TRIBESINFO![CLAN] = TCLANNUMBER
TRIBESINFO![TRIBE] = TCLANNUMBER
TRIBESINFO![Village] = "Tribe"
TRIBESINFO![CURRENT TERRAIN] = "PRAIRIE"
TRIBESINFO![GOODS_CLAN] = TCLANNUMBER
TRIBESINFO![GOODS TRIBE] = TCLANNUMBER
TRIBESINFO![COST CLAN] = TCLANNUMBER
TRIBESINFO![WARRIORS] = 80
TRIBESINFO![ACTIVES] = 80
TRIBESINFO![INACTIVES] = 80
TRIBESINFO![MORALE] = 1
TRIBESINFO.UPDATE
TRIBESINFO.Close

    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "CATTLE", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "GOAT", "ADD", 3700)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "HORSE", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "COAL", "ADD", 3000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "IRON", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SILVER", "ADD", 10000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "CLUB", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SHIELD", "ADD", 30)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "JERKIN", "ADD", 200)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SWORD", "ADD", 30)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "PROVS", "ADD", 50000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SLING", "ADD", 300)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "WAGON", "ADD", 300)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BARK", "ADD", 1000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BONES", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "GUT", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "LEATHER", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "WAX", "ADD", 20)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SKIN", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BRONZE", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "LOG", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BRASS", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "TRAP", "ADD", 500)


    
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ADMINISTRATION", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "BONEWORK", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "BONING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "CURING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "DIPLOMACY", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ECONOMICS", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ENGINEERING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "FORESTRY", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "GARRISON", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "GUTTING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "HERDING", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "HUNTING", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "LEADERSHIP", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "LEATHERWORK", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "QUARRYING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "SCOUTING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "SKINNING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "TANNING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "WOODWORK", 3)

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TRIBE_CHECKING;")
qdfCurrent.Execute

QUERY_STRING = "INSERT INTO TRIBE_CHECKING ( CLAN, TRIBE, [CURRENT HEX], WARRIORS, ACTIVES,"
QUERY_STRING = QUERY_STRING & " INACTIVES, SLAVE ) SELECT CLAN, TRIBE, [CURRENT HEX],"
QUERY_STRING = QUERY_STRING & " WARRIORS, ACTIVES, INACTIVES, SLAVE"
QUERY_STRING = QUERY_STRING & " FROM TRIBES_GENERAL_INFO;"
Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute

End Function

Function New_Clan()
Dim HEX_MAP As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
TCLANNUMBER = InputBox("What is the new CLAN number?", "NEWCLAN", "0000")

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.AddNew
TRIBESINFO![CLAN] = TCLANNUMBER
TRIBESINFO![TRIBE] = TCLANNUMBER
TRIBESINFO![Village] = "Tribe"
TRIBESINFO![CURRENT TERRAIN] = "PRAIRIE"
TRIBESINFO![GOODS_CLAN] = TCLANNUMBER
TRIBESINFO![GOODS TRIBE] = TCLANNUMBER
TRIBESINFO![COST CLAN] = TCLANNUMBER
TRIBESINFO![WARRIORS] = 5890
TRIBESINFO![ACTIVES] = 5890
TRIBESINFO![INACTIVES] = 5890
TRIBESINFO![MORALE] = 1
' GET THE HEX THAT THE TRIBE IS IN.
HEX_MAP = InputBox("What HEX is the new CLAN to live in?", "NEWCLAN", "AA 0101")
TRIBESINFO![CURRENT HEX] = HEX_MAP
TRIBESINFO.UPDATE
TRIBESINFO.Close

    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "CATTLE", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "GOAT", "ADD", 3700)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "HORSE", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "COAL", "ADD", 3000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "IRON", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SILVER", "ADD", 10000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "CLUB", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SHIELD", "ADD", 30)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "JERKIN", "ADD", 200)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SWORD", "ADD", 30)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "PROVS", "ADD", 50000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SLING", "ADD", 300)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "WAGON", "ADD", 300)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BARK", "ADD", 1000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BONES", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "GUT", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "LEATHER", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "WAX", "ADD", 20)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SKIN", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BRONZE", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "LOG", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BRASS", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "TRAP", "ADD", 500)


    
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ADMINISTRATION", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "BONEWORK", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "BONING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "CURING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "DIPLOMACY", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ECONOMICS", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ENGINEERING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "FORESTRY", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "GARRISON", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "GUTTING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "HERDING", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "HUNTING", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "LEADERSHIP", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "LEATHERWORK", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "QUARRYING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "SCOUTING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "SKINNING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "TANNING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "WOODWORK", 3)

OutLine = "You may distribute a further 30 skill points (new, or add to existing skills) but no skill level may exceed 7."
OutLine = OutLine & "{enter}"

Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 1, OutLine, "No")


OutLine = "You may choose to take either 900 iron or 1200 bronze."
OutLine = OutLine & "{enter}{enter}"

Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 2, OutLine, "No")

OutLine = "Special skills / items:"
OutLine = OutLine & "{enter}"

Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 3, OutLine, "No")

OutLine = "Horse Bow:  can make with Wpn6, horse archers can participate in missile phase and charge phase."
OutLine = OutLine & "{enter}"

Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 4, OutLine, "No")
OutLine = "Elephants:      can handle, may use elephants for transport."
OutLine = OutLine & "{enter}"

Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 5, OutLine, "No")

Call Tribe_Checking("Update_All", "", "", "")

End Function

Function NEW_GROUP()

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.Seek "=", Forms![NEW GROUP]![PARENT CLAN], Forms![NEW GROUP]![PARENT TRIBE]
CURRENT_HEX = TRIBESINFO![CURRENT HEX]
CURRENT_COST_CLAN = TRIBESINFO![COST CLAN]
CURRENT_TERRAIN = TRIBESINFO![CURRENT TERRAIN]
If IsNull(TRIBESINFO![RELIGION]) Then
   RELIGION = ""
Else
   RELIGION = TRIBESINFO![RELIGION]
End If
If IsNull(TRIBESINFO![CULT]) Then
   CULT = ""
Else
   CULT = TRIBESINFO![CULT]
End If

TRIBESINFO.AddNew
TRIBESINFO![CLAN] = Forms![NEW GROUP]![New Clan]
TRIBESINFO![TRIBE] = Forms![NEW GROUP]![New Tribe]
TRIBESINFO![TRIBE NAME] = Null
TRIBESINFO![CURRENT HEX] = CURRENT_HEX
TRIBESINFO![CURRENT TERRAIN] = CURRENT_TERRAIN
TRIBESINFO![Village] = Forms![NEW GROUP]![Village]
If Not IsNull(RELIGION) And Not (RELIGION = "") Then
   TRIBESINFO![RELIGION] = RELIGION
End If
If Not IsNull(CULT) And Not (CULT = "") Then
   TRIBESINFO![CULT] = CULT
End If
TRIBESINFO![CREDIT] = 0
TRIBESINFO![AMT RECEIVED] = 0
TRIBESINFO![Cost] = 0
TRIBESINFO![OWNER] = Null
TRIBESINFO![EMAIL] = "N"
TRIBESINFO![GOODS_CLAN] = Null
TRIBESINFO![GOODS TRIBE] = Null
TRIBESINFO![POP TRIBE] = Null
TRIBESINFO![COST CLAN] = CURRENT_COST_CLAN
TRIBESINFO![MORALE] = 1
TRIBESINFO.UPDATE
TRIBESINFO.Close

Set TRIBES_CHECKING = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TRIBES_CHECKING.index = "PRIMARYKEY"
TRIBES_CHECKING.AddNew
TRIBES_CHECKING![CLAN] = Forms![NEW GROUP]![New Clan]
TRIBES_CHECKING![TRIBE] = Forms![NEW GROUP]![New Tribe]
TRIBES_CHECKING![CURRENT HEX] = CURRENT_HEX
TRIBES_CHECKING.UPDATE
TRIBES_CHECKING.Close

Set Globaltable = TVDBGM.OpenRecordset("GLOBAL")
Globaltable.MoveFirst
Current_Turn = Globaltable![CURRENT TURN]
Globaltable.Close

Set FarmingTable = TVDBGM.OpenRecordset("TRIBE_FARMING")
FarmingTable.index = "PRIMARYKEY"
FarmingTable.Seek "=", Forms![NEW GROUP]![New Clan], Forms![NEW GROUP]![New Tribe], Current_Turn

If FarmingTable.NoMatch Then
   FarmingTable.AddNew
   FarmingTable![CLAN] = Forms![NEW GROUP]![New Clan]
   FarmingTable![TRIBE] = Forms![NEW GROUP]![New Tribe]
   FarmingTable![TURN] = Current_Turn
   FarmingTable![ITEM] = "START"
   FarmingTable.UPDATE
   FarmingTable.Close
End If

Call EXIT_FORMS("NEW GROUP")
Call OPEN_FORMS("NEW GROUP")

End Function

Function NEW_MAP_NUMBER(MAPNUM, Direction)
Dim x As Long
Dim Y As Long

x = Asc(Mid(MAPNUM, 1, 1)) - 65
Y = Asc(Mid(MAPNUM, 2, 1)) - 65
If Direction = "UP" Then x = x - 1
If Direction = "DOWN" Then x = x + 1
If Direction = "LEFT" Then Y = Y - 1
If Direction = "RIGHT" Then Y = Y + 1
If Direction = "UP&LEFT" Then x = x - 1 & Y = Y - 1
If Direction = "DOWN&RIGHT" Then x = x + 1 & Y = Y + 1
If x < 0 Then x = 15
If Y < 0 Then Y = 15
If x > 15 Then x = 0
If Y > 15 Then Y = 0
NEW_MAP_NUMBER = Chr$(x + 65) & Chr$(Y + 65)

End Function

Function OPEN_TRIBEVIB_PRINT()
    Set wrdDoc = wrdApp.Documents.Add

    wrdApp.ActiveDocument.SaveAs DIRECTPATH & "\tribevib"
    wrdApp.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
    wrdApp.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)
    wrdApp.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1)
    wrdApp.ActiveDocument.PageSetup.BottomMargin = CentimetersToPoints(1.5)
    wrdApp.ActiveDocument.PageSetup.PaperSize = wdPaperA4
    wrdApp.ActiveDocument.Range.ParagraphFormat.SpaceAfter = 0

    With wrdApp.Selection
       
       
       .Font.Name = "Times New Roman"
       .Font.Size = 10
       ' clear all tabstops
       .Paragraphs(1).TabStops.ClearAll
       ' add in tabstops
       .Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(1), Alignment:=wdAlignTabLeft
       .Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(2.5), Alignment:=wdAlignTabLeft
       .Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(3.5), Alignment:=wdAlignTabLeft
       .Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(4.5), Alignment:=wdAlignTabLeft
       
    End With

    wrdDoc.Save
    wrdDoc.Activate

End Function

Function PAD_HEXES()
On Error GoTo ERR_PAD

Dim MAPNUMBER As String
Dim ACROSS As Long
Dim DOWN As Long
Dim HEXNUMBER As String
ReDim NEW_MAP(26, 26) As String
Dim NUMBER As String

NEW_MAP(1, 1) = "AA"
NEW_MAP(1, 2) = "AB"
NEW_MAP(1, 3) = "AC"
NEW_MAP(1, 4) = "AD"
NEW_MAP(1, 5) = "AE"
NEW_MAP(1, 6) = "AF"
NEW_MAP(1, 7) = "AG"
NEW_MAP(1, 8) = "AH"
NEW_MAP(1, 9) = "AI"
NEW_MAP(1, 10) = "AJ"
NEW_MAP(1, 11) = "AK"
NEW_MAP(1, 12) = "AL"
NEW_MAP(1, 13) = "AM"
NEW_MAP(1, 14) = "AN"
NEW_MAP(1, 15) = "AO"
NEW_MAP(1, 16) = "AP"
NEW_MAP(1, 17) = "AQ"
NEW_MAP(1, 18) = "AR"
NEW_MAP(1, 19) = "AS"
NEW_MAP(1, 20) = "AT"
NEW_MAP(1, 21) = "AU"
NEW_MAP(1, 22) = "AV"
NEW_MAP(1, 23) = "AW"
NEW_MAP(1, 24) = "AX"
NEW_MAP(1, 25) = "AY"
NEW_MAP(1, 26) = "AZ"
NEW_MAP(2, 1) = "BA"
NEW_MAP(2, 2) = "BB"
NEW_MAP(2, 3) = "BC"
NEW_MAP(2, 4) = "BD"
NEW_MAP(2, 5) = "BE"
NEW_MAP(2, 6) = "BF"
NEW_MAP(2, 7) = "BG"
NEW_MAP(2, 8) = "BH"
NEW_MAP(2, 9) = "BI"
NEW_MAP(2, 10) = "BJ"
NEW_MAP(2, 11) = "BK"
NEW_MAP(2, 12) = "BL"
NEW_MAP(2, 13) = "BM"
NEW_MAP(2, 14) = "BN"
NEW_MAP(2, 15) = "BO"
NEW_MAP(2, 16) = "BP"
NEW_MAP(2, 17) = "BQ"
NEW_MAP(2, 18) = "BR"
NEW_MAP(2, 19) = "BS"
NEW_MAP(2, 20) = "BT"
NEW_MAP(2, 21) = "BU"
NEW_MAP(2, 22) = "BV"
NEW_MAP(2, 23) = "BW"
NEW_MAP(2, 24) = "BX"
NEW_MAP(2, 25) = "BY"
NEW_MAP(2, 26) = "BZ"
NEW_MAP(3, 1) = "CA"
NEW_MAP(3, 2) = "CB"
NEW_MAP(3, 3) = "CC"
NEW_MAP(3, 4) = "CD"
NEW_MAP(3, 5) = "CE"
NEW_MAP(3, 6) = "CF"
NEW_MAP(3, 7) = "CG"
NEW_MAP(3, 8) = "CH"
NEW_MAP(3, 9) = "CI"
NEW_MAP(3, 10) = "CJ"
NEW_MAP(3, 11) = "CK"
NEW_MAP(3, 12) = "CL"
NEW_MAP(3, 13) = "CM"
NEW_MAP(3, 14) = "CN"
NEW_MAP(3, 15) = "CO"
NEW_MAP(3, 16) = "CP"
NEW_MAP(3, 17) = "CQ"
NEW_MAP(3, 18) = "CR"
NEW_MAP(3, 19) = "CS"
NEW_MAP(3, 20) = "CT"
NEW_MAP(3, 21) = "CU"
NEW_MAP(3, 22) = "CV"
NEW_MAP(3, 23) = "CW"
NEW_MAP(3, 24) = "CX"
NEW_MAP(3, 25) = "CY"
NEW_MAP(3, 26) = "CZ"
NEW_MAP(4, 1) = "DA"
NEW_MAP(4, 2) = "DB"
NEW_MAP(4, 3) = "DC"
NEW_MAP(4, 4) = "DD"
NEW_MAP(4, 5) = "DE"
NEW_MAP(4, 6) = "DF"
NEW_MAP(4, 7) = "DG"
NEW_MAP(4, 8) = "DH"
NEW_MAP(4, 9) = "DI"
NEW_MAP(4, 10) = "DJ"
NEW_MAP(4, 11) = "DK"
NEW_MAP(4, 12) = "DL"
NEW_MAP(4, 13) = "DM"
NEW_MAP(4, 14) = "DN"
NEW_MAP(4, 15) = "DO"
NEW_MAP(4, 16) = "DP"
NEW_MAP(4, 17) = "DQ"
NEW_MAP(4, 18) = "DR"
NEW_MAP(4, 19) = "DS"
NEW_MAP(4, 20) = "DT"
NEW_MAP(4, 21) = "DU"
NEW_MAP(4, 22) = "DV"
NEW_MAP(4, 23) = "DW"
NEW_MAP(4, 24) = "DX"
NEW_MAP(4, 25) = "DY"
NEW_MAP(4, 26) = "DZ"
NEW_MAP(5, 1) = "EA"
NEW_MAP(5, 2) = "EB"
NEW_MAP(5, 3) = "EC"
NEW_MAP(5, 4) = "ED"
NEW_MAP(5, 5) = "EE"
NEW_MAP(5, 6) = "EF"
NEW_MAP(5, 7) = "EG"
NEW_MAP(5, 8) = "EH"
NEW_MAP(5, 9) = "EI"
NEW_MAP(5, 10) = "EJ"
NEW_MAP(5, 11) = "EK"
NEW_MAP(5, 12) = "EL"
NEW_MAP(5, 13) = "EM"
NEW_MAP(5, 14) = "EN"
NEW_MAP(5, 15) = "EO"
NEW_MAP(5, 16) = "EP"
NEW_MAP(5, 17) = "EQ"
NEW_MAP(5, 18) = "ER"
NEW_MAP(5, 19) = "ES"
NEW_MAP(5, 20) = "ET"
NEW_MAP(5, 21) = "EU"
NEW_MAP(5, 22) = "EV"
NEW_MAP(5, 23) = "EW"
NEW_MAP(5, 24) = "EX"
NEW_MAP(5, 25) = "EY"
NEW_MAP(5, 26) = "EZ"
NEW_MAP(6, 1) = "FA"
NEW_MAP(6, 2) = "FB"
NEW_MAP(6, 3) = "FC"
NEW_MAP(6, 4) = "FD"
NEW_MAP(6, 5) = "FE"
NEW_MAP(6, 6) = "FF"
NEW_MAP(6, 7) = "FG"
NEW_MAP(6, 8) = "FH"
NEW_MAP(6, 9) = "FI"
NEW_MAP(6, 10) = "FJ"
NEW_MAP(6, 11) = "FK"
NEW_MAP(6, 12) = "FL"
NEW_MAP(6, 13) = "FM"
NEW_MAP(6, 14) = "FN"
NEW_MAP(6, 15) = "FO"
NEW_MAP(6, 16) = "FP"
NEW_MAP(6, 17) = "FQ"
NEW_MAP(6, 18) = "FR"
NEW_MAP(6, 19) = "FS"
NEW_MAP(6, 20) = "FT"
NEW_MAP(6, 21) = "FU"
NEW_MAP(6, 22) = "FV"
NEW_MAP(6, 23) = "FW"
NEW_MAP(6, 24) = "FX"
NEW_MAP(6, 25) = "FY"
NEW_MAP(6, 26) = "FZ"
NEW_MAP(7, 1) = "GA"
NEW_MAP(7, 2) = "GB"
NEW_MAP(7, 3) = "GC"
NEW_MAP(7, 4) = "GD"
NEW_MAP(7, 5) = "GE"
NEW_MAP(7, 6) = "GF"
NEW_MAP(7, 7) = "GG"
NEW_MAP(7, 8) = "GH"
NEW_MAP(7, 9) = "GI"
NEW_MAP(7, 10) = "GJ"
NEW_MAP(7, 11) = "GK"
NEW_MAP(7, 12) = "GL"
NEW_MAP(7, 13) = "GM"
NEW_MAP(7, 14) = "GN"
NEW_MAP(7, 15) = "GO"
NEW_MAP(7, 16) = "GP"
NEW_MAP(7, 17) = "GQ"
NEW_MAP(7, 18) = "GR"
NEW_MAP(7, 19) = "GS"
NEW_MAP(7, 20) = "GT"
NEW_MAP(7, 21) = "GU"
NEW_MAP(7, 22) = "GV"
NEW_MAP(7, 23) = "GW"
NEW_MAP(7, 24) = "GX"
NEW_MAP(7, 25) = "GY"
NEW_MAP(7, 26) = "GZ"
NEW_MAP(8, 1) = "HA"
NEW_MAP(8, 2) = "HB"
NEW_MAP(8, 3) = "HC"
NEW_MAP(8, 4) = "HD"
NEW_MAP(8, 5) = "HE"
NEW_MAP(8, 6) = "HF"
NEW_MAP(8, 7) = "HG"
NEW_MAP(8, 8) = "HH"
NEW_MAP(8, 9) = "HI"
NEW_MAP(8, 10) = "HJ"
NEW_MAP(8, 11) = "HK"
NEW_MAP(8, 12) = "HL"
NEW_MAP(8, 13) = "HM"
NEW_MAP(8, 14) = "HN"
NEW_MAP(8, 15) = "HO"
NEW_MAP(8, 16) = "HP"
NEW_MAP(8, 17) = "HQ"
NEW_MAP(8, 18) = "HR"
NEW_MAP(8, 19) = "HS"
NEW_MAP(8, 20) = "HT"
NEW_MAP(8, 21) = "HU"
NEW_MAP(8, 22) = "HV"
NEW_MAP(8, 23) = "HW"
NEW_MAP(8, 24) = "HX"
NEW_MAP(8, 25) = "HY"
NEW_MAP(8, 26) = "HZ"
NEW_MAP(9, 1) = "IA"
NEW_MAP(9, 2) = "IB"
NEW_MAP(9, 3) = "IC"
NEW_MAP(9, 4) = "ID"
NEW_MAP(9, 5) = "IE"
NEW_MAP(9, 6) = "IF"
NEW_MAP(9, 7) = "IG"
NEW_MAP(9, 8) = "IH"
NEW_MAP(9, 9) = "II"
NEW_MAP(9, 10) = "IJ"
NEW_MAP(9, 11) = "IK"
NEW_MAP(9, 12) = "IL"
NEW_MAP(9, 13) = "IM"
NEW_MAP(9, 14) = "IN"
NEW_MAP(9, 15) = "IO"
NEW_MAP(9, 16) = "IP"
NEW_MAP(9, 17) = "IQ"
NEW_MAP(9, 18) = "IR"
NEW_MAP(9, 19) = "IS"
NEW_MAP(9, 20) = "IT"
NEW_MAP(9, 21) = "IU"
NEW_MAP(9, 22) = "IV"
NEW_MAP(9, 23) = "IW"
NEW_MAP(9, 24) = "IX"
NEW_MAP(9, 25) = "IY"
NEW_MAP(9, 26) = "IZ"
NEW_MAP(10, 1) = "JA"
NEW_MAP(10, 2) = "JB"
NEW_MAP(10, 3) = "JC"
NEW_MAP(10, 4) = "JD"
NEW_MAP(10, 5) = "JE"
NEW_MAP(10, 6) = "JF"
NEW_MAP(10, 7) = "JG"
NEW_MAP(10, 8) = "JH"
NEW_MAP(10, 9) = "JI"
NEW_MAP(10, 10) = "JJ"
NEW_MAP(10, 11) = "JK"
NEW_MAP(10, 12) = "JL"
NEW_MAP(10, 13) = "JM"
NEW_MAP(10, 14) = "JN"
NEW_MAP(10, 15) = "JO"
NEW_MAP(10, 16) = "JP"
NEW_MAP(10, 17) = "JQ"
NEW_MAP(10, 18) = "JR"
NEW_MAP(10, 19) = "JS"
NEW_MAP(10, 20) = "JT"
NEW_MAP(10, 21) = "JU"
NEW_MAP(10, 22) = "JV"
NEW_MAP(10, 23) = "JW"
NEW_MAP(10, 24) = "JX"
NEW_MAP(10, 25) = "JY"
NEW_MAP(10, 26) = "JZ"
NEW_MAP(11, 1) = "KA"
NEW_MAP(11, 2) = "KB"
NEW_MAP(11, 3) = "KC"
NEW_MAP(11, 4) = "KD"
NEW_MAP(11, 5) = "KE"
NEW_MAP(11, 6) = "KF"
NEW_MAP(11, 7) = "KG"
NEW_MAP(11, 8) = "KH"
NEW_MAP(11, 9) = "KI"
NEW_MAP(11, 10) = "KJ"
NEW_MAP(11, 11) = "KK"
NEW_MAP(11, 12) = "KL"
NEW_MAP(11, 13) = "KM"
NEW_MAP(11, 14) = "KN"
NEW_MAP(11, 15) = "KO"
NEW_MAP(11, 16) = "KP"
NEW_MAP(11, 17) = "KQ"
NEW_MAP(11, 18) = "KR"
NEW_MAP(11, 19) = "KS"
NEW_MAP(11, 20) = "KT"
NEW_MAP(11, 21) = "KU"
NEW_MAP(11, 22) = "KV"
NEW_MAP(11, 23) = "KW"
NEW_MAP(11, 24) = "KX"
NEW_MAP(11, 25) = "KY"
NEW_MAP(11, 26) = "KZ"
NEW_MAP(12, 1) = "LA"
NEW_MAP(12, 2) = "LB"
NEW_MAP(12, 3) = "LC"
NEW_MAP(12, 4) = "LD"
NEW_MAP(12, 5) = "LE"
NEW_MAP(12, 6) = "LF"
NEW_MAP(12, 7) = "LG"
NEW_MAP(12, 8) = "LH"
NEW_MAP(12, 9) = "LI"
NEW_MAP(12, 10) = "LJ"
NEW_MAP(12, 11) = "LK"
NEW_MAP(12, 12) = "LL"
NEW_MAP(12, 13) = "LM"
NEW_MAP(12, 14) = "LN"
NEW_MAP(12, 15) = "LO"
NEW_MAP(12, 16) = "LP"
NEW_MAP(12, 17) = "LQ"
NEW_MAP(12, 18) = "LR"
NEW_MAP(12, 19) = "LS"
NEW_MAP(12, 20) = "LT"
NEW_MAP(12, 21) = "LU"
NEW_MAP(12, 22) = "LV"
NEW_MAP(12, 23) = "LW"
NEW_MAP(12, 24) = "LX"
NEW_MAP(12, 25) = "LY"
NEW_MAP(12, 26) = "LZ"
NEW_MAP(13, 1) = "MA"
NEW_MAP(13, 2) = "MB"
NEW_MAP(13, 3) = "MC"
NEW_MAP(13, 4) = "MD"
NEW_MAP(13, 5) = "ME"
NEW_MAP(13, 6) = "MF"
NEW_MAP(13, 7) = "MG"
NEW_MAP(13, 8) = "MH"
NEW_MAP(13, 9) = "MI"
NEW_MAP(13, 10) = "MJ"
NEW_MAP(13, 11) = "MK"
NEW_MAP(13, 12) = "ML"
NEW_MAP(13, 13) = "MM"
NEW_MAP(13, 14) = "MN"
NEW_MAP(13, 15) = "MO"
NEW_MAP(13, 16) = "MP"
NEW_MAP(13, 17) = "MQ"
NEW_MAP(13, 18) = "MR"
NEW_MAP(13, 19) = "MS"
NEW_MAP(13, 20) = "MT"
NEW_MAP(13, 21) = "MU"
NEW_MAP(13, 22) = "MV"
NEW_MAP(13, 23) = "MW"
NEW_MAP(13, 24) = "MX"
NEW_MAP(13, 25) = "MY"
NEW_MAP(13, 26) = "MZ"
NEW_MAP(14, 1) = "NA"
NEW_MAP(14, 2) = "NB"
NEW_MAP(14, 3) = "NC"
NEW_MAP(14, 4) = "ND"
NEW_MAP(14, 5) = "NE"
NEW_MAP(14, 6) = "NF"
NEW_MAP(14, 7) = "NG"
NEW_MAP(14, 8) = "NH"
NEW_MAP(14, 9) = "NI"
NEW_MAP(14, 10) = "NJ"
NEW_MAP(14, 11) = "NK"
NEW_MAP(14, 12) = "NL"
NEW_MAP(14, 13) = "NM"
NEW_MAP(14, 14) = "NN"
NEW_MAP(14, 15) = "NO"
NEW_MAP(14, 16) = "NP"
NEW_MAP(14, 17) = "NQ"
NEW_MAP(14, 18) = "NR"
NEW_MAP(14, 19) = "NS"
NEW_MAP(14, 20) = "NT"
NEW_MAP(14, 21) = "NU"
NEW_MAP(14, 22) = "NV"
NEW_MAP(14, 23) = "NW"
NEW_MAP(14, 24) = "NX"
NEW_MAP(14, 25) = "NY"
NEW_MAP(14, 26) = "NZ"
NEW_MAP(15, 1) = "OA"
NEW_MAP(15, 2) = "OB"
NEW_MAP(15, 3) = "OC"
NEW_MAP(15, 4) = "OD"
NEW_MAP(15, 5) = "OE"
NEW_MAP(15, 6) = "OF"
NEW_MAP(15, 7) = "OG"
NEW_MAP(15, 8) = "OH"
NEW_MAP(15, 9) = "OI"
NEW_MAP(15, 10) = "OJ"
NEW_MAP(15, 11) = "OK"
NEW_MAP(15, 12) = "OL"
NEW_MAP(15, 13) = "OM"
NEW_MAP(15, 14) = "ON"
NEW_MAP(15, 15) = "OO"
NEW_MAP(15, 16) = "OP"
NEW_MAP(15, 17) = "OQ"
NEW_MAP(15, 18) = "OR"
NEW_MAP(15, 19) = "OS"
NEW_MAP(15, 20) = "OT"
NEW_MAP(15, 21) = "OU"
NEW_MAP(15, 22) = "OV"
NEW_MAP(15, 23) = "OW"
NEW_MAP(15, 24) = "OX"
NEW_MAP(15, 25) = "OY"
NEW_MAP(15, 26) = "OZ"
NEW_MAP(16, 1) = "PA"
NEW_MAP(16, 2) = "PB"
NEW_MAP(16, 3) = "PC"
NEW_MAP(16, 4) = "PD"
NEW_MAP(16, 5) = "PE"
NEW_MAP(16, 6) = "PF"
NEW_MAP(16, 7) = "PG"
NEW_MAP(16, 8) = "PH"
NEW_MAP(16, 9) = "PI"
NEW_MAP(16, 10) = "PJ"
NEW_MAP(16, 11) = "PK"
NEW_MAP(16, 12) = "PL"
NEW_MAP(16, 13) = "PM"
NEW_MAP(16, 14) = "PN"
NEW_MAP(16, 15) = "PO"
NEW_MAP(16, 16) = "PP"
NEW_MAP(16, 17) = "PQ"
NEW_MAP(16, 18) = "PR"
NEW_MAP(16, 19) = "PS"
NEW_MAP(16, 20) = "PT"
NEW_MAP(16, 21) = "PU"
NEW_MAP(16, 22) = "PV"
NEW_MAP(16, 23) = "PW"
NEW_MAP(16, 24) = "PX"
NEW_MAP(16, 25) = "PY"
NEW_MAP(16, 26) = "PZ"
NEW_MAP(17, 1) = "QA"
NEW_MAP(17, 2) = "QB"
NEW_MAP(17, 3) = "QC"
NEW_MAP(17, 4) = "QD"
NEW_MAP(17, 5) = "QE"
NEW_MAP(17, 6) = "QF"
NEW_MAP(17, 7) = "QG"
NEW_MAP(17, 8) = "QH"
NEW_MAP(17, 9) = "QI"
NEW_MAP(17, 10) = "QJ"
NEW_MAP(17, 11) = "QK"
NEW_MAP(17, 12) = "QL"
NEW_MAP(17, 13) = "QM"
NEW_MAP(17, 14) = "QN"
NEW_MAP(17, 15) = "QO"
NEW_MAP(17, 16) = "QP"
NEW_MAP(17, 17) = "QQ"
NEW_MAP(17, 18) = "QR"
NEW_MAP(17, 19) = "QS"
NEW_MAP(17, 20) = "QT"
NEW_MAP(17, 21) = "QU"
NEW_MAP(17, 22) = "QV"
NEW_MAP(17, 23) = "QW"
NEW_MAP(17, 24) = "QX"
NEW_MAP(17, 25) = "QY"
NEW_MAP(17, 26) = "QZ"
NEW_MAP(18, 1) = "RA"
NEW_MAP(18, 2) = "RB"
NEW_MAP(18, 3) = "RC"
NEW_MAP(18, 4) = "RD"
NEW_MAP(18, 5) = "RE"
NEW_MAP(18, 6) = "RF"
NEW_MAP(18, 7) = "RG"
NEW_MAP(18, 8) = "RH"
NEW_MAP(18, 9) = "RI"
NEW_MAP(18, 10) = "RJ"
NEW_MAP(18, 11) = "RK"
NEW_MAP(18, 12) = "RL"
NEW_MAP(18, 13) = "RM"
NEW_MAP(18, 14) = "RN"
NEW_MAP(18, 15) = "RO"
NEW_MAP(18, 16) = "RP"
NEW_MAP(18, 17) = "RQ"
NEW_MAP(18, 18) = "RR"
NEW_MAP(18, 19) = "RS"
NEW_MAP(18, 20) = "RT"
NEW_MAP(18, 21) = "RU"
NEW_MAP(18, 22) = "RV"
NEW_MAP(18, 23) = "RW"
NEW_MAP(18, 24) = "RX"
NEW_MAP(18, 25) = "RY"
NEW_MAP(18, 26) = "RZ"
NEW_MAP(19, 1) = "SA"
NEW_MAP(19, 2) = "SB"
NEW_MAP(19, 3) = "SC"
NEW_MAP(19, 4) = "SD"
NEW_MAP(19, 5) = "SE"
NEW_MAP(19, 6) = "SF"
NEW_MAP(19, 7) = "SG"
NEW_MAP(19, 8) = "SH"
NEW_MAP(19, 9) = "SI"
NEW_MAP(19, 10) = "SJ"
NEW_MAP(19, 11) = "SK"
NEW_MAP(19, 12) = "SL"
NEW_MAP(19, 13) = "SM"
NEW_MAP(19, 14) = "SN"
NEW_MAP(19, 15) = "SO"
NEW_MAP(19, 16) = "SP"
NEW_MAP(19, 17) = "SQ"
NEW_MAP(19, 18) = "SR"
NEW_MAP(19, 19) = "SS"
NEW_MAP(19, 20) = "ST"
NEW_MAP(19, 21) = "SU"
NEW_MAP(19, 22) = "SV"
NEW_MAP(19, 23) = "SW"
NEW_MAP(19, 24) = "SX"
NEW_MAP(19, 25) = "SY"
NEW_MAP(19, 26) = "SZ"
NEW_MAP(20, 1) = "TA"
NEW_MAP(20, 2) = "TB"
NEW_MAP(20, 3) = "TC"
NEW_MAP(20, 4) = "TD"
NEW_MAP(20, 5) = "TE"
NEW_MAP(20, 6) = "TF"
NEW_MAP(20, 7) = "TG"
NEW_MAP(20, 8) = "TH"
NEW_MAP(20, 9) = "TI"
NEW_MAP(20, 10) = "TJ"
NEW_MAP(20, 11) = "TK"
NEW_MAP(20, 12) = "TL"
NEW_MAP(20, 13) = "TM"
NEW_MAP(20, 14) = "TN"
NEW_MAP(20, 15) = "TO"
NEW_MAP(20, 16) = "TP"
NEW_MAP(20, 17) = "TQ"
NEW_MAP(20, 18) = "TR"
NEW_MAP(20, 19) = "TS"
NEW_MAP(20, 20) = "TT"
NEW_MAP(20, 21) = "TU"
NEW_MAP(20, 22) = "TV"
NEW_MAP(20, 23) = "TW"
NEW_MAP(20, 24) = "TX"
NEW_MAP(20, 25) = "TY"
NEW_MAP(20, 26) = "TZ"
NEW_MAP(21, 1) = "UA"
NEW_MAP(21, 2) = "UB"
NEW_MAP(21, 3) = "UC"
NEW_MAP(21, 4) = "UD"
NEW_MAP(21, 5) = "UE"
NEW_MAP(21, 6) = "UF"
NEW_MAP(21, 7) = "UG"
NEW_MAP(21, 8) = "UH"
NEW_MAP(21, 9) = "UI"
NEW_MAP(21, 10) = "UJ"
NEW_MAP(21, 11) = "UK"
NEW_MAP(21, 12) = "UL"
NEW_MAP(21, 13) = "UM"
NEW_MAP(21, 14) = "UN"
NEW_MAP(21, 15) = "UO"
NEW_MAP(21, 16) = "UP"
NEW_MAP(21, 17) = "UQ"
NEW_MAP(21, 18) = "UR"
NEW_MAP(21, 19) = "US"
NEW_MAP(21, 20) = "UT"
NEW_MAP(21, 21) = "UU"
NEW_MAP(21, 22) = "UV"
NEW_MAP(21, 23) = "UW"
NEW_MAP(21, 24) = "UX"
NEW_MAP(21, 25) = "UY"
NEW_MAP(21, 26) = "UZ"
NEW_MAP(22, 1) = "VA"
NEW_MAP(22, 2) = "VB"
NEW_MAP(22, 3) = "VC"
NEW_MAP(22, 4) = "VD"
NEW_MAP(22, 5) = "VE"
NEW_MAP(22, 6) = "VF"
NEW_MAP(22, 7) = "VG"
NEW_MAP(22, 8) = "VH"
NEW_MAP(22, 9) = "VI"
NEW_MAP(22, 10) = "VJ"
NEW_MAP(22, 11) = "VK"
NEW_MAP(22, 12) = "VL"
NEW_MAP(22, 13) = "VM"
NEW_MAP(22, 14) = "VN"
NEW_MAP(22, 15) = "VO"
NEW_MAP(22, 16) = "VP"
NEW_MAP(22, 17) = "VQ"
NEW_MAP(22, 18) = "VR"
NEW_MAP(22, 19) = "VS"
NEW_MAP(22, 20) = "VT"
NEW_MAP(22, 21) = "VU"
NEW_MAP(22, 22) = "VV"
NEW_MAP(22, 23) = "VW"
NEW_MAP(22, 24) = "VX"
NEW_MAP(22, 25) = "VY"
NEW_MAP(22, 26) = "VZ"
NEW_MAP(23, 1) = "WA"
NEW_MAP(23, 2) = "WB"
NEW_MAP(23, 3) = "WC"
NEW_MAP(23, 4) = "WD"
NEW_MAP(23, 5) = "WE"
NEW_MAP(23, 6) = "WF"
NEW_MAP(23, 7) = "WG"
NEW_MAP(23, 8) = "WH"
NEW_MAP(23, 9) = "WI"
NEW_MAP(23, 10) = "WJ"
NEW_MAP(23, 11) = "WK"
NEW_MAP(23, 12) = "WL"
NEW_MAP(23, 13) = "WM"
NEW_MAP(23, 14) = "WN"
NEW_MAP(23, 15) = "WO"
NEW_MAP(23, 16) = "WP"
NEW_MAP(23, 17) = "WQ"
NEW_MAP(23, 18) = "WR"
NEW_MAP(23, 19) = "WS"
NEW_MAP(23, 20) = "WT"
NEW_MAP(23, 21) = "WU"
NEW_MAP(23, 22) = "WV"
NEW_MAP(23, 23) = "WW"
NEW_MAP(23, 24) = "WX"
NEW_MAP(23, 25) = "WY"
NEW_MAP(23, 26) = "WZ"
NEW_MAP(24, 1) = "XA"
NEW_MAP(24, 2) = "XB"
NEW_MAP(24, 3) = "XC"
NEW_MAP(24, 4) = "XD"
NEW_MAP(24, 5) = "XE"
NEW_MAP(24, 6) = "XF"
NEW_MAP(24, 7) = "XG"
NEW_MAP(24, 8) = "XH"
NEW_MAP(24, 9) = "XI"
NEW_MAP(24, 10) = "XJ"
NEW_MAP(24, 11) = "XK"
NEW_MAP(24, 12) = "XL"
NEW_MAP(24, 13) = "XM"
NEW_MAP(24, 14) = "XN"
NEW_MAP(24, 15) = "XO"
NEW_MAP(24, 16) = "XP"
NEW_MAP(24, 17) = "XQ"
NEW_MAP(24, 18) = "XR"
NEW_MAP(24, 19) = "XS"
NEW_MAP(24, 20) = "XT"
NEW_MAP(24, 21) = "XU"
NEW_MAP(24, 22) = "XV"
NEW_MAP(24, 23) = "XW"
NEW_MAP(24, 24) = "XX"
NEW_MAP(24, 25) = "XY"
NEW_MAP(24, 26) = "XZ"
NEW_MAP(25, 1) = "YA"
NEW_MAP(25, 2) = "YB"
NEW_MAP(25, 3) = "YC"
NEW_MAP(25, 4) = "YD"
NEW_MAP(25, 5) = "YE"
NEW_MAP(25, 6) = "YF"
NEW_MAP(25, 7) = "YG"
NEW_MAP(25, 8) = "YH"
NEW_MAP(25, 9) = "YI"
NEW_MAP(25, 10) = "YJ"
NEW_MAP(25, 11) = "YK"
NEW_MAP(25, 12) = "YL"
NEW_MAP(25, 13) = "YM"
NEW_MAP(25, 14) = "YN"
NEW_MAP(25, 15) = "YO"
NEW_MAP(25, 16) = "YP"
NEW_MAP(25, 17) = "YQ"
NEW_MAP(25, 18) = "YR"
NEW_MAP(25, 19) = "YS"
NEW_MAP(25, 20) = "YT"
NEW_MAP(25, 21) = "YU"
NEW_MAP(25, 22) = "YV"
NEW_MAP(25, 23) = "YW"
NEW_MAP(25, 24) = "YX"
NEW_MAP(25, 25) = "YY"
NEW_MAP(25, 26) = "YZ"
NEW_MAP(26, 1) = "ZA"
NEW_MAP(26, 2) = "ZB"
NEW_MAP(26, 3) = "ZC"
NEW_MAP(26, 4) = "ZD"
NEW_MAP(26, 5) = "ZE"
NEW_MAP(26, 6) = "ZF"
NEW_MAP(26, 7) = "ZG"
NEW_MAP(26, 8) = "ZH"
NEW_MAP(26, 9) = "ZI"
NEW_MAP(26, 10) = "ZJ"
NEW_MAP(26, 11) = "ZK"
NEW_MAP(26, 12) = "ZL"
NEW_MAP(26, 13) = "ZM"
NEW_MAP(26, 14) = "ZN"
NEW_MAP(26, 15) = "ZO"
NEW_MAP(26, 16) = "ZP"
NEW_MAP(26, 17) = "ZQ"
NEW_MAP(26, 18) = "ZR"
NEW_MAP(26, 19) = "ZS"
NEW_MAP(26, 20) = "ZT"
NEW_MAP(26, 21) = "ZU"
NEW_MAP(26, 22) = "ZV"
NEW_MAP(26, 23) = "ZW"
NEW_MAP(26, 24) = "ZX"
NEW_MAP(26, 25) = "ZY"
NEW_MAP(26, 26) = "ZZ"

DOWN = 1
ACROSS = 1
NUMBER = "0101"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"

Do
   HEXTABLE.AddNew
   HEXTABLE![MAP] = NEW_MAP(DOWN, ACROSS) & " " & NUMBER
   HEXTABLE.UPDATE

   If NUMBER = "3021" Then
      If DOWN = 26 Then
         If ACROSS = 26 Then
            Exit Do
         Else
            ACROSS = ACROSS + 1
            NUMBER = "0101"
         End If
      ElseIf ACROSS = 26 Then
         ACROSS = 1
         NUMBER = "0101"
         DOWN = DOWN + 1
      Else
         ACROSS = ACROSS + 1
         NUMBER = "0101"
      End If
   ElseIf Right(NUMBER, 2) = "21" Then
      If Left(NUMBER, 1) = "0" Then
         If Mid(NUMBER, 2, 1) = "9" Then
            NUMBER = Left(NUMBER, 2) + 1 & "01"
            If NUMBER = "01001" Then
               Msg = "1 - " & NUMBER
               Response = MsgBox(Msg, True)
            End If
         Else
            NUMBER = "0" & Mid(NUMBER, 2, 1) + 1 & "01"
            If NUMBER = "01001" Then
               Msg = "2 - " & NUMBER
               Response = MsgBox(Msg, True)
            End If
         End If
      Else
         NUMBER = Mid(NUMBER, 1, 2) + 1 & "01"
            If NUMBER = "01001" Then
               Msg = "3 - " & NUMBER
               Response = MsgBox(Msg, True)
            End If
      End If
   ElseIf Left(NUMBER, 1) = "0" Then
         If Mid(NUMBER, 2, 1) = "9" Then
            NUMBER = Left(NUMBER, 2) + 1 & "01"
            If NUMBER = "01001" Then
               Msg = "4 - " & NUMBER
               Response = MsgBox(Msg, True)
            End If
         ElseIf Mid(NUMBER, 3, 2) = "21" Then
            NUMBER = "0" & Mid(NUMBER, 2, 1) + 1 & "01"
            If NUMBER = "01001" Then
               Msg = "5 - " & NUMBER
               Response = MsgBox(Msg, True)
            End If
         ElseIf Mid(NUMBER, 3, 1) = "0" Then
            If Right(NUMBER, 1) = "9" Then
               NUMBER = "0" & Mid(NUMBER, 2, 1) & "10"
               If NUMBER = "01001" Then
                  Msg = "6 - " & NUMBER
                  Response = MsgBox(Msg, True)
               End If
            Else
               NUMBER = "0" & Mid(NUMBER, 2, 2) & Right(NUMBER, 1) + 1
               If NUMBER = "01001" Then
                  Msg = "7 - " & NUMBER
                  Response = MsgBox(Msg, True)
               End If
            End If
         ElseIf Mid(NUMBER, 4, 1) = "9" Then
            NUMBER = "0" & Mid(NUMBER, 2, 1) & Mid(NUMBER, 3, 1) + 1 & "0"
         Else
            NUMBER = "0" & Mid(NUMBER, 2, 2) & Right(NUMBER, 1) + 1
            If NUMBER = "01001" Then
               Msg = "8 - " & NUMBER
               Response = MsgBox(Msg, True)
            End If
         End If
   Else
      NUMBER = NUMBER + 1
      If NUMBER = "01001" Then
         Msg = "9 - " & NUMBER
         Response = MsgBox(Msg, True)
      End If
   End If
   
'   MSG1 = "NUMBER = " & NUMBER
'   MSG2 = "DOWN = " & DOWN
'   MSG3 = "ACROSS = " & ACROSS
'   MSG4 = "ERROR MSG = " & Err
'   RESPONSE = MsgBox(MSG1 & MSG2 & MSG3 & MSG4, True)

Loop

ERR_PAD_CLOSE:
   Exit Function

ERR_PAD:
If Err = 3022 Then
   Resume Next
Else
   MSG1 = "NUMBER = " & NUMBER
   MSG2 = "DOWN = " & DOWN
   MSG3 = "ACROSS = " & ACROSS
   MSG4 = "ERROR MSG = " & Err
   Response = MsgBox(MSG1 & MSG2 & MSG3 & MSG4, True)
   Resume ERR_PAD_CLOSE
End If

End Function

Function set_peter()
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GMTABLE.Delete
GMTABLE.AddNew
GMTABLE![Name] = "PETER"
GMTABLE![FILE] = "TVDATAPR.accdb"
GMTABLE![ODBC] = "TV_DATA_PR"
GMTABLE.UPDATE
GMTABLE.Close

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Call DELETE_ATTACHED_TABLES

Call REATTACH_TABLES

End Function

Function START_UP()

End Function

Function TRANSFER_SKILLS()

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "PRIMARYKEY"
SKILLSTABLE.MoveFirst

If Forms![TRANSFER SKILLS]![NEW SKILL LEVEL] = 0 Then
   SKILLSTABLE.Seek "=", Forms![TRANSFER SKILLS]![CURRENT TRIBE], Forms![TRANSFER SKILLS]![Skill]
   SKILLSTABLE.Delete
   SKILLSTABLE.AddNew
   SKILLSTABLE![TRIBE] = Forms![TRANSFER SKILLS]![New Tribe]
   SKILLSTABLE![Skill] = Forms![TRANSFER SKILLS]![Skill]
   SKILLSTABLE![SKILL LEVEL] = Forms![TRANSFER SKILLS]![SKILL LEVEL]
   SKILLSTABLE.UPDATE
Else
   SKILLSTABLE.Seek "=", Forms![TRANSFER SKILLS]![CURRENT TRIBE], Forms![TRANSFER SKILLS]![Skill]
   MSG1 = "Current Tribe = " & Forms![TRANSFER SKILLS]![CURRENT TRIBE]
   MSG2 = "Current Skill = " & Forms![TRANSFER SKILLS]![Skill]
   Response = MsgBox(MSG1 & MSG2, True)
   SKILLSTABLE.Edit
   SKILLSTABLE![SKILL LEVEL] = Forms![TRANSFER SKILLS]![SKILL LEVEL] - Forms![TRANSFER SKILLS]![NEW SKILL LEVEL]
   SKILLSTABLE.UPDATE
   SKILLSTABLE.AddNew
   SKILLSTABLE![TRIBE] = Forms![TRANSFER SKILLS]![New Tribe]
   SKILLSTABLE![Skill] = Forms![TRANSFER SKILLS]![Skill]
   SKILLSTABLE![SKILL LEVEL] = Forms![TRANSFER SKILLS]![NEW SKILL LEVEL]
   SKILLSTABLE.UPDATE
End If

TRIBE = Forms![TRANSFER SKILLS]![CURRENT TRIBE]

DoCmd.Close A_FORM, "TRANSFER SKILLS"
DoCmd.OpenForm "TRANSFER SKILLS"
DoCmd.FindRecord TRIBE

End Function

Function UPDATE_FRESH_WATER()
Dim CurrentHex As String
Dim HEX_N As String, HEX_NE As String, HEX_SE As String
Dim HEX_S As String, HEX_SW As String, HEX_NW As String
Dim AMT_OF_WATER As Long

AMT_OF_WATER = 0

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst

Do Until HEXTABLE.EOF
   CurrentHex = HEXTABLE![MAP]
   HEX_N = GET_MAP_NORTH(CurrentHex)
   HEX_NE = GET_MAP_NORTH_EAST(CurrentHex)
   HEX_SE = GET_MAP_SOUTH_EAST(CurrentHex)
   HEX_S = GET_MAP_SOUTH(CurrentHex)
   HEX_SW = GET_MAP_SOUTH_WEST(CurrentHex)
   HEX_NW = GET_MAP_NORTH_WEST(CurrentHex)
   HEXTABLE.Edit
   If Mid(HEXTABLE![RIVERS], 1, 1) = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   If Mid(HEXTABLE![RIVERS], 2, 1) = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   If Mid(HEXTABLE![RIVERS], 3, 1) = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   If Mid(HEXTABLE![RIVERS], 4, 1) = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   If Mid(HEXTABLE![RIVERS], 5, 1) = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   If Mid(HEXTABLE![RIVERS], 6, 1) = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   If HEXTABLE![SPRINGS] = "Y" Then
      AMT_OF_WATER = AMT_OF_WATER + 1
   End If
   
   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_N
   If Not HEXTABLE.NoMatch Then
      If HEXTABLE![TERRAIN] = "OCEAN" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
      If HEXTABLE![TERRAIN] = "LAKE" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_NE
   If Not HEXTABLE.NoMatch Then
      If HEXTABLE![TERRAIN] = "OCEAN" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
      If HEXTABLE![TERRAIN] = "LAKE" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_SE
   If Not HEXTABLE.NoMatch Then
      If HEXTABLE![TERRAIN] = "OCEAN" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
      If HEXTABLE![TERRAIN] = "LAKE" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_S
   If Not HEXTABLE.NoMatch Then
      If HEXTABLE![TERRAIN] = "OCEAN" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
      If HEXTABLE![TERRAIN] = "LAKE" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_SW
   If Not HEXTABLE.NoMatch Then
      If HEXTABLE![TERRAIN] = "OCEAN" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
      If HEXTABLE![TERRAIN] = "LAKE" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_NW
   If Not HEXTABLE.NoMatch Then
      If HEXTABLE![TERRAIN] = "OCEAN" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
      If HEXTABLE![TERRAIN] = "LAKE" Then
         AMT_OF_WATER = AMT_OF_WATER + 1
      End If
   End If

   HEXTABLE.Seek "=", CurrentHex
   HEXTABLE.Edit
   If AMT_OF_WATER = 0 Then
      HEXTABLE![FRESH WATER] = 0
   ElseIf AMT_OF_WATER = 1 Then
      HEXTABLE![FRESH WATER] = 1.1
   ElseIf AMT_OF_WATER = 2 Then
      HEXTABLE![FRESH WATER] = 1.11
   ElseIf AMT_OF_WATER = 3 Then
      HEXTABLE![FRESH WATER] = 1.12
   ElseIf AMT_OF_WATER = 4 Then
      HEXTABLE![FRESH WATER] = 1.13
   ElseIf AMT_OF_WATER = 5 Then
      HEXTABLE![FRESH WATER] = 1.14
   ElseIf AMT_OF_WATER = 6 Then
      HEXTABLE![FRESH WATER] = 1.15
   ElseIf AMT_OF_WATER > 6 Then
      HEXTABLE![FRESH WATER] = 1.16
   End If
   HEXTABLE.UPDATE
   AMT_OF_WATER = 0
   HEXTABLE.MoveNext
Loop

HEXTABLE.Close

End Function

Function UPDATE_HEX()
Dim CurrentHex As String
Dim HEX_N As String, HEX_NE As String, HEX_SE As String
Dim HEX_S As String, HEX_SW As String, HEX_NW As String
Dim RIVER_N As String, RIVER_NE As String, RIVER_SE As String
Dim RIVER_S As String, RIVER_SW As String, RIVER_NW As String
Dim CLIFF_N As String, CLIFF_NE As String, CLIFF_SE As String
Dim CLIFF_S As String, CLIFF_SW As String, CLIFF_NW As String
Dim BEACH_N As String, BEACH_NE As String, BEACH_SE As String
Dim BEACH_S As String, BEACH_SW As String, BEACH_NW As String
Dim ROAD_N As String, ROAD_NE As String, ROAD_SE As String
Dim ROAD_S As String, ROAD_SW As String, ROAD_NW As String
Dim WATERFALL_N As String, WATERFALL_NE As String, WATERFALL_SE As String
Dim WATERFALL_S As String, WATERFALL_SW As String, WATERFALL_NW As String
Dim CANYON_N As String, CANYON_NE As String, CANYON_SE As String
Dim CANYON_S As String, CANYON_SW As String, CANYON_NW As String
Dim STREAM_N As String, STREAM_NE As String, STREAM_SE As String
Dim STREAM_S As String, STREAM_SW As String, STREAM_NW As String

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst

Do Until HEXTABLE.EOF
   HEXTABLE.Edit
   
   CurrentHex = HEXTABLE![MAP]
   RIVER_N = Mid(HEXTABLE![RIVERS], 1, 1)
   RIVER_NE = Mid(HEXTABLE![RIVERS], 2, 1)
   RIVER_SE = Mid(HEXTABLE![RIVERS], 3, 1)
   RIVER_S = Mid(HEXTABLE![RIVERS], 4, 1)
   RIVER_SW = Mid(HEXTABLE![RIVERS], 5, 1)
   RIVER_NW = Mid(HEXTABLE![RIVERS], 6, 1)
   FORD_N = Mid(HEXTABLE![FORDS], 1, 1)
   FORD_NE = Mid(HEXTABLE![FORDS], 2, 1)
   FORD_SE = Mid(HEXTABLE![FORDS], 3, 1)
   FORD_S = Mid(HEXTABLE![FORDS], 4, 1)
   FORD_SW = Mid(HEXTABLE![FORDS], 5, 1)
   FORD_NW = Mid(HEXTABLE![FORDS], 6, 1)
   PASS_N = Mid(HEXTABLE![PASSES], 1, 1)
   PASS_NE = Mid(HEXTABLE![PASSES], 2, 1)
   PASS_SE = Mid(HEXTABLE![PASSES], 3, 1)
   PASS_S = Mid(HEXTABLE![PASSES], 4, 1)
   PASS_SW = Mid(HEXTABLE![PASSES], 5, 1)
   PASS_NW = Mid(HEXTABLE![PASSES], 6, 1)
   CLIFF_N = Mid(HEXTABLE![CLIFFS], 1, 1)
   CLIFF_NE = Mid(HEXTABLE![CLIFFS], 2, 1)
   CLIFF_SE = Mid(HEXTABLE![CLIFFS], 3, 1)
   CLIFF_S = Mid(HEXTABLE![CLIFFS], 4, 1)
   CLIFF_SW = Mid(HEXTABLE![CLIFFS], 5, 1)
   CLIFF_NW = Mid(HEXTABLE![CLIFFS], 6, 1)
   BEACH_N = Mid(HEXTABLE![BEACHES], 1, 1)
   BEACH_NE = Mid(HEXTABLE![BEACHES], 2, 1)
   BEACH_SE = Mid(HEXTABLE![BEACHES], 3, 1)
   BEACH_S = Mid(HEXTABLE![BEACHES], 4, 1)
   BEACH_SW = Mid(HEXTABLE![BEACHES], 5, 1)
   BEACH_NW = Mid(HEXTABLE![BEACHES], 6, 1)
   ROAD_N = Mid(HEXTABLE![ROADS], 1, 1)
   ROAD_NE = Mid(HEXTABLE![ROADS], 2, 1)
   ROAD_SE = Mid(HEXTABLE![ROADS], 3, 1)
   ROAD_S = Mid(HEXTABLE![ROADS], 4, 1)
   ROAD_SW = Mid(HEXTABLE![ROADS], 5, 1)
   ROAD_NW = Mid(HEXTABLE![ROADS], 6, 1)
   WATERFALL_N = Mid(HEXTABLE![WATERFALLS], 1, 1)
   WATERFALL_NE = Mid(HEXTABLE![WATERFALLS], 2, 1)
   WATERFALL_SE = Mid(HEXTABLE![WATERFALLS], 3, 1)
   WATERFALL_S = Mid(HEXTABLE![WATERFALLS], 4, 1)
   WATERFALL_SW = Mid(HEXTABLE![WATERFALLS], 5, 1)
   WATERFALL_NW = Mid(HEXTABLE![WATERFALLS], 6, 1)
   CANYON_N = Mid(HEXTABLE![CANYONS], 1, 1)
   CANYON_NE = Mid(HEXTABLE![CANYONS], 2, 1)
   CANYON_SE = Mid(HEXTABLE![CANYONS], 3, 1)
   CANYON_S = Mid(HEXTABLE![CANYONS], 4, 1)
   CANYON_SW = Mid(HEXTABLE![CANYONS], 5, 1)
   CANYON_NW = Mid(HEXTABLE![CANYONS], 6, 1)
   STREAM_N = Mid(HEXTABLE![STREAMS], 1, 1)
   STREAM_NE = Mid(HEXTABLE![STREAMS], 2, 1)
   STREAM_SE = Mid(HEXTABLE![STREAMS], 3, 1)
   STREAM_S = Mid(HEXTABLE![STREAMS], 4, 1)
   STREAM_SW = Mid(HEXTABLE![STREAMS], 5, 1)
   STREAM_NW = Mid(HEXTABLE![STREAMS], 6, 1)

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_N
   If Not HEXTABLE.NoMatch Then
      HEXTABLE.Edit
      HEXTABLE![RIVERS] = Mid(HEXTABLE![RIVERS], 1, 3) & RIVER_N & Mid(HEXTABLE![RIVERS], 5, 2)
      HEXTABLE![FORDS] = Mid(HEXTABLE![FORDS], 1, 3) & FORD_N & Mid(HEXTABLE![FORDS], 5, 2)
      HEXTABLE![PASSES] = Mid(HEXTABLE![PASSES], 1, 3) & PASS_N & Mid(HEXTABLE![PASSES], 5, 2)
      HEXTABLE![CLIFFS] = Mid(HEXTABLE![CLIFFS], 1, 3) & CLIFF_N & Mid(HEXTABLE![CLIFFS], 5, 2)
      HEXTABLE![BEACHES] = Mid(HEXTABLE![BEACHES], 1, 3) & BEACH_N & Mid(HEXTABLE![BEACHES], 5, 2)
      HEXTABLE![ROADS] = Mid(HEXTABLE![ROADS], 1, 3) & ROAD_N & Mid(HEXTABLE![ROADS], 5, 2)
      HEXTABLE![WATERFALLS] = Mid(HEXTABLE![WATERFALLS], 1, 3) & WATERFALL_N & Mid(HEXTABLE![WATERFALLS], 5, 2)
      HEXTABLE![CANYONS] = Mid(HEXTABLE![CANYONS], 1, 3) & CANYON_N & Mid(HEXTABLE![CANYONS], 5, 2)
      HEXTABLE![STREAMS] = Mid(HEXTABLE![STREAMS], 1, 3) & STREAM_N & Mid(HEXTABLE![STREAMS], 5, 2)
      HEXTABLE.UPDATE
   End If
   
   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_NE
   If Not HEXTABLE.NoMatch Then
      HEXTABLE.Edit
      HEXTABLE![RIVERS] = Mid(HEXTABLE![RIVERS], 1, 4) & RIVER_NE & Mid(HEXTABLE![RIVERS], 6, 1)
      HEXTABLE![FORDS] = Mid(HEXTABLE![FORDS], 1, 4) & FORD_NE & Mid(HEXTABLE![FORDS], 6, 1)
      HEXTABLE![PASSES] = Mid(HEXTABLE![PASSES], 1, 4) & PASS_NE & Mid(HEXTABLE![PASSES], 6, 1)
      HEXTABLE![CLIFFS] = Mid(HEXTABLE![CLIFFS], 1, 4) & CLIFF_NE & Mid(HEXTABLE![CLIFFS], 6, 1)
      HEXTABLE![BEACHES] = Mid(HEXTABLE![BEACHES], 1, 4) & BEACH_NE & Mid(HEXTABLE![BEACHES], 6, 1)
      HEXTABLE![ROADS] = Mid(HEXTABLE![ROADS], 1, 4) & ROAD_NE & Mid(HEXTABLE![ROADS], 6, 1)
      HEXTABLE![WATERFALLS] = Mid(HEXTABLE![WATERFALLS], 1, 4) & WATERFALL_NE & Mid(HEXTABLE![WATERFALLS], 6, 1)
      HEXTABLE![CANYONS] = Mid(HEXTABLE![CANYONS], 1, 4) & CANYON_NE & Mid(HEXTABLE![CANYONS], 6, 1)
      HEXTABLE![STREAMS] = Mid(HEXTABLE![STREAMS], 1, 4) & STREAM_NE & Mid(HEXTABLE![STREAMS], 6, 1)
      HEXTABLE.UPDATE
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_SE
   If Not HEXTABLE.NoMatch Then
      HEXTABLE.Edit
      HEXTABLE![RIVERS] = Mid(HEXTABLE![RIVERS], 1, 5) & RIVER_SE
      HEXTABLE![FORDS] = Mid(HEXTABLE![FORDS], 1, 5) & FORD_SE
      HEXTABLE![PASSES] = Mid(HEXTABLE![PASSES], 1, 5) & PASS_SE
      HEXTABLE![CLIFFS] = Mid(HEXTABLE![CLIFFS], 1, 5) & CLIFF_SE
      HEXTABLE![BEACHES] = Mid(HEXTABLE![BEACHES], 1, 5) & BEACH_SE
      HEXTABLE![ROADS] = Mid(HEXTABLE![ROADS], 1, 5) & ROAD_SE
      HEXTABLE![WATERFALLS] = Mid(HEXTABLE![WATERFALLS], 1, 5) & WATERFALL_SE
      HEXTABLE![CANYONS] = Mid(HEXTABLE![CANYONS], 1, 5) & CANYON_SE
      HEXTABLE![STREAMS] = Mid(HEXTABLE![STREAMS], 1, 5) & STREAM_SE
      HEXTABLE.UPDATE
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_S
   If Not HEXTABLE.NoMatch Then
      HEXTABLE.Edit
      HEXTABLE![RIVERS] = RIVER_S & Mid(HEXTABLE![RIVERS], 2, 5)
      HEXTABLE![FORDS] = FORD_S & Mid(HEXTABLE![FORDS], 2, 5)
      HEXTABLE![PASSES] = PASS_S & Mid(HEXTABLE![PASSES], 2, 5)
      HEXTABLE![CLIFFS] = CLIFF_S & Mid(HEXTABLE![CLIFFS], 2, 5)
      HEXTABLE![BEACHES] = BEACH_S & Mid(HEXTABLE![BEACHES], 2, 5)
      HEXTABLE![ROADS] = ROAD_S & Mid(HEXTABLE![ROADS], 2, 5)
      HEXTABLE![WATERFALLS] = WATERFALL_S & Mid(HEXTABLE![WATERFALLS], 2, 5)
      HEXTABLE![CANYONS] = CANYON_S & Mid(HEXTABLE![CANYONS], 2, 5)
      HEXTABLE![STREAMS] = STREAM_S & Mid(HEXTABLE![STREAMS], 2, 5)
      HEXTABLE.UPDATE
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_SW
   If Not HEXTABLE.NoMatch Then
      HEXTABLE.Edit
      HEXTABLE![RIVERS] = Mid(HEXTABLE![RIVERS], 1, 1) & RIVER_SW & Mid(HEXTABLE![RIVERS], 3, 4)
      HEXTABLE![FORDS] = Mid(HEXTABLE![FORDS], 1, 1) & FORD_SW & Mid(HEXTABLE![FORDS], 3, 4)
      HEXTABLE![PASSES] = Mid(HEXTABLE![PASSES], 1, 1) & PASS_SW & Mid(HEXTABLE![PASSES], 3, 4)
      HEXTABLE![CLIFFS] = Mid(HEXTABLE![CLIFFS], 1, 1) & CLIFF_SW & Mid(HEXTABLE![CLIFFS], 3, 4)
      HEXTABLE![BEACHES] = Mid(HEXTABLE![BEACHES], 1, 1) & BEACH_SW & Mid(HEXTABLE![BEACHES], 3, 4)
      HEXTABLE![ROADS] = Mid(HEXTABLE![ROADS], 1, 1) & ROAD_SW & Mid(HEXTABLE![ROADS], 3, 4)
      HEXTABLE![WATERFALLS] = Mid(HEXTABLE![WATERFALLS], 1, 1) & WATERFALL_SW & Mid(HEXTABLE![WATERFALLS], 3, 4)
      HEXTABLE![CANYONS] = Mid(HEXTABLE![CANYONS], 1, 1) & CANYON_SW & Mid(HEXTABLE![CANYONS], 3, 4)
      HEXTABLE![STREAMS] = Mid(HEXTABLE![STREAMS], 1, 1) & STREAM_SW & Mid(HEXTABLE![STREAMS], 3, 4)
      HEXTABLE.UPDATE
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", HEX_NW
   If Not HEXTABLE.NoMatch Then
      HEXTABLE.Edit
      HEXTABLE![RIVERS] = Mid(HEXTABLE![RIVERS], 1, 2) & RIVER_NW & Mid(HEXTABLE![RIVERS], 4, 3)
      HEXTABLE![FORDS] = Mid(HEXTABLE![FORDS], 1, 2) & FORD_NW & Mid(HEXTABLE![FORDS], 4, 3)
      HEXTABLE![PASSES] = Mid(HEXTABLE![PASSES], 1, 2) & PASS_NW & Mid(HEXTABLE![PASSES], 4, 3)
      HEXTABLE![CLIFFS] = Mid(HEXTABLE![CLIFFS], 1, 2) & CLIFF_NW & Mid(HEXTABLE![CLIFFS], 4, 3)
      HEXTABLE![BEACHES] = Mid(HEXTABLE![BEACHES], 1, 2) & BEACH_NW & Mid(HEXTABLE![BEACHES], 4, 3)
      HEXTABLE![ROADS] = Mid(HEXTABLE![ROADS], 1, 2) & ROAD_NW & Mid(HEXTABLE![ROADS], 4, 3)
      HEXTABLE![WATERFALLS] = Mid(HEXTABLE![WATERFALLS], 1, 2) & WATERFALL_NW & Mid(HEXTABLE![WATERFALLS], 4, 3)
      HEXTABLE![CANYONS] = Mid(HEXTABLE![CANYONS], 1, 2) & CANYON_NW & Mid(HEXTABLE![CANYONS], 4, 3)
      HEXTABLE![STREAMS] = Mid(HEXTABLE![STREAMS], 1, 2) & STREAM_NW & Mid(HEXTABLE![STREAMS], 4, 3)
      HEXTABLE.UPDATE
   End If

   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", CurrentHex
   HEXTABLE.MoveNext
Loop

HEXTABLE.Close

End Function

Function UPDATE_MINERAL_AMTS()

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_minerals")
HEXMAPMINERALS.MoveFirst

Do Until HEXMAPMINERALS.EOF
   HEXMAPMINERALS.Edit

'  INCLUDE CHECKING FOR MINERALS IN ALL THREE PLACES - (ORE TYPE) (SECOND ORE) (THIRD ORE)
   If HEXMAPMINERALS![ORE_TYPE] = "COAL" Then
      HEXMAPMINERALS![MINING] = 6
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "COPPER ORE" Then
      HEXMAPMINERALS![MINING] = 4
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "GEMS" Then
      HEXMAPMINERALS![MINING] = 0.5
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "GOLD" Then
      HEXMAPMINERALS![MINING] = 1
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "IRON ORE" Then
      HEXMAPMINERALS![MINING] = 3
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "LEAD ORE" Then
      HEXMAPMINERALS![MINING] = 3.5
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "SALT" Then
      HEXMAPMINERALS![MINING] = 4
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "SILVER" Then
      HEXMAPMINERALS![MINING] = 10
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "TIN ORE" Then
      HEXMAPMINERALS![MINING] = 3.5
   ElseIf HEXMAPMINERALS![ORE_TYPE] = "ZINC ORE" Then
      HEXMAPMINERALS![MINING] = 3
   End If
   
   If HEXMAPMINERALS![SECOND_ORE] = "COAL" Then
      HEXMAPMINERALS![SECOND_MINING] = 6
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "COPPER ORE" Then
      HEXMAPMINERALS![SECOND_MINING] = 4
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "GEMS" Then
      HEXMAPMINERALS![SECOND_MINING] = 0.5
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "GOLD" Then
      HEXMAPMINERALS![SECOND_MINING] = 1
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "IRON ORE" Then
      HEXMAPMINERALS![SECOND_MINING] = 3
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "LEAD ORE" Then
      HEXMAPMINERALS![SECOND_MINING] = 3.5
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "SALT" Then
      HEXMAPMINERALS![SECOND_MINING] = 4
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "SILVER" Then
      HEXMAPMINERALS![SECOND_MINING] = 10
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "TIN ORE" Then
      HEXMAPMINERALS![SECOND_MINING] = 3.5
   ElseIf HEXMAPMINERALS![SECOND_ORE] = "ZINC ORE" Then
      HEXMAPMINERALS![SECOND_MINING] = 3
   End If
   
   If HEXMAPMINERALS![THIRD_ORE] = "COAL" Then
      HEXMAPMINERALS![THIRD_MINING] = 6
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "COPPER ORE" Then
      HEXMAPMINERALS![THIRD_MINING] = 4
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "GEMS" Then
      HEXMAPMINERALS![THIRD_MINING] = 0.5
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "GOLD" Then
      HEXMAPMINERALS![THIRD_MINING] = 1
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "IRON ORE" Then
      HEXMAPMINERALS![THIRD_MINING] = 3
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "LEAD ORE" Then
      HEXMAPMINERALS![THIRD_MINING] = 3.5
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "SALT" Then
      HEXMAPMINERALS![THIRD_MINING] = 4
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "SILVER" Then
      HEXMAPMINERALS![THIRD_MINING] = 10
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "TIN ORE" Then
      HEXMAPMINERALS![THIRD_MINING] = 3.5
   ElseIf HEXMAPMINERALS![THIRD_ORE] = "ZINC ORE" Then
      HEXMAPMINERALS![THIRD_MINING] = 3
   End If

   HEXMAPMINERALS.UPDATE

   HEXMAPMINERALS.MoveNext
Loop

End Function

Function WHO_IS_IN_HEX(CLAN, TRIBE, CurrentHex, MOVE_TRIBE)
Dim CURRENT_HEX As String
Dim TRIBES_IN_HEX As String

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBEINFO = TVDBGM.OpenRecordset("TRIBES_General_Info")
TRIBEINFO.MoveFirst
TRIBEINFO.index = "MAP"
TRIBEINFO.Seek "=", CurrentHex

If IsNull(CurrentHex) Then
   TRIBEINFO.MoveFirst
   TRIBEINFO.index = "PRIMARYKEY"
   TRIBEINFO.Seek "=", CLAN, TRIBE
   CURRENT_HEX = TRIBEINFO![CURRENT HEX]
   TRIBEINFO.MoveFirst
   TRIBEINFO.index = "MAP"
   TRIBEINFO.Seek "=", CurrentHex

Else
   CURRENT_HEX = CurrentHex
End If

TRIBES_IN_HEX = "EMPTY"

If Not TRIBEINFO.NoMatch Then
Do

   If MOVE_TRIBE = "N" Then
      If Not TRIBEINFO![TRIBE] = TRIBE Then
         If TRIBES_IN_HEX = "EMPTY" Then
            TRIBES_IN_HEX = TRIBEINFO![TRIBE]
         Else
            TRIBES_IN_HEX = TRIBES_IN_HEX & ", " & TRIBEINFO![TRIBE]
         End If
      End If
   ElseIf TRIBES_IN_HEX = "EMPTY" Then
      TRIBES_IN_HEX = TRIBEINFO![TRIBE]
   Else
      TRIBES_IN_HEX = TRIBES_IN_HEX & ", " & TRIBEINFO![TRIBE]
   End If

   TRIBEINFO.MoveNext

   If TRIBEINFO.EOF Then
      Exit Do
   End If
   If Not TRIBEINFO![CURRENT HEX] = CURRENT_HEX Then
      Exit Do
   End If
Loop
End If

TRIBEINFO.Close

Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst
HEXTABLE.Seek "=", CURRENT_HEX

If HEXTABLE.NoMatch Then
   'ignore
ElseIf HEXTABLE![ROAMING HERD] = "Y" Then
   TRIBES_IN_HEX = TRIBES_IN_HEX & ", ROAMING HERD"
End If

If TRIBES_IN_HEX = "EMPTY" Then
   WHO_IS_IN_HEX = "EMPTY"
Else
   WHO_IS_IN_HEX = TRIBES_IN_HEX
End If
End Function

Function WRITE_TURN_ACTIVITY(CLAN, TRIBE, Section_Ident, LINENUMBER, OutLine, Additional)
  
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set Turns_Activities_Table = TVDBGM.OpenRecordset("Turns_Activities")
Turns_Activities_Table.index = "PRIMARYKEY"
Turns_Activities_Table.Seek "=", CLAN, TRIBE, Section_Ident, LINENUMBER

If Turns_Activities_Table.NoMatch Then
   Turns_Activities_Table.AddNew
   Turns_Activities_Table![CLAN] = CLAN
   Turns_Activities_Table![TRIBE] = TRIBE
   Turns_Activities_Table![Section] = Section_Ident
   Turns_Activities_Table![LINE NUMBER] = LINENUMBER
   Turns_Activities_Table![line detail] = OutLine

   Turns_Activities_Table.UPDATE
Else
   If Additional = "No" Then
      Turns_Activities_Table.Edit
      Turns_Activities_Table![line detail] = OutLine
      Turns_Activities_Table.UPDATE
   Else
      Turns_Activities_Table.Edit
      Turns_Activities_Table![line detail] = Turns_Activities_Table![line detail] & OutLine
      Turns_Activities_Table.UPDATE
   End If
   
End If

Turns_Activities_Table.Close

End Function
Public Function SET_GM(GM_NAME, SCREEN)
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

If SCREEN = "YES" Then
   Forms![WHICH GM ARE YOU]![Status] = "Set GM " + GM_NAME + " starts"
   Forms![WHICH GM ARE YOU].Repaint
End If

' GET THE CURRENT DIRECTORY
  CURRENT_DIRECTORY = CurDir$
   
If GM_NAME = "JEFF" Then
   GMTABLE.Delete
   GMTABLE.AddNew
   GMTABLE![Name] = "JEFF"
   GMTABLE![FILE] = "TVDATAJF.accdb"
   GMTABLE![ODBC] = "TV_DATA_JF"
   GMTABLE![DIRECTORY] = CURRENT_DIRECTORY
   GMTABLE![cALc_costs_PROCESSED] = "N"
   GMTABLE![FINAL_ACTIVITIES_PROCESSED] = "N"
   GMTABLE.UPDATE
   GMTABLE.Close
   FILEGM = CurDir$ & "\TVDATAJF.accdb"

ElseIf GM_NAME = "PETER" Then
   GMTABLE.Delete
   GMTABLE.AddNew
   GMTABLE![Name] = "PETER"
   GMTABLE![FILE] = "TVDATAPR.accdb"
   GMTABLE![ODBC] = "TV_DATA_PR"
   GMTABLE![DIRECTORY] = CURRENT_DIRECTORY
   GMTABLE![cALc_costs_PROCESSED] = "N"
   GMTABLE![FINAL_ACTIVITIES_PROCESSED] = "N"
   GMTABLE.UPDATE
   GMTABLE.Close
   FILEGM = CurDir$ & "\TVDATAPR.accdb"

End If

Call DELETE_ATTACHED_TABLES

Call REATTACH_TABLES

If SCREEN = "YES" Then
   Forms![WHICH GM ARE YOU]![Status] = "Set GM " + GM_NAME + " ends"
   Forms![WHICH GM ARE YOU].Repaint
End If

End Function

Public Function CALC_CLAN_RATINGS()

   Call CLAN_STATISTICS("NO")
   DoCmd.OpenQuery "STATS - GOODS STATS"
   Call CALC_CLAN_RATING
   DoCmd.Hourglass False

End Function

Public Function Open_Table(TABLE_NAME As String)

DoCmd.OpenTable (TABLE_NAME)

End Function

Public Function Check_GL_Level_against_Completed_Research()

On Error GoTo ERR_GL_CHECK

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.MoveFirst

Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTAB.index = "primarykey"
COMPRESTAB.MoveFirst

Do Until TRIBESINFO.EOF
   Forms![TRIBEVIBES]![Status] = "Fix Government Level" + TRIBESINFO![TRIBE]
   Forms![TRIBEVIBES].Repaint
   
   COMPRESTAB.Seek "=", TRIBESINFO![TRIBE], "GOVERNMENT LEVEL 1"
   If Not COMPRESTAB.NoMatch Then
      COMPRESTAB.Seek "=", TRIBESINFO![TRIBE], "GOVERNMENT LEVEL 2"
      If Not COMPRESTAB.NoMatch Then
         COMPRESTAB.Seek "=", TRIBESINFO![TRIBE], "GOVERNMENT LEVEL 3"
         If Not COMPRESTAB.NoMatch Then
            COMPRESTAB.Seek "=", TRIBESINFO![TRIBE], "GOVERNMENT LEVEL 4"
            If Not COMPRESTAB.NoMatch Then
               COMPRESTAB.Seek "=", TRIBESINFO![TRIBE], "GOVERNMENT LEVEL 5"
               If Not COMPRESTAB.NoMatch Then
                  TRIBESINFO.Edit
                  TRIBESINFO![GOVT LEVEL] = 5
                  TRIBESINFO.UPDATE
               Else
                  TRIBESINFO.Edit
                  TRIBESINFO![GOVT LEVEL] = 4
                  TRIBESINFO.UPDATE
               End If
            Else
               TRIBESINFO.Edit
               TRIBESINFO![GOVT LEVEL] = 3
               TRIBESINFO.UPDATE
            End If
         Else
            TRIBESINFO.Edit
            TRIBESINFO![GOVT LEVEL] = 2
            TRIBESINFO.UPDATE
         End If
      Else
         TRIBESINFO.Edit
         TRIBESINFO![GOVT LEVEL] = 1
         TRIBESINFO.UPDATE
      End If
   End If
   TRIBESINFO.MoveNext
Loop

Forms![TRIBEVIBES]![Status] = " "
Forms![TRIBEVIBES].Repaint
   
ERR_GL_CHECK_CLOSE:
   Exit Function

ERR_GL_CHECK:
If (Err = 3022) Then
   Resume Next
Else
   MSG1 = "ERROR = " & Err
   MSG2 = "HEXNUMBER = " & HEXNUMBER
   Response = MsgBox(Msg & MSG1 & MSG2, True)
   Resume ERR_GL_CHECK_CLOSE
End If

End Function

Public Function REATTACH_TABLES()
On Error GoTo ERR_REATTACH_TAB
'======================================
' Commented out as Split database can reattach ALL tables
' andrew.d.bentley@gmail.com
'======================================
'DebugOP "f - REATTACH_TABLES()"
'
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "CLAN_STATS", "CLAN_STATS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "COMPLETED_RESEARCH", "COMPLETED_RESEARCH"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "DICE_ROLLS", "DICE_ROLLS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GAMES_WEATHER", "GAMES_WEATHER"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GLOBAL", "GLOBAL"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GOODS_TRIBES_PROCESSING", "GOODS_TRIBES_PROCESSING"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GOODS_STATS", "GOODS_STATS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GM_COSTS_TABLE", "GM_COSTS_TABLE"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HERD_SWAPS", "HERD_SWAPS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEX_MAP", "HEX_MAP"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEX_MAP_CITY", "HEX_MAP_CITY"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEX_MAP_CONST", "HEX_MAP_CONST"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEX_MAP_MINERALS", "HEX_MAP_MINERALS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEX_MAP_POLITICS", "HEX_MAP_POLITICS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEXMAP_FARMING", "HEXMAP_FARMING"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEXMAP_PERMANENT_FARMING", "HEXMAP_PERMANENT_FARMING"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "MASSTRANSFERS", "MASSTRANSFERS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "MODIFIERS", "MODIFIERS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "PACIFICATION_TABLE", "PACIFICATION_TABLE"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "PERMANENT_MESSAGES_Table", "PERMANENT_MESSAGES_Table"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "POPULATION_INCREASE", "POPULATION_INCREASE"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Scout_Movement", "Process_Scout_Movement"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Research", "Process_Research"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Skills", "Process_Skills"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribe_Movement", "Process_Tribe_Movement"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Activity", "Process_Tribes_Activity"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Activity_Copy", "Process_Tribes_Activity_Copy"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Item_Allocation", "Process_Tribes_Item_Allocation"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Item_Allocation_Copy", "Process_Tribes_Item_Allocation_Copy"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Transfers", "Process_TRIBES_TRANSFERS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "PROVS_AVAILABILITY", "PROVS_AVAILABILITY"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "RESEARCH_ATTEMPTS", "RESEARCH_ATTEMPTS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Scout_Movement", "Scout_Movement"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "SEEKING_RETURNS_TABLE", "SEEKING_RETURNS_TABLE"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "SHIP_DAMAGE", "SHIP_DAMAGE"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "SKILL_ATTEMPTS", "SKILL_ATTEMPTS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "SKILLS", "SKILLS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "SKILLS_STATS", "SKILLS_STATS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "SPECIAL_TRANSFER_ROUTES", "SPECIAL_TRANSFER_ROUTES"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TEMP_TRADING_POST", "TEMP_TRADING_POST"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TERRAIN_COMBAT", "TERRAIN_COMBAT"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRADING_POST_GOODS", "TRADING_POST_GOODS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBE_CHECKING", "TRIBE_CHECKING"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBE_RESEARCH", "TRIBE_RESEARCH"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBES_PROCESSING", "TRIBES_PROCESSING"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBES_BOOKS", "TRIBES_BOOKS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBES_general_info", "TRIBES_general_info"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBES_GOODS", "TRIBES_GOODS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBES_SPECIALISTS", "TRIBES_SPECIALISTS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TRIBES_TURNS_ACTIVITIES", "TRIBES_TURNS_ACTIVITIES"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TURNS_ACTIVITIES", "TURNS_ACTIVITIES"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TURN_INFO_REQD_NEXT_TURN", "TURN_INFO_REQD_NEXT_TURN"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "TURNS_trading_post_activity", "TURNS_trading_post_activity"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "UNDER_CONSTRUCTION", "UNDER_CONSTRUCTION"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "VALID_GOODS", "VALID_GOODS"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "WEAPON_ARMOUR", "WEAPON_ARMOUR"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "WEATHER", "WEATHER"
'DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "WEATHER_COMBAT", "WEATHER_COMBAT"

ERR_REATTACH_TAB_CLOSE:
   Exit Function

ERR_REATTACH_TAB:
   Resume Next


End Function

Public Function TV_accdb_CLEAN_UP()
Dim QUERY As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM AVAILABLE_RESEARCH;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM CLAN_RATINGS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM IMPLEMENT_USAGE;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM PRINTING_SWITCHS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM SEEKING_RETURNS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TRIBES_ACTIVITY_PHASE;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TRIBES_GOODS_USAGE;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM UNDER_CONSTRUCTION_TEMP;")
qdfCurrent.Execute

DoCmd.Hourglass False

End Function

Function POPULATE_CAPACITIES()
On Error GoTo ERR_POPULATE_CAPACITIES
TRIBE_STATUS = "Populate Capacities"

Dim GOODS_TRIBE(10) As String
Dim GOOD_TRIBE As String
Dim GT_WEIGHT(10) As Double
Dim GT_MNT_CAPACITY(10) As Double
Dim GT_WLK_CAPACITY(10) As Double
Dim OLD_CLAN As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set TRIBECHECK = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TRIBECHECK.index = "PRIMARYKEY"
TRIBECHECK.MoveFirst
         
Do Until TRIBECHECK.EOF
   Call Determine_Capacities("GROUP", TRIBECHECK![CLAN], TRIBECHECK![TRIBE])
   TRIBECHECK.MoveNext
Loop

' Now perform GT's Capacity Calc

OLD_CLAN = "EMPTY"

count = 1

Do While count < 11
   GT_WEIGHT(count) = 0
   GT_MNT_CAPACITY(count) = 0
   GT_WLK_CAPACITY(count) = 0
   GOODS_TRIBE(count) = "Empty"
   count = count + 1
Loop

Set TRIBESTABLE = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESTABLE.index = "PRIMARYKEY"

TRIBECHECK.MoveFirst
Do Until TRIBECHECK.EOF
If TRIBECHECK![CLAN] = OLD_CLAN Then
   'bypass
Else
   'process
OLD_CLAN = TRIBECHECK![CLAN]
TRIBESTABLE.MoveFirst
TRIBESTABLE.Seek "=", TRIBECHECK![CLAN], TRIBECHECK![TRIBE]

If Not TRIBESTABLE.NoMatch Then
   Do Until Not TRIBESTABLE![CLAN] = TRIBECHECK![CLAN]
      If IsNull(TRIBESTABLE![GOODS TRIBE]) Then
         GOOD_TRIBE = TRIBESTABLE![TRIBE]
      Else
         GOOD_TRIBE = TRIBESTABLE![GOODS TRIBE]
      End If
      count = 1
      Do While count < 11
         If GOODS_TRIBE(count) = "EMPTY" Then
            ' 1st Goods Tribe
            GOODS_TRIBE(count) = GOOD_TRIBE
            GT_MNT_CAPACITY(count) = TRIBESTABLE![CAPACITY]
            GT_WLK_CAPACITY(count) = TRIBESTABLE![Walking_Capacity]
            count = 11
         ElseIf GOODS_TRIBE(count) = GOOD_TRIBE Then
            ' 2nd tribe with same Goods Tribe as previous
            GT_MNT_CAPACITY(count) = GT_MNT_CAPACITY(count) + TRIBESTABLE![CAPACITY]
            GT_WLK_CAPACITY(count) = GT_WLK_CAPACITY(count) + TRIBESTABLE![Walking_Capacity]
            count = 11
         Else
            ' new goods tribe
            count = count + 1
         End If
      Loop
      TRIBESTABLE.MoveNext
      If TRIBESTABLE.EOF Then
         Exit Do
      End If
   Loop
End If
   
   
count = 1
Do While count < 11
   If GOODS_TRIBE(count) = "EMPTY" Then
      count = 11
   Else
      TRIBESTABLE.MoveFirst
      TRIBESTABLE.Seek "=", TRIBECHECK![CLAN], GOODS_TRIBE(count)
      If Not TRIBESTABLE.NoMatch Then
         TRIBESTABLE.Edit
         TRIBESTABLE![GT_MOUNTED_CAPACITY] = GT_MNT_CAPACITY(count)
         TRIBESTABLE![GT_WALKING_CAPACITY] = GT_WLK_CAPACITY(count)
         TRIBESTABLE.UPDATE
      End If
      count = count + 1
   End If
Loop

count = 1

Do While count < 11
   GT_WEIGHT(count) = 0
   GT_MNT_CAPACITY(count) = 0
   GT_WLK_CAPACITY(count) = 0
   GOODS_TRIBE(count) = "Empty"
   count = count + 1
Loop
End If

   TRIBECHECK.MoveNext
Loop


TRIBECHECK.Close
TRIBESTABLE.Close

ERR_POPULATE_CAPACITIES_CLOSE:
   Exit Function

ERR_POPULATE_CAPACITIES:
If Err = 3163 Then
   Resume Next
Else
   MSG4 = "Populate Capacities - ERROR MSG = " & Err
   Response = MsgBox(MSG4, True)
   Resume ERR_POPULATE_CAPACITIES_CLOSE
End If

End Function

Public Function POPULATE_WEIGHTS()
On Error GoTo ERR_POPULATE_WEIGHTS
TRIBE_STATUS = "Populate Weights"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set TRIBECHECK = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TRIBECHECK.index = "PRIMARYKEY"
TRIBECHECK.MoveFirst
         
Do Until TRIBECHECK.EOF
   Call Determine_Weights(TRIBECHECK![CLAN], TRIBECHECK![TRIBE])
   TRIBECHECK.MoveNext
Loop

ERR_POPULATE_WEIGHTS_CLOSE:
   Exit Function

ERR_POPULATE_WEIGHTS:
If Err = 3163 Then
   Resume Next
Else
   MSG4 = "Populate Weights - ERROR MSG = " & Err
   Response = MsgBox(MSG4, True)
   Resume ERR_POPULATE_WEIGHTS_CLOSE
End If
End Function



Public Function Ensure_TPs_Are_Setup()

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.MoveFirst

Do
  If HEXMAPCONST![CONSTRUCTION] = "TRADING POST" Then
     Tribes_Current_Hex = HEXMAPCONST![MAP]
     CONSTCLAN = HEXMAPCONST![CLAN]
     CONSTTRIBE = HEXMAPCONST![TRIBE]
     HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "MONTHS TP OPEN"
     If HEXMAPCONST.NoMatch Then
        HEXMAPCONST.AddNew
        HEXMAPCONST![MAP] = Tribes_Current_Hex
        HEXMAPCONST![CLAN] = CONSTCLAN
        HEXMAPCONST![TRIBE] = CONSTTRIBE
        HEXMAPCONST![CONSTRUCTION] = "MONTHS TP OPEN"
        HEXMAPCONST.UPDATE
        HEXMAPCONST.MoveFirst
     Else
        HEXMAPCONST.Seek "=", Tribes_Current_Hex, CONSTCLAN, CONSTTRIBE, "TRADING POST"
     End If
  End If
  HEXMAPCONST.MoveNext
  If HEXMAPCONST.EOF Then
     Exit Do
  End If
Loop
End Function

Public Function Reset_Implements_and_Goods_Usage_Tables()
Dim QUERY As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Implement_Usage;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TRIBES_GOODS_USAGE;")
qdfCurrent.Execute

Call Populate_Implement_Usage_Table

Call Populate_Tribes_Goods_Usage_Table

DoCmd.Hourglass False


End Function
Function CALC_CLAN_RATING()
Dim qdfCurrent As QueryDef
Dim CLANSTATS As Recordset
Dim CLANRATINGS As Recordset
Dim CLANSKILLRATINGS As Recordset
Dim CLANCOMBATRATINGS As Recordset
Dim CLANNAVALRATINGS As Recordset
Dim ratings As String
Dim CLAN_RATING(4000) As Double
Dim CLAN_SKILL_RATING(4000) As Double
Dim CLAN_COMBAT_RATING(4000) As Double
Dim CLAN_NAVAL_RATING(4000) As Double
Dim CLAN_RANK(200, 2) As Double
Dim CLAN_SKILL_RANK(200, 2) As Double
Dim CLAN_COMBAT_RANK(200, 2) As Double
Dim CLAN_NAVAL_RANK(200, 2) As Double 'Dim CLAN_RANKS(200, 2) As Double


Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM CLAN_RATINGS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM CLAN_SKILL_RATINGS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM CLAN_COMBAT_RATINGS;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM CLAN_NAVAL_RATINGS;")
qdfCurrent.Execute

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.MoveFirst

Set SKILLSTABLE = TVDBGM.OpenRecordset("skills")
SKILLSTABLE.index = "PRIMARYKEY"
SKILLSTABLE.MoveFirst

Set COMPRESTABLE = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTABLE.index = "TRIBE"
COMPRESTABLE.MoveFirst

Set CLANSTATS = TVDBGM.OpenRecordset("CLAN_STATS")

Set CLANRATINGS = TVDB.OpenRecordset("CLAN_RATINGS")
Set CLANSKILLRATINGS = TVDB.OpenRecordset("CLAN_SKILL_RATINGS")
Set CLANCOMBATRATINGS = TVDB.OpenRecordset("CLAN_COMBAT_RATINGS")
Set CLANNAVALRATINGS = TVDB.OpenRecordset("CLAN_NAVAL_RATINGS")

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"

For count = 0 To 4000
   CLAN_RATING(count) = 0
   CLAN_SKILL_RATING(count) = 0
   CLAN_COMBAT_RATING(count) = 0
   CLAN_NAVAL_RATING(count) = 0
Next

count = 0

Do Until SKILLSTABLE.EOF
   If Not Left(SKILLSTABLE![TRIBE], 1) = "B" Then
      If Not Left(SKILLSTABLE![TRIBE], 1) = "M" Then
         If SKILLSTABLE![SKILL LEVEL] > 5 Then
            If Len(SKILLSTABLE![TRIBE]) = 3 Then
               count = "0" & Mid(SKILLSTABLE![TRIBE], 2, 2)
            ElseIf Len(SKILLSTABLE![TRIBE]) = 4 Then
               count = "0" & Mid(SKILLSTABLE![TRIBE], 2, 3)
            End If
            CLAN_RATING(count) = CLAN_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            CLAN_SKILL_RATING(count) = CLAN_SKILL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Select Case SKILLSTABLE![Skill]
            Case "ARCHERY"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "CAPTAINCY"
                  CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "CHARIOTRY"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "COMBAT"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "HEAVY WEAPONS"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "HORSEMANSHIP"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "LEADERSHIP"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "MARINER"
                  CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "NAVIGATION"
                  CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "ROWING"
                  CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "SAILING"
                  CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "SCOUTING"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "Seamanship"
                  CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "SECURITY"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "SPYING"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            Case "TACTICS"
                  CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + SKILLSTABLE![SKILL LEVEL] * 10
            End Select
         End If
      End If
   End If
   SKILLSTABLE.MoveNext
Loop

Do Until COMPRESTABLE.EOF
   If COMPRESTABLE![TRIBE] = "ZZZZ" Then
      Exit Do
   End If
   If Not Left(COMPRESTABLE![TRIBE], 1) = "B" Then
      If Not Left(COMPRESTABLE![TRIBE], 1) = "M" Then
         If Len(COMPRESTABLE![TRIBE]) = 3 Then
            count = "0" & Mid(COMPRESTABLE![TRIBE], 2, 2)
         ElseIf Len(COMPRESTABLE![TRIBE]) = 4 Then
            count = "0" & Mid(COMPRESTABLE![TRIBE], 2, 3)
         End If
         CLAN_RATING(count) = CLAN_RATING(count) + 100
         CLAN_SKILL_RATING(count) = CLAN_SKILL_RATING(count) + 100
      End If
   End If
   COMPRESTABLE.MoveNext
   If COMPRESTABLE.EOF Then
      Exit Do
   End If
   If COMPRESTABLE![TRIBE] = "ZZZZ" Then
      Exit Do
   End If
Loop

Do Until CLANSTATS.EOF
   VALIDGOODS.MoveFirst
   VALIDGOODS.Seek "=", CLANSTATS![GOOD]
   If Not VALIDGOODS.NoMatch Then
      If Not Left(CLANSTATS![CLAN], 1) = "B" Then
         If Not Left(CLANSTATS![CLAN], 1) = "M" Then
            If Not Left(CLANSTATS![CLAN], 1) = "Z" Then
               count = CLANSTATS![CLAN]
               RATING = VALIDGOODS![RATING]
               CLAN_RATING(count) = CLAN_RATING(count) + (RATING * CLANSTATS![NUMBER])
               If VALIDGOODS![TABLE] = "WAR" Then
                   CLAN_COMBAT_RATING(count) = CLAN_COMBAT_RATING(count) + (RATING * CLANSTATS![NUMBER])
               ElseIf VALIDGOODS![TABLE] = "SHIP" Then
                   CLAN_NAVAL_RATING(count) = CLAN_NAVAL_RATING(count) + (RATING * CLANSTATS![NUMBER])
               End If
            End If
         End If
      End If
   End If
   CLANSTATS.MoveNext

Loop

   VALIDGOODS.MoveFirst
   VALIDGOODS.Seek "=", "SLAVE"
   RATING = VALIDGOODS![RATING]
Do Until TRIBESINFO.EOF
   If Not Left(TRIBESINFO![CLAN], 1) = "B" Then
      If Not Left(TRIBESINFO![CLAN], 1) = "M" Then
      If Not Left(TRIBESINFO![CLAN], 1) = "Z" Then
         If Len(TRIBESINFO![CLAN]) = 3 Then
            count = "0" & Mid(TRIBESINFO![CLAN], 2, 2)
         ElseIf Len(TRIBESINFO![CLAN]) = 4 Then
            count = "0" & Mid(TRIBESINFO![CLAN], 2, 3)
         End If
         CLAN_RATING(count) = CLAN_RATING(count) + (RATING * TRIBESINFO![SLAVE])
      End If
      End If
   End If
   TRIBESINFO.MoveNext

Loop

For count = 0 To 200
   CLAN_RANK(count, 1) = 0
   CLAN_RANK(count, 2) = 0
Next


' General Clan Ratings
For count = 0 To 200
   RATING = CLAN_RATING(count)
   RECORD_COUNT = count
   For Counter = 0 To 4000
     If CLAN_RATING(Counter) > RATING Then
        RATING = CLAN_RATING(Counter)
        RECORD_COUNT = Counter
     End If
   Next
   CLAN_RATING(RECORD_COUNT) = 0
   CLAN_RANK(count, 1) = RATING
   CLAN_RANK(count, 2) = RECORD_COUNT

Next

' Skill Clan Ratings
For count = 0 To 200
   RATING = CLAN_SKILL_RATING(count)
   RECORD_COUNT = count
   For Counter = 0 To 4000
     If CLAN_SKILL_RATING(Counter) > RATING Then
        RATING = CLAN_SKILL_RATING(Counter)
        RECORD_COUNT = Counter
     End If
   Next
   CLAN_SKILL_RATING(RECORD_COUNT) = 0
   CLAN_SKILL_RANK(count, 1) = RATING
   CLAN_SKILL_RANK(count, 2) = RECORD_COUNT

Next

' Combat Clan Ratings
For count = 0 To 200
   RATING = CLAN_COMBAT_RATING(count)
   RECORD_COUNT = count
   For Counter = 0 To 4000
     If CLAN_COMBAT_RATING(Counter) > RATING Then
        RATING = CLAN_COMBAT_RATING(Counter)
        RECORD_COUNT = Counter
     End If
   Next
   CLAN_COMBAT_RATING(RECORD_COUNT) = 0
   CLAN_COMBAT_RANK(count, 1) = RATING
   CLAN_COMBAT_RANK(count, 2) = RECORD_COUNT

Next

' Naval Clan Ratings
For count = 0 To 200
   RATING = CLAN_NAVAL_RATING(count)
   RECORD_COUNT = count
   For Counter = 0 To 4000
     If CLAN_NAVAL_RATING(Counter) > RATING Then
        RATING = CLAN_NAVAL_RATING(Counter)
        RECORD_COUNT = Counter
     End If
   Next
   CLAN_NAVAL_RATING(RECORD_COUNT) = 0
   CLAN_NAVAL_RANK(count, 1) = RATING
   CLAN_NAVAL_RANK(count, 2) = RECORD_COUNT

Next

MSG1 = "Clan ratings : "
 For count = 1 To 200
   If CLAN_RANK(count, 1) > 0 Then
      CLANRATINGS.AddNew
      CLANRATINGS![CLAN] = CLAN_RANK(count, 2)
      CLANRATINGS![RATING] = CLAN_RANK(count, 1)
      MSG1 = MSG1 & CLAN_RANK(count, 2) & ", "
      MSG2 = "?? " & CLAN_RANK(count, 1)
      'RESPONSE = MsgBox(MSG1 & MSG2, True)
      CLANRATINGS.UPDATE
   End If
Next

'MsgBox (MSG)
'MsgBox (MSG1)
ratings = InputBox("Current Rankings?", "RANKINGS", MSG1)

MSG1 = "Clan Skill ratings : "
For count = 0 To 200
   If CLAN_SKILL_RANK(count, 1) > 0 Then
      CLANSKILLRATINGS.AddNew
      CLANSKILLRATINGS![CLAN] = CLAN_SKILL_RANK(count, 2)
      CLANSKILLRATINGS![RATING] = CLAN_SKILL_RANK(count, 1)
      MSG1 = MSG1 & CLAN_SKILL_RANK(count, 2) & ", "
      MSG2 = "?? " & CLAN_SKILL_RANK(count, 1)
      'RESPONSE = MsgBox(MSG1 & MSG2, True)
      CLANSKILLRATINGS.UPDATE
   End If
Next

'MsgBox (MSG)
'MsgBox (MSG1)
ratings = InputBox("Current Rankings?", "SKILL RANKINGS", MSG1)

MSG1 = "Clan Combat ratings : "
For count = 0 To 200
   If CLAN_COMBAT_RANK(count, 1) > 0 Then
      CLANCOMBATRATINGS.AddNew
      CLANCOMBATRATINGS![CLAN] = CLAN_COMBAT_RANK(count, 2)
      CLANCOMBATRATINGS![RATING] = CLAN_COMBAT_RANK(count, 1)
      MSG1 = MSG1 & CLAN_COMBAT_RANK(count, 2) & ", "
      MSG2 = "?? " & CLAN_COMBAT_RANK(count, 1)
      'RESPONSE = MsgBox(MSG1 & MSG2, True)
      CLANCOMBATRATINGS.UPDATE
   End If
Next

'MsgBox (MSG)
'MsgBox (MSG1)
ratings = InputBox("Current Rankings?", "COMBAT RANKINGS", MSG1)

MSG1 = "Clan Naval ratings : "
For count = 0 To 200
   If CLAN_NAVAL_RANK(count, 1) > 0 Then
      CLANNAVALRATINGS.AddNew
      CLANNAVALRATINGS![CLAN] = CLAN_NAVAL_RANK(count, 2)
      CLANNAVALRATINGS![RATING] = CLAN_NAVAL_RANK(count, 1)
      MSG1 = MSG1 & CLAN_NAVAL_RANK(count, 2) & ", "
      MSG2 = "?? " & CLAN_NAVAL_RANK(count, 1)
      'RESPONSE = MsgBox(MSG1 & MSG2, True)
      CLANNAVALRATINGS.UPDATE
   End If
Next

'MsgBox (MSG)
'MsgBox (MSG1)
ratings = InputBox("Current Rankings?", "NAVAL RANKINGS", MSG1)

End Function



Function HEX_POPULATION(CLAN, TRIBE, CurrentHex)
Dim CURRENT_HEX As String
Dim TRIBES_IN_HEX As String
Dim hex_pop As Long
Dim Tribes(30) As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
hex_pop = 0

' Current Hexes population

Set HEXMAPPOLITICS = TVDBGM.OpenRecordset("HEX_MAP_politics")
HEXMAPPOLITICS.index = "SECONDARYKEY"
If Not HEXMAPPOLITICS.EOF Then
   HEXMAPPOLITICS.MoveFirst
   HEXMAPPOLITICS.Seek "=", CLAN, TRIBE

   If Not HEXMAPPOLITICS.NoMatch Then
      hex_pop = HEXMAPPOLITICS![POPULATION]
   Else
      ' no hex - wow - real problem
   End If
End If

' Identify who is in the hex

TRIBESINHEX = WHO_IS_IN_HEX(CLAN, TRIBE, CurrentHex, "N")

'got an error, there is spaces after the comma

Iteration_Count(1) = 1
Tribes(Iteration_Count(1)) = TRIBE
Iteration_Count(1) = Iteration_Count(1) + 1

Do Until Iteration_Count(1) > 30
   BRACKET = 0
   BRACKET = InStr(TRIBESINHEX, ",")
   If BRACKET > 0 Then
      Tribes(Iteration_Count(1)) = Left(TRIBESINHEX, (BRACKET - 1))
   Else
      Tribes(Iteration_Count(1)) = TRIBESINHEX
      TRIBESINHEX = " EMPTY"
   End If
   TRIBESINHEX = Right(TRIBESINHEX, (Len(TRIBESINHEX) - (BRACKET + 1)))
   Iteration_Count(1) = Iteration_Count(1) + 1
   If TRIBESINHEX = "EMPTY" Then
    Exit Do
   End If
Loop

' Now add warriors, actives and inactives of all members of the clan

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.MoveFirst

Iteration_Count(1) = 1
Do Until Iteration_Count(1) > 30

   TRIBESINFO.Seek "=", CLAN, Tribes(Iteration_Count(1))
   
   If TRIBESINFO.NoMatch Then
      ' ignore
   Else
      hex_pop = hex_pop + TRIBESINFO![WARRIORS]
      hex_pop = hex_pop + TRIBESINFO![ACTIVES]
      hex_pop = hex_pop + TRIBESINFO![INACTIVES]
   End If

   Iteration_Count(1) = Iteration_Count(1) + 1
   TRIBESINFO.MoveFirst

Loop

HEX_POPULATION = hex_pop

HEXMAPPOLITICS.Close
TRIBESINFO.Close

End Function

Function TEST_HEX_POPULATION()
Dim hex_pop As Long

hex_pop = HEX_POPULATION("0330", "0330", "GJ 1720")

End Function

Function CLEAN_UP_BLANK_ROWS(WHICH_TABLES As String)
Dim QUERY As String
Dim strSQL As String

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "TRIBES_ACTIVITY" Then
   ' clean up process_tribes_activity
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBES_ACTIVITY"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBES_ACTIVITY.TRIBE IS NULL);"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute

   QUERY_STRING = "DELETE * FROM PROCESS_TRIBES_ACTIVITY"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBES_ACTIVITY.TRIBE='');"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute

   ' clean up process_tribes_item_allocation
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBES_ITEM_ALLOCATION"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBES_ITEM_ALLOCATION.TRIBE IS NULL);"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
  
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBES_ITEM_ALLOCATION"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBES_ITEM_ALLOCATION.TRIBE='');"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "TRIBE_MOVEMENT" Then
   ' clean up process_tribe_movement
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBE_MOVEMENT"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBE_MOVEMENT.TRIBE IS NULL);"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
  
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBE_MOVEMENT"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBE_MOVEMENT.TRIBE='');"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "TRIBE_MOVEMENT_COPY" Then
   ' clean up process_tribe_movement
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBE_MOVEMENT_COPY"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBE_MOVEMENT_COPY.TRIBE IS NULL);"
   Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
  
   QUERY_STRING = "DELETE * FROM PROCESS_TRIBE_MOVEMENT_COPY"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_TRIBE_MOVEMENT_COPY.TRIBE='');"
   Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "SKILLS" Then
   ' clean up process_skills
   QUERY_STRING = "DELETE * FROM PROCESS_SKILLS"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_SKILLS.TRIBE IS NULL);"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
  
   QUERY_STRING = "DELETE * FROM PROCESS_SKILLS"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_SKILLS.TRIBE='');"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "RESEARCH" Then
   ' clean up process_research
   QUERY_STRING = "DELETE * FROM PROCESS_RESEARCH"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_RESEARCH.TRIBE IS NULL);"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
     
   QUERY_STRING = "DELETE * FROM PROCESS_RESEARCH"
   QUERY_STRING = QUERY_STRING & " WHERE (PROCESS_RESEARCH.TRIBE='');"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "TRANSFERS" Then
   ' clean up process_transfers
   QUERY_STRING = "DELETE * FROM MASSTRANSFERS"
   QUERY_STRING = QUERY_STRING & " WHERE (MASSTRANSFERS.FROM IS NULL);"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
     
   QUERY_STRING = "DELETE * FROM MASSTRANSFERS"
   QUERY_STRING = QUERY_STRING & " WHERE (MASSTRANSFERS.FROM='');"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If

If WHICH_TABLES = "ALL" Or WHICH_TABLES = "TRANSFERS_COPY" Then
   ' clean up process_transfers
   QUERY_STRING = "DELETE * FROM MASSTRANSFERS_COPY"
   QUERY_STRING = QUERY_STRING & " WHERE (MASSTRANSFERS_COPY.FROM IS NULL);"
   Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
     
   QUERY_STRING = "DELETE * FROM MASSTRANSFERS_COPY"
   QUERY_STRING = QUERY_STRING & " WHERE (MASSTRANSFERS_COPY.FROM='');"
   Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
End If


DoCmd.Hourglass False

End Function

Function Tribe_Checking(ACTION, CLAN, TRIBE, HEX)
On Error GoTo ERR_TC_CHECK
TRIBE_STATUS = "Tribes Checking"
DebugOP "GLOBAL FUNCTIONS > TribeChecking()"

Dim qdfCurrent As QueryDef
Dim QUERY As String
Dim strSQL As String

' functions
' Update all records in Tribe_Checking
' Update only Current_Hex
' Retrieve Current_Hex
' Retrieve Current_Population
' Clan, Tribe, Current Hex, Provs, Warriors, Actives, Inactives, Slave
' Provs is empty
' Indexes - HEX, PrimaryKey (Clan, Tribe), TRIBE

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)


If ACTION = "Update_All" Then
   ' Empty table
   Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM TRIBE_CHECKING;")
   qdfCurrent.Execute

   ' Repopulate table
   QUERY_STRING = "INSERT INTO TRIBE_CHECKING ( CLAN, TRIBE, [CURRENT HEX], WARRIORS, ACTIVES,"
   QUERY_STRING = QUERY_STRING & " INACTIVES, SLAVE ) SELECT CLAN, TRIBE, [CURRENT HEX],"
   QUERY_STRING = QUERY_STRING & " WARRIORS, ACTIVES, INACTIVES, SLAVE"
   QUERY_STRING = QUERY_STRING & " FROM TRIBES_GENERAL_INFO;"
   Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
   
   ' Still need to populate Provs

   Set TRIBECHECK = TVDBGM.OpenRecordset("Tribe_CHECKING")
   TRIBECHECK.index = "PRIMARYKEY"
   TRIBECHECK.MoveFirst

   Do While Not TRIBECHECK.EOF
      Num_Goods = GET_TRIBES_GOOD_QUANTITY(TRIBECHECK![CLAN], TRIBECHECK![TRIBE], "PROVS")
      TRIBECHECK.Edit
      TRIBECHECK![Provs] = Num_Goods
      TRIBECHECK.UPDATE

      TRIBECHECK.MoveNext
   Loop
ElseIf ACTION = "Update_Hex" Then
   Set TRIBECHECK = TVDBGM.OpenRecordset("Tribe_CHECKING")
   TRIBECHECK.index = "PRIMARYKEY"
   TRIBECHECK.MoveFirst
   TRIBECHECK.Seek "=", CLAN, TRIBE
   TRIBECHECK.Edit
   TRIBECHECK![CURRENT HEX] = HEX
   TRIBECHECK.UPDATE

ElseIf ACTION = "Get_Hex" Then
   Set TRIBECHECK = TVDBGM.OpenRecordset("Tribe_CHECKING")
   TRIBECHECK.index = "PRIMARYKEY"
   TRIBECHECK.MoveFirst
   TRIBECHECK.Seek "=", CLAN, TRIBE
   Tribe_Checking_Hex = TRIBECHECK![CURRENT HEX]

ElseIf ACTION = "Get_People" Then
   Set TRIBECHECK = TVDBGM.OpenRecordset("Tribe_CHECKING")
   TRIBECHECK.index = "PRIMARYKEY"
   TRIBECHECK.MoveFirst
   TRIBECHECK.Seek "=", CLAN, TRIBE
   Tribe_Checking_People = TRIBECHECK![WARRIORS] + TRIBECHECK![ACTIVES] + TRIBECHECK![INACTIVES] + TRIBECHECK![SLAVE]

ElseIf ACTION = "Get_Provs" Then
   Set TRIBECHECK = TVDBGM.OpenRecordset("Tribe_CHECKING")
   TRIBECHECK.index = "PRIMARYKEY"
   TRIBECHECK.MoveFirst
   TRIBECHECK.Seek "=", CLAN, TRIBE
   Tribe_Checking_Provs = TRIBECHECK![Provs]

Else
   ' NEW???

End If

TRIBECHECK.MoveFirst

ERR_TC_CHECK_CLOSE:
   Exit Function

ERR_TC_CHECK:
If (Err = 3022) Or (Err = 3163) Then
   Resume Next
Else
   MSG1 = "ERROR = " & Err
   Response = MsgBox(MSG1, True)
   Resume ERR_TC_CHECK_CLOSE
End If

End Function

Function Unit_Check(ACTION, TRIBE)
'*===============================================================================*'
'*****    This function is used to retrieve the information related to the   *****'
'*****    tribe or subtribe for a specific unit type                         *****'
'*****                                                                       *****'
'*****     i.e. skill tribe (movement)                                       *****'
'*****          dice tribe  (movement)                                       *****'
'*****                                                                       *****'
'*****                                                                       *****'
'*****                                                                       *****'
'*****                                                                       *****'
'*****                                                                       *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 14/02/19  *  Insert Maintenance Log                                         **'
'**           *                                                                 **'
'**           *                                                                 **'
'*===============================================================================*'
On Error GoTo ERR_UC_CHECK
Dim ValidUnit As Recordset
Dim ValidDenomiations As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

'Set ValidUnit = TVDB.OpenRecordset("VALID_UNIT")
'ValidUnit.Index = "PRIMARYKEY"
'ValidUnit.MoveFirst
'Assume all units have the first four numbers to denote the tribe/subtribe
'Assume the next 1 through 4 will be charactes to denote unit type
'Assume the last 1 through 2 numbers will be the number of that unit type for the subtribe

ValidDenomiations = ""
'Do Until ValidUnit.EOF
'   If Not IsNull(ValidUnit![Denomination]) Then
'      ValidDenomiations = ValidDenomiations & ValidUnit![Denomination]
'   End If
'   ValidUnit.MoveNext
'Loop

' for movement module

If Len(TRIBE) = 4 Then ' c, e, f & g
   ' tribe or subtribe or bandit
   If Left(TRIBE, 1) = "B" Then
      SKILL_MOVE_TRIBE = TRIBE
      DICE_TRIBE = Right(TRIBE, 3)
   Else
      SKILL_MOVE_TRIBE = Left(TRIBE, 4)
      DICE_TRIBE = Left(TRIBE, 4)
   End If
ElseIf Len(TRIBE) >= 5 Then
   SKILL_MOVE_TRIBE = Left(TRIBE, 4)
   DICE_TRIBE = Left(TRIBE, 4)
Else
   ' should never drop to here
   SKILL_MOVE_TRIBE = TRIBE
   DICE_TRIBE = TRIBE
End If

If ACTION = "DICE" Or ACTION = "TRIBE" Then
 If InStr(TRIBE, "ELE") Then
    Unit_Check = Left(TRIBE, 4)
 ElseIf InStr(TRIBE, "GAR") Then
    Unit_Check = Left(TRIBE, 4)
 ElseIf InStr(TRIBE, "FLE") Then
    Unit_Check = Left(TRIBE, 4)
 ElseIf InStr(TRIBE, "COU") Then
    Unit_Check = Left(TRIBE, 4)
 ElseIf InStr(TRIBE, "AGE") Then
    Unit_Check = Left(TRIBE, 4)
 ElseIf InStr(TRIBE, "ARC") Then
    Unit_Check = Left(TRIBE, 4)
 ElseIf Left(TRIBE, 1) = "B" Then
    Unit_Check = Right(TRIBE, 3)
 ElseIf Left(TRIBE, 1) = "M" Then
    Unit_Check = TRIBE
 ElseIf Len(TRIBE) > 4 Then
    Unit_Check = Left(TRIBE, 4)
 Else
    Unit_Check = TRIBE
  End If
End If

ERR_UC_CHECK_CLOSE:
   Exit Function

ERR_UC_CHECK:
If (Err = 3022) Then
   Resume Next
Else
   MSG1 = "ERROR = " & Err
   Response = MsgBox(MSG1, True)
   Resume ERR_UC_CHECK_CLOSE
End If

End Function
Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function
Function Export_Tables()
On Error GoTo ERR_EXP_TAB
Dim DocDir, fileName As String

CURRENT_DIRECTORY = Dir(DIRECTPATH, vbDirectory)
If IsNull(CURRENT_DIRECTORY) Or CURRENT_DIRECTORY = "" Then
  MkDir (DIRECTPATH)
End If

DocDir = CurDir$ & "\New Code\"
fileName = DocDir & "ACTIVITIES.txt"
If FileExists(fileName) Then
   fileName = DocDir & "ACTIVITIES1.txt"
End If
DoCmd.TransferText acExportDelim, , "ACTIVITIES", fileName, True
fileName = DocDir & "ACTIVITY.txt"
If FileExists(fileName) Then
   fileName = DocDir & "ACTIVITY1.txt"
End If
DoCmd.TransferText acExportDelim, , "ACTIVITY", fileName, True
fileName = DocDir & "IMPLEMENTS.txt"
If FileExists(fileName) Then
   fileName = DocDir & "IMPLEMENTS1.txt"
End If
DoCmd.TransferText acExportDelim, , "IMPLEMENTS", fileName, True
fileName = DocDir & "RELIGION.txt"
If FileExists(fileName) Then
   fileName = DocDir & "RELIGION1.txt"
End If
DoCmd.TransferText acExportDelim, , "RELIGION", fileName, True
fileName = DocDir & "RESEARCH.txt"
If FileExists(fileName) Then
   fileName = DocDir & "RESEARCH1.txt"
End If
DoCmd.TransferText acExportDelim, , "RESEARCH", fileName, True
fileName = DocDir & "Valid_Animals.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Animals1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Animals", fileName, True
fileName = DocDir & "Valid_Borders.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Borders1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Borders", fileName, True
fileName = DocDir & "Valid_Buildings.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Buildings1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Buildings", fileName, True
fileName = DocDir & "Valid_Crops.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Crops1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Crops", fileName, True
fileName = DocDir & "Valid_Minerals.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Minerals1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Minerals", fileName, True
fileName = DocDir & "Valid_Modifiers.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Modifiers1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Modifiers", fileName, True
fileName = DocDir & "Valid_Ships.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Ships1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Ships", fileName, True
fileName = DocDir & "Valid_Skills.txt"
If FileExists(fileName) Then
   fileName = DocDir & "Valid_Skills1.txt"
End If
DoCmd.TransferText acExportDelim, , "Valid_Skills", fileName, True
   
ERR_EXP_TAB_CLOSE:
   Exit Function

ERR_EXP_TAB:
   Resume Next

End Function

Public Function FileExists(ByVal path_ As String) As Boolean
    FileExists = (Len(Dir(path_)) > 0)
End Function

Public Function Close_Table(TABLE_NAME As String)

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDB.OpenRecordset(TABLE_NAME)
TRIBESINFO.MoveFirst
TRIBESINFO.Close

End Function

Public Function Update_Weights_Capacities()

Call POPULATE_CAPACITIES

Call POPULATE_WEIGHTS

Call Tribe_Checking("Update_All", "", "", "")



End Function

Public Function ConvertBuildingsTable()
' Converts HEX_MAP_CONST table to new format where absence of container
' building marked as -1 instead of 0
Dim Counter As Long
Dim MSG1 As String
Set HEXCONSTTABLE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXCONSTTABLE.index = "TRIBECONST"
HEXCONSTTABLE.MoveFirst

Set VALIDBUILDINGS = TVDB.OpenRecordset("VALID_BUILDINGS")
VALIDBUILDINGS.index = "PRIMARYKEY"


Do Until HEXCONSTTABLE.EOF

   VALIDBUILDINGS.MoveFirst
   VALIDBUILDINGS.Seek "=", HEXCONSTTABLE![CONSTRUCTION]
   If Not VALIDBUILDINGS.NoMatch Then
     If VALIDBUILDINGS![LIMITS] >= 10 Then
        If HEXCONSTTABLE![1] <> -1 And _
           HEXCONSTTABLE![2] <> -1 And _
           HEXCONSTTABLE![3] <> -1 And _
           HEXCONSTTABLE![4] <> -1 And _
           HEXCONSTTABLE![5] <> -1 And _
           HEXCONSTTABLE![6] <> -1 And _
           HEXCONSTTABLE![7] <> -1 And _
           HEXCONSTTABLE![8] <> -1 And _
           HEXCONSTTABLE![9] <> -1 And _
           HEXCONSTTABLE![10] <> -1 Then
           HEXCONSTTABLE.Edit
           If HEXCONSTTABLE![1] = 0 Then
                 HEXCONSTTABLE![1] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![2] = 0 Then
                 HEXCONSTTABLE![2] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![3] = 0 Then
                 HEXCONSTTABLE![3] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![4] = 0 Then
                 HEXCONSTTABLE![4] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![5] = 0 Then
                 HEXCONSTTABLE![5] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![6] = 0 Then
                 HEXCONSTTABLE![6] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![7] = 0 Then
                 HEXCONSTTABLE![7] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![8] = 0 Then
                 HEXCONSTTABLE![8] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![9] = 0 Then
                 HEXCONSTTABLE![9] = -1
                 Counter = Counter + 1
              End If
              If HEXCONSTTABLE![10] = 0 Then
                 HEXCONSTTABLE![10] = -1
                 Counter = Counter + 1
              End If
              HEXCONSTTABLE.UPDATE
         End If
     End If
  Else
   MSG1 = "Buildings table conversion. Construction " & HEXCONSTTABLE![CONSTRUCTION] & " was not found in Valid Buildings table"
   Response = MsgBox(MSG1, True)
  End If
   HEXCONSTTABLE.MoveNext
   If HEXCONSTTABLE.EOF Then
         Exit Do
   End If
Loop

If (Counter > 0) Then
   MSG1 = "Buildings table conversion. " & Counter & " converted."
   Response = MsgBox(MSG1, True)
End If
HEXCONSTTABLE.Close
VALIDBUILDINGS.Close
End Function

