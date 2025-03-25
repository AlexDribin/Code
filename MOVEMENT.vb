Attribute VB_Name = "MOVEMENT"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*                          VERSION 3.1.1                                        *'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 29/01/96  *  Allow for Scouting Finds                                       **'
'** 23/02/96  *  Amended Naval movement                                         **'
'** 05/03/25  *  Movement points calculation for mounted units fixed (AlexD)    **'
'** 05/03/25  *  Prevents moving wagons without pulling animals. (AlexD)        **'
'** 16/03/25  *  Mounted capacity check is currently disabled (AlexD)           **'
'*===============================================================================*'
 



Sub ADD_NEW_HEX(CURRENT_MAP, TERRAIN)
Dim FORMARG As String

If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = ADD NEW HEX" & crlf
   Response = MsgBox((MSG1), True)
End If

hexmaptable.AddNew
'MSG = "CURRENT MAP = " & CURRENT_MAP
'RESPONSE = MsgBox(MAP, True)
hexmaptable![MAP] = CURRENT_MAP

hexmaptable.UPDATE

FORMARG = "[MAP] = """ & CURRENT_MAP & """"

DoCmd.Hourglass False

DoCmd.OpenForm "HEX_MAP", , , FORMARG, A_EDIT, A_DIALOG

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.MoveFirst
hexmaptable.index = "PRIMARYKEY"

DoCmd.Hourglass True

End Sub


Public Function Obtain_Skill_Levels()
Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "PRIMARYKEY"
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "SCOUTING"
If SKILLSTABLE.NoMatch Then
   SCOUTING_LEVEL = 1             ' SCOUTING LEVEL CANNOT BE ZERO FOR SCOUTING
Else
   SCOUTING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "APIARISM"
If SKILLSTABLE.NoMatch Then
   APIARISM_LEVEL = 0
Else
   APIARISM_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "FORESTRY"
If SKILLSTABLE.NoMatch Then
   FORESTRY_LEVEL = 0
Else
   FORESTRY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "DIPLOMACY"
If SKILLSTABLE.NoMatch Then
   DIPLOMACY_LEVEL = 0
Else
   DIPLOMACY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "NAVIGATION"
If SKILLSTABLE.NoMatch Then
   NAVIGATION_LEVEL = 0
Else
   NAVIGATION_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "SAILING"
If SKILLSTABLE.NoMatch Then
   SAILING_LEVEL = 0
Else
   SAILING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "ROWING"
If SKILLSTABLE.NoMatch Then
   ROWING_LEVEL = 0
Else
   ROWING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "SEAMANSHIP"
If SKILLSTABLE.NoMatch Then
   SEAMANSHIP_LEVEL = 0
Else
   SEAMANSHIP_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", SKILL_MOVE_TRIBE, "POLITICS"
If SKILLSTABLE.NoMatch Then
   POLITICS_LEVEL = 0
Else
   POLITICS_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

If codetrack = 1 Then
   MSG1 = "SCOUTING_LEVEL = " & SCOUTING_LEVEL & crlf
   MSG2 = "APIARISM_LEVEL = " & APIARISM_LEVEL & crlf
   MSG3 = "FORESTRY_LEVEL = " & FORESTRY_LEVEL & crlf
   MSG4 = "DIPLOMACY_LEVEL = " & DIPLOMACY_LEVEL & crlf
   MSG5 = "NAVIGATION_LEVEL = " & NAVIGATION_LEVEL & crlf
   MSG6 = "SAILING_LEVEL = " & SAILING_LEVEL & crlf
   MSG7 = "ROWING_LEVEL = " & ROWING_LEVEL & crlf
   MSG8 = "SEAMANSHIP_LEVEL = " & SEAMANSHIP_LEVEL & crlf
   Response = MsgBox((MSG1 & MSG2 & MSG3 & MSG4 & MSG5 & MSG6 & MSG7 & MSG8), True)
End If

SKILLSTABLE.Close


End Function

Sub CALC_SCOUTING_FINDS()
On Error GoTo CALC_SCOUTING_FINDS_ERROR
CSF_POS = "START"

If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = CALC_SCOUTING_FINDS" & crlf
   Response = MsgBox((MSG1), True)
End If

If SCOUT_MISSION = "NONE" Then
   If WEATHER = "WIND" Or WEATHER = "L-SNOW" Then
      FIND_CHANCE = (3 + SCOUTING_LEVEL) - 1
   ElseIf WEATHER = "H-SNOW" Or WEATHER = "H-RAIN" Or WEATHER = "L-RAIN" Then
      FIND_CHANCE = (3 + SCOUTING_LEVEL) - 2
   Else
      FIND_CHANCE = (3 + SCOUTING_LEVEL)
   End If
Else
   FIND_CHANCE = (3 + SCOUTING_LEVEL) / 2
End If

CURRENT_SEASON = GET_SEASON(TURN_CURRENT)
roll1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
roll2 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)

CSF_POS = "CALC DEATH CHANCE"
'DEATH_CHANCE = (((1 / SCOUTING_LEVEL) * (1 / (SCOUTS_USED ^ (1 / 3)))) * 75) 'previous formula
DEATH_CHANCE = (1 / (SCOUTING_LEVEL ^ 0.8)) * (1 / (SCOUTS_USED ^ 0.3)) * 0.01

' 20230804 - commented out due to code giving topsy turvy results - Andy636
'If DEATH_CHANCE < 1 Then
'   DEATH_CHANCE = 0
'Else
'   DEATH_CHANCE = (5 + 5) / DEATH_CHANCE
'End If

Randomize
If Rnd() <= DEATH_CHANCE Then
   
   TRIBESINFO.Edit
   TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] - SCOUTS_USED
   TRIBESINFO.UPDATE
   If HORSES_USED > 0 Then
      Call UPDATE_TRIBES_TABLES("HORSE", "SUBTRACT", HORSES_USED)
   End If
   If ELEPHANTS_USED > 0 Then
      Call UPDATE_TRIBES_TABLES("ELEPHANT", "SUBTRACT", ELEPHANTS_USED)
   End If
   If CAMELS_USED > 0 Then
      Call UPDATE_TRIBES_TABLES("CAMEL", "SUBTRACT", CAMELS_USED)
   End If
   MOVEMENT_LINE = "Scout Group did not return"
   Scout_Result.AddNew
   Scout_Result![TRIBE] = MOVE_TRIBE
   Scout_Result![SCOUT] = SCOUT_NUMBER
   Scout_Result![MISSION] = "Scouts Died"
   Scout_Result![FOUND] = ""
   Scout_Result![Results] = "Already deducted warriors and mounts"
   Scout_Result.UPDATE
   MSG1 = MOVE_TRIBE & " lost scout group " & SCOUT_NUMBER & crlf
   MSG2 = "Warriors and animals have been deleted "
   Response = MsgBox((MSG1 & MSG2), True)
   GoTo END_FINDS:
End If
  
CSF_POS = "LOCATE MISSION"
'  ALLOW FOR LOCATE/SPY/RAID/PACIFY
If SCOUT_MISSION = "LOCATE" Then
   ' WHO IS BEING LOCATED
   TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
   If TRIBESINHEX = "EMPTY" Then
       MOVEMENT_LINE = MOVEMENT_LINE & " No Groups Located"
   Else
      Scout_Result.AddNew
      Scout_Result![TRIBE] = MOVE_TRIBE
      Scout_Result![SCOUT] = SCOUT_NUMBER
      Scout_Result![MISSION] = SCOUT_MISSION
      Scout_Result![FOUND] = TRIBESINHEX
      Scout_Result![Results] = "Had locate orders and found these tribes"
      Scout_Result.UPDATE
      TRIBESINHEX_NEW = Check_Truced(MOVE_CLAN, Truced_Clans, TRIBESINHEX)
      MOVEMENT_LINE = MOVEMENT_LINE & " Located " & TRIBESINHEX
      If Len(TRIBESINHEX_NEW) > 3 Then
         MSG1 = "Locate Orders from " & MOVE_TRIBE & " found the following groups " & TRIBESINHEX_NEW
         Response = MsgBox((MSG1), True)
      End If
   
   End If
   GoTo END_FINDS:
End If

CSF_POS = "SPY MISSION"
If SCOUT_MISSION = "SPY" Then
   ' WHO IS BEING SPIED ON
   TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
   If TRIBESINHEX = "EMPTY" Then
       MOVEMENT_LINE = MOVEMENT_LINE & " No Groups Spied On"
   Else
      Scout_Result.AddNew
      Scout_Result![TRIBE] = MOVE_TRIBE
      Scout_Result![SCOUT] = SCOUT_NUMBER
      Scout_Result![MISSION] = SCOUT_MISSION
      Scout_Result![FOUND] = TRIBESINHEX
      Scout_Result![Results] = "Had spy orders and found these tribes"
      Scout_Result.UPDATE
      TRIBESINHEX_NEW = Check_Truced(MOVE_CLAN, Truced_Clans, TRIBESINHEX)
      MOVEMENT_LINE = MOVEMENT_LINE & " May have spied on " & TRIBESINHEX
      If Len(TRIBESINHEX_NEW) > 3 Then
         MSG1 = "Spying Orders from " & MOVE_TRIBE & " found the following groups " & TRIBESINHEX_NEW
         Response = MsgBox((MSG1), True)
      End If
   End If

   GoTo END_FINDS:
End If

CSF_POS = "RAID MISSION"
If SCOUT_MISSION = "RAID" Then
   ' WHO IS BEING RAIDED
   TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
   If TRIBESINHEX = "EMPTY" Then
       MOVEMENT_LINE = MOVEMENT_LINE & " No Groups Raided"
   Else
      Scout_Result.AddNew
      Scout_Result![TRIBE] = MOVE_TRIBE
      Scout_Result![SCOUT] = SCOUT_NUMBER
      Scout_Result![MISSION] = SCOUT_MISSION
      Scout_Result![FOUND] = TRIBESINHEX
      Scout_Result![Results] = "Had raid orders and found these tribes"
      Scout_Result.UPDATE
      TRIBESINHEX_NEW = Check_Truced(MOVE_CLAN, Truced_Clans, TRIBESINHEX)
      MOVEMENT_LINE = MOVEMENT_LINE & " May Raid " & TRIBESINHEX
      If Len(TRIBESINHEX_NEW) > 3 Then
         MSG1 = "Raid Orders from " & MOVE_TRIBE & " found the following groups " & TRIBESINHEX_NEW
         Response = MsgBox((MSG1), True)
      End If
   End If

   GoTo END_FINDS:
End If

CSF_POS = "PATROL MISSION"
If SCOUT_MISSION = "PATROL" Then
   ' WHO IS BEING RAIDED
   TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
   If TRIBESINHEX = "EMPTY" Then
       MOVEMENT_LINE = MOVEMENT_LINE & " Nothing of interest found"
   Else
      Scout_Result.AddNew
      Scout_Result![TRIBE] = MOVE_TRIBE
      Scout_Result![SCOUT] = SCOUT_NUMBER
      Scout_Result![MISSION] = SCOUT_MISSION
      Scout_Result![FOUND] = TRIBESINHEX
      Scout_Result![Results] = "Patrolled and found these tribes"
      Scout_Result.UPDATE
      TRIBESINHEX_NEW = Check_Truced(MOVE_CLAN, Truced_Clans, TRIBESINHEX)
      MOVEMENT_LINE = MOVEMENT_LINE & " Patrolled and found " & TRIBESINHEX
      If Len(TRIBESINHEX_NEW) > 3 Then
         MSG1 = "Patrol Orders from " & MOVE_TRIBE & " found the following groups " & TRIBESINHEX_NEW
         Response = MsgBox((MSG1), True)
      End If
   End If

   GoTo END_FINDS:
End If

CSF_POS = "RECRUITS MISSION"
If SCOUT_MISSION = "RECRUITS" Then
   If roll1 > FIND_CHANCE Then
      MOVEMENT_LINE = MOVEMENT_LINE & ", Find no Recruits"
   Else
      roll2 = DROLL(6, 1, 10, 0, DICE_TRIBE, 1, 0)
      NUMBER_OF_RECRUITS = CLng(roll2 + Sqr(SCOUTS_USED)) * ((2 + DIPLOMACY_LEVEL) / 10)
      If NUMBER_OF_RECRUITS = 0 Then
         NUMBER_OF_RECRUITS = 1
      End If
      ' REQUIRE GENERAL ROLL FOR EVERY 3 PEOPLE
      TRIBESINFO.Edit
      TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] + CLng(NUMBER_OF_RECRUITS)
      TRIBESINFO.UPDATE
      MOVEMENT_LINE = MOVEMENT_LINE & " Find " & CLng(NUMBER_OF_RECRUITS) & " Recruits"
      AMOUNT_OF_FINDS = NUMBER_OF_RECRUITS / 3
      CURRENT_TERRAIN = hexmaptable![TERRAIN]
      If CURRENT_TERRAIN = "PRAIRIE" Then
         CURRENT_TERRAIN = "PRAIRIE"
      ElseIf CURRENT_TERRAIN = "GRASSY HILLS" Then
         CURRENT_TERRAIN = "GRASSY HILLS"
      ElseIf CURRENT_TERRAIN = "SWAMP" Then
         CURRENT_TERRAIN = "SWAMP"
      ElseIf InStr(CURRENT_TERRAIN, "TUNDRA") Then
         CURRENT_TERRAIN = "MOUNTAINS"
      ElseIf InStr(CURRENT_TERRAIN, "MOUNTAIN") Then
         CURRENT_TERRAIN = "MOUNTAINS"
      ElseIf InStr(CURRENT_TERRAIN, "ARID") Then
         CURRENT_TERRAIN = "ARID"
      ElseIf InStr(CURRENT_TERRAIN, "DESERT") Then
         CURRENT_TERRAIN = "DESERT"
      ElseIf InStr(CURRENT_TERRAIN, "SNOW") Then
         CURRENT_TERRAIN = "MOUNTAINS"
      ElseIf InStr(CURRENT_TERRAIN, "VOLCANO") Then
         CURRENT_TERRAIN = "ARID"
      ElseIf InStr(CURRENT_TERRAIN, "SNOW") Then
         CURRENT_TERRAIN = "MOUNTAINS"
      ElseIf InStr(CURRENT_TERRAIN, "JUNGLE") Then
         CURRENT_TERRAIN = "FOREST"
      ElseIf InStr(CURRENT_TERRAIN, "CONIFER") Then
         CURRENT_TERRAIN = "FOREST"
      ElseIf InStr(CURRENT_TERRAIN, "DECIDUOUS") Then
         CURRENT_TERRAIN = "FOREST"
      Else
         CURRENT_TERRAIN = "NONE"
         GoTo END_FINDS:
      End If
   
      If AMOUNT_OF_FINDS > 0 Then
         For cnt1 = 1 To AMOUNT_OF_FINDS
            Find_Roll = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
           
            'Set SCOUTING_TABLE = TVDB.OpenRecordset("SCOUTING_FINDS")
            SCOUTING_TABLE.index = "SECONDARY INDEX"
            SCOUTING_TABLE.MoveFirst
            SCOUTING_TABLE.Seek "=", CURRENT_TERRAIN
            
            Do Until (Find_Roll >= SCOUTING_TABLE![LOWEST ROLL]) And (Find_Roll <= SCOUTING_TABLE![HIGHEST ROLL])
               SCOUTING_TABLE.MoveNext
            Loop

            If SCOUTING_TABLE![TYPE OF CALC] = "SQUARE" Then
               NUMBER_OF_ITEMS = CLng(SCOUTING_TABLE![MIN NUMBER] * Sqr(SCOUTS_USED))
               Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
               MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
                   
            ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE" Then
               Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 1, 0)
               If InStr(SCOUTING_TABLE![DICE ROLLS REQ], Item_Roll) Then
                  NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
                  Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                  MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
               Else
                  SCOUTING_TABLE.MoveNext
                  If (Find_Roll >= SCOUTING_TABLE![LOWEST ROLL]) And (Find_Roll <= SCOUTING_TABLE![HIGHEST ROLL]) Then
                     If InStr(SCOUTING_TABLE![DICE ROLLS REQ], Item_Roll) Then
                        NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
                        Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                        MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
                     Else
                        SCOUTING_TABLE.MoveNext
                        If (Find_Roll >= SCOUTING_TABLE![LOWEST ROLL]) And (Find_Roll <= SCOUTING_TABLE![HIGHEST ROLL]) Then
                           If InStr(SCOUTING_TABLE![DICE ROLLS REQ], Item_Roll) Then
                              NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
                              Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                              MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
                           End If
                         End If
                      End If
                   End If
                End If
         
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE ADD" Then
                Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
                NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] + Item_Roll
                Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
          
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE MULT" Then
                Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
                NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] * Item_Roll
                Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
          
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE CALC" Then
                Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
                NUMBER_OF_ITEMS = (Item_Roll - 1) * 3 + SCOUTING_TABLE![MIN NUMBER]
                Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
            
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE BODY 1" Then
                Call UPDATE_TRIBES_TABLES("SLING", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("CLUB", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("HOOD", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("JERKIN", "ADD", 1)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find 1 Sling, Club, Hood & Jerkin "
            
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE BODY 2" Then
                Call UPDATE_TRIBES_TABLES("BONE AXE", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("BONE ARMOUR", "ADD", 1)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find 1 Bone Axe & 1 Bone Armour "
         
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "BODY 1" Then
                Call UPDATE_TRIBES_TABLES("SLING", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("CLUB", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("HOOD", "ADD", 1)
                Call UPDATE_TRIBES_TABLES("JERKIN", "ADD", 1)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find 1 Sling, Club, Hood & Jerkin "
         
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "PER SCOUT" Then
                NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] * SCOUTS_USED
                Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
           
             ElseIf SCOUTING_TABLE![TYPE OF CALC] = "PER 5 SCOUT" Then
                NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] * CLng(SCOUTS_USED / 5)
                Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
         
             ElseIf SCOUTING_TABLE![ITEM] = "SILVER" Then
                NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
                Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM] & " coins"
        
             ElseIf SCOUTING_TABLE![ITEM] = "Empty Barrel" Then
                NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
                Call UPDATE_TRIBES_TABLES("BARREL", "ADD", NUMBER_OF_ITEMS)
                MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
        
           Else
              NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
              Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
              MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
                    
           End If
        Next
     End If
        
   End If
   GoTo END_FINDS:
End If
  
If roll1 > FIND_CHANCE Then
   ' check for minerals
   Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
   HEXMAPMINERALS.index = "PRIMARYKEY"
   HEXMAPMINERALS.MoveFirst
   HEXMAPMINERALS.Seek "=", CURRENT_MAP
   If IsNull(HEXMAPMINERALS![ORE_TYPE]) Then
      MOVEMENT_LINE = MOVEMENT_LINE & " Find nothing while searching"
   End If
   GoTo END_FINDS:
End If
   
'GENERAL ROLL
roll3 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
CURRENT_TERRAIN = hexmaptable![TERRAIN]
If CURRENT_TERRAIN = "PRAIRIE" Then
   CURRENT_TERRAIN = "PRAIRIE"
ElseIf CURRENT_TERRAIN = "GRASSY HILLS" Then
   CURRENT_TERRAIN = "GRASSY HILLS"
ElseIf CURRENT_TERRAIN = "SWAMP" Then
   CURRENT_TERRAIN = "SWAMP"
ElseIf InStr(CURRENT_TERRAIN, "TUNDRA") Then
   CURRENT_TERRAIN = "TUNDRA"
ElseIf InStr(CURRENT_TERRAIN, "MOUNTAIN") Then
   CURRENT_TERRAIN = "MOUNTAINS"
ElseIf InStr(CURRENT_TERRAIN, "ARID") Then
   CURRENT_TERRAIN = "ARID"
ElseIf InStr(CURRENT_TERRAIN, "DESERT") Then
   CURRENT_TERRAIN = "DESERT"
ElseIf InStr(CURRENT_TERRAIN, "SNOW") Then
   CURRENT_TERRAIN = "MOUNTAINS"
ElseIf InStr(CURRENT_TERRAIN, "VOLCANO") Then
   CURRENT_TERRAIN = "ARID"
ElseIf InStr(CURRENT_TERRAIN, "SNOW") Then
   CURRENT_TERRAIN = "MOUNTAINS"
ElseIf InStr(CURRENT_TERRAIN, "JUNGLE") Then
   CURRENT_TERRAIN = "FOREST"
ElseIf InStr(CURRENT_TERRAIN, "CONIFER") Then
   CURRENT_TERRAIN = "FOREST"
ElseIf InStr(CURRENT_TERRAIN, "DECIDUOUS") Then
   CURRENT_TERRAIN = "FOREST"
ElseIf InStr(CURRENT_TERRAIN, "SNOW") Then
   CURRENT_TERRAIN = "MOUNTAINS"
Else
   CURRENT_TERRAIN = "NONE"
   GoTo END_FINDS:
End If
   
Find_Roll = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
           
'Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
'Set SCOUTING_TABLE = TVDB.OpenRecordset("SCOUTING_FINDS")
SCOUTING_TABLE.index = "SECONDARY INDEX"
SCOUTING_TABLE.MoveFirst
SCOUTING_TABLE.Seek "=", CURRENT_TERRAIN
           
If Not SCOUTING_TABLE.NoMatch Then
Do Until (Find_Roll >= SCOUTING_TABLE![LOWEST ROLL]) And (Find_Roll <= SCOUTING_TABLE![HIGHEST ROLL])
   SCOUTING_TABLE.MoveNext
   If SCOUTING_TABLE.EOF Then
      Exit Do
   End If
Loop

If SCOUTING_TABLE![TYPE OF CALC] = "SQUARE" Then
   NUMBER_OF_ITEMS = CLng(SCOUTING_TABLE![MIN NUMBER] * Sqr(SCOUTS_USED))
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
                  
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE" Then
   Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
   If InStr(SCOUTING_TABLE![DICE ROLLS REQ], Item_Roll) Then
      NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
      Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
      MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   Else
      SCOUTING_TABLE.MoveNext
      If (Find_Roll >= SCOUTING_TABLE![LOWEST ROLL]) And (Find_Roll <= SCOUTING_TABLE![HIGHEST ROLL]) Then
         If InStr(SCOUTING_TABLE![DICE ROLLS REQ], Item_Roll) Then
            NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
            Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
            MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
         Else
            SCOUTING_TABLE.MoveNext
            If (Find_Roll >= SCOUTING_TABLE![LOWEST ROLL]) And (Find_Roll <= SCOUTING_TABLE![HIGHEST ROLL]) Then
               If InStr(SCOUTING_TABLE![DICE ROLLS REQ], Item_Roll) Then
                  NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
                  Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
                  MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
               End If
            End If
         End If
      End If
   End If
   GoTo END_FINDS:
         
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE ADD" Then
   Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
   NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] + Item_Roll
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
        
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE MULT" Then
   Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
   NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] * Item_Roll
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
         
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE CALC" Then
   Item_Roll = DROLL(6, 1, 6, 0, DICE_TRIBE, 0, 0)
   NUMBER_OF_ITEMS = (Item_Roll - 1) * 3 + SCOUTING_TABLE![MIN NUMBER]
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
        
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE BODY 1" Then
   Call UPDATE_TRIBES_TABLES("SLING", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("CLUB", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("HOOD", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("JERKIN", "ADD", 1)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find 1 Sling, Club, Hood & Jerkin "
   GoTo END_FINDS:
         
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "DICE BODY 2" Then
   Call UPDATE_TRIBES_TABLES("BONE AXE", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("BONE ARMOUR", "ADD", 1)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find 1 Bone Axe & 1 Bone Armour "
   GoTo END_FINDS:
        
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "BODY 1" Then
   Call UPDATE_TRIBES_TABLES("SLING", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("CLUB", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("HOOD", "ADD", 1)
   Call UPDATE_TRIBES_TABLES("JERKIN", "ADD", 1)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find 1 Sling, Club, Hood & Jerkin "
   GoTo END_FINDS:
         
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "PER SCOUT" Then
   NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] * SCOUTS_USED
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
         
ElseIf SCOUTING_TABLE![TYPE OF CALC] = "PER 5 SCOUT" Then
   NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER] * CLng(SCOUTS_USED / 5)
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
         
Else
   NUMBER_OF_ITEMS = SCOUTING_TABLE![MIN NUMBER]
   Call UPDATE_TRIBES_TABLES(SCOUTING_TABLE![ITEM], "ADD", NUMBER_OF_ITEMS)
   MOVEMENT_LINE = MOVEMENT_LINE & " Find " & NUMBER_OF_ITEMS & " " & SCOUTING_TABLE![ITEM]
   GoTo END_FINDS:
                    
End If
End If

MOVEMENT_LINE = MOVEMENT_LINE & " Find nothing"

END_FINDS:

CALC_SCOUTING_FINDS_ERROR_CLOSE:
   Exit Sub


CALC_SCOUTING_FINDS_ERROR:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Msg = "Error Occurred in section " & CSF_POS
   MsgBox (Msg)
   Resume CALC_SCOUTING_FINDS_ERROR_CLOSE
End If

End Sub

Sub CAN_FLEET_MOVE(CURRENT_MAP, SCOUTS)
On Error GoTo CAN_FLEET_MOVE_ERROR

If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = CAN_GROUP_MOVE" & crlf
   Response = MsgBox((MSG1), True)
End If

START_TIME = Time

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", CURRENT_MAP

JETTY_AVAILABLE = "N"
MOVEMENT_COST = 0

If HEXMAPCONST.NoMatch Then
   JETTY_AVAILABLE = "N"
Else
   Do While HEXMAPCONST![MAP] = CURRENT_MAP
      If HEXMAPCONST![CONSTRUCTION] = "JETTY" Then
         JETTY_AVAILABLE = "Y"
         Exit Do
      End If
      HEXMAPCONST.MoveNext
   Loop
End If

TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE

   If MOVEMENT_LINE = "Tribe Movement: Move " Then
      MOVEMENT_LINE = wind & " " & WIND_DIRECTION & " Fleet Movement: Move "
   End If
  
' CALCULATE THE LAUNCHING COSTS

If Direction = "L" Then
   If SHIP_TYPE = "BOAT" Then
      If JETTY_AVAILABLE = "N" Then
         MOVEMENT_COST = MOVEMENT_COST + 1
      End If

   ElseIf Left(SHIP_TYPE, 6) = "FISHER" Then
      If JETTY_AVAILABLE = "N" Then
         MOVEMENT_COST = MOVEMENT_COST + 3
      End If
    
   ElseIf InStr(1, SHIP_TYPE, "GALLEY", 1) Then
      If JETTY_AVAILABLE = "N" Then
         MOVEMENT_COST = MOVEMENT_COST + 5
      End If
  
   ElseIf Left(SHIP_TYPE, 8) = "LONGSHIP" Then
      If JETTY_AVAILABLE = "N" Then
         MOVEMENT_COST = MOVEMENT_COST + 3
      End If

   ElseIf Left(SHIP_TYPE, 8) = "MERCHANT" Then
      If JETTY_AVAILABLE = "N" Then
         MOVEMENT_COST = MOVEMENT_COST + 5
      End If

   ElseIf Left(SHIP_TYPE, 7) = "WARSHIP" Then
      If JETTY_AVAILABLE = "N" Then
         MOVEMENT_COST = MOVEMENT_COST + 5
      End If

   End If
   TERRAIN = TERRAIN & "L "
   MOVEMENT_COUNT = MOVEMENT_COUNT + 1
   Call GET_NEXT_TRIBE_MOVE
End If

NEW_HEX_N = GET_MAP_NORTH(CURRENT_MAP)
hexmaptable.Seek "=", NEW_HEX_N
If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "LAKE" Then
   OCEAN_N = "Y"
   COAST_N = "N"
Else
   OCEAN_N = "N"
   COAST_N = "Y"
End If
NEW_HEX_NE = GET_MAP_NORTH_EAST(CURRENT_MAP)
hexmaptable.Seek "=", NEW_HEX_NE
If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "LAKE" Then
   OCEAN_NE = "Y"
   COAST_NE = "N"
Else
   OCEAN_NE = "N"
   COAST_NE = "Y"
End If
NEW_HEX_SE = GET_MAP_SOUTH_EAST(CURRENT_MAP)
hexmaptable.Seek "=", NEW_HEX_SE
If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "LAKE" Then
   OCEAN_SE = "Y"
   COAST_SE = "N"
Else
   OCEAN_SE = "N"
   COAST_SE = "Y"
End If
NEW_HEX_S = GET_MAP_SOUTH(CURRENT_MAP)
hexmaptable.Seek "=", NEW_HEX_S
If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "LAKE" Then
   OCEAN_S = "Y"
   COAST_S = "N"
Else
   OCEAN_S = "N"
   COAST_S = "Y"
End If
NEW_HEX_SW = GET_MAP_SOUTH_WEST(CURRENT_MAP)
hexmaptable.Seek "=", NEW_HEX_SW
If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "LAKE" Then
   OCEAN_SW = "Y"
   COAST_SW = "N"
Else
   OCEAN_SW = "N"
   COAST_SW = "Y"
End If
NEW_HEX_NW = GET_MAP_NORTH_WEST(CURRENT_MAP)
hexmaptable.Seek "=", NEW_HEX_NW
If hexmaptable![TERRAIN] = "OCEAN" Or hexmaptable![TERRAIN] = "LAKE" Then
   OCEAN_NW = "Y"
   COAST_NW = "N"
Else
   OCEAN_NW = "N"
   COAST_NW = "Y"
End If
hexmaptable.Seek "=", CURRENT_MAP

If Direction = "FCL" Or Direction = "FLL" Then
   If COAST_N = "N" Then
      If COAST_NW = "Y" Then
         Direction = "N"
      ElseIf COAST_SW = "Y" Then
         Direction = "NW"
      ElseIf COAST_S = "Y" Then
         Direction = "SW"
      ElseIf COAST_SE = "Y" Then
         Direction = "S"
      Else
         Direction = "SE"
      End If
   ElseIf COAST_NE = "N" Then
      Direction = "NE"
   ElseIf COAST_SE = "N" Then
      Direction = "SE"
   ElseIf COAST_S = "N" Then
      Direction = "S"
   ElseIf COAST_SW = "N" Then
      Direction = "SW"
   ElseIf COAST_NW = "N" Then
      Direction = "NW"
   Else
      NEW_ORDERS = "STOP"
   End If
ElseIf Direction = "FCR" Or Direction = "FLR" Then
   If COAST_N = "N" Then
      If COAST_NE = "Y" Then
         Direction = "N"
      ElseIf COAST_SE = "Y" Then
         Direction = "NE"
      ElseIf COAST_S = "Y" Then
         Direction = "SE"
      ElseIf COAST_SW = "Y" Then
         Direction = "S"
      Else
         Direction = "SW"
      End If
   ElseIf COAST_NW = "N" Then
      Direction = "NW"
   ElseIf COAST_SW = "N" Then
      Direction = "SW"
   ElseIf COAST_S = "N" Then
      Direction = "S"
   ElseIf COAST_SE = "N" Then
      Direction = "SE"
   ElseIf COAST_NE = "N" Then
      Direction = "NE"
   Else
      NEW_ORDERS = "STOP"
   End If
End If

If Direction = "FOL" Then
   If OCEAN_N = "N" Then
      If OCEAN_NW = "Y" Then
         Direction = "N"
      ElseIf OCEAN_SW = "Y" Then
         Direction = "NW"
      ElseIf OCEAN_S = "Y" Then
         Direction = "SW"
      ElseIf OCEAN_SE = "Y" Then
         Direction = "S"
      Else
         Direction = "SE"
      End If
   ElseIf OCEAN_NE = "N" Then
      Direction = "NE"
   ElseIf OCEAN_SE = "N" Then
      Direction = "SE"
   ElseIf OCEAN_S = "N" Then
      Direction = "S"
   ElseIf OCEAN_SW = "N" Then
      Direction = "SW"
   ElseIf OCEAN_NW = "N" Then
      Direction = "NW"
   Else
      NEW_ORDERS = "STOP"
   End If
ElseIf Direction = "FOR" Then
   If OCEAN_N = "N" Then
      If OCEAN_NE = "Y" Then
         Direction = "N"
      ElseIf OCEAN_SE = "Y" Then
         Direction = "NE"
      ElseIf OCEAN_S = "Y" Then
         Direction = "SE"
      ElseIf OCEAN_SW = "Y" Then
         Direction = "S"
      Else
         Direction = "SW"
      End If
   ElseIf OCEAN_NW = "N" Then
      Direction = "NW"
   ElseIf OCEAN_SW = "N" Then
      Direction = "SW"
   ElseIf OCEAN_S = "N" Then
      Direction = "S"
   ElseIf OCEAN_SE = "N" Then
      Direction = "SE"
   ElseIf OCEAN_NE = "N" Then
      Direction = "NE"
   Else
      NEW_ORDERS = "STOP"
   End If
End If

' INITIALISE VARIABLES

  GROUP_MOVE = "F"
  CURRENT_TERRAIN = hexmaptable![TERRAIN]
  PASS_AVAILABLE = "N"

  SAILING = "N"

' WEATHER OR WIND COSTS

   If wind = "CALM" Or wind = "NONE" Then
      If ROWING_POSSIBLE = "YES" _
      And SAILING_POSSIBLE = "YES" Then
         MOVEMENT_COST = MOVEMENT_COST + 4
         MOVEMENT_POINTS = ROWING_MOVEMENT
      ElseIf SAILING_POSSIBLE = "NO" _
      And ROWING_POSSIBLE = "YES" Then
         MOVEMENT_COST = MOVEMENT_COST + 4
         MOVEMENT_POINTS = ROWING_MOVEMENT
      ElseIf SAILING_POSSIBLE = "YES" _
      And ROWING_POSSIBLE = "NO" Then
         MOVEMENT_COST = 10000
         MOVEMENT_POINTS = SAILING_MOVEMENT
      ElseIf SAILING_POSSIBLE = "BOTH" _
      And ROWING_POSSIBLE = "BOTH" Then
         MOVEMENT_COST = 10000
         MOVEMENT_POINTS = SAILING_MOVEMENT
      Else
         'SHIT SHOULD NOT BE HERE
      End If
      
   ElseIf wind = "STRONG" Then
      If SAILING_POSSIBLE = "YES" Then
         If WIND_DIRECTION = "N" Then
            If Direction = "N" Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "NE") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "SE") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "S" Then
               MOVEMENT_COST = MOVEMENT_COST + 4
            End If
         ElseIf WIND_DIRECTION = "NE" Then
            If Direction = "NE" Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "N") Or (Direction = "SE") Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "S") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "SW" Then
               MOVEMENT_COST = MOVEMENT_COST + 4
            End If
         ElseIf WIND_DIRECTION = "SE" Then
            If Direction = "SE" Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "NE") Or (Direction = "S") Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "SW") Or (Direction = "N") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "NW" Then
               MOVEMENT_COST = MOVEMENT_COST + 4
            End If
         ElseIf WIND_DIRECTION = "S" Then
            If Direction = "S" Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "SE") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "NE") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "N" Then
               MOVEMENT_COST = MOVEMENT_COST + 4
            End If
         ElseIf WIND_DIRECTION = "SW" Then
            If Direction = "SW" Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "S") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "SE") Or (Direction = "N") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "NE" Then
               MOVEMENT_COST = MOVEMENT_COST + 4
            End If
         ElseIf WIND_DIRECTION = "NW" Then
            If Direction = "NW" Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "N") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "S") Or (Direction = "NE") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "SE" Then
               MOVEMENT_COST = MOVEMENT_COST + 4
            End If
         End If
      ElseIf ROWING_POSSIBLE = "YES" Then
         MOVEMENT_COST = MOVEMENT_COST + 5
      ElseIf ROWING_POSSIBLE = "BOTH" Then
         MOVEMENT_COST = MOVEMENT_COST + 5
      Else
         MOVEMENT_LINE = MOVEMENT_LINE & " is not possible with this ship combination"
         MOVEMENT_COST = 10000
      End If
      
   ElseIf wind = "GALE" Then
      If SAILING_POSSIBLE = "YES" Then
         If WIND_DIRECTION = "N" Then
            If Direction = "N" Then
               MOVEMENT_COST = MOVEMENT_COST + 30
            ElseIf (Direction = "NE") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "SE") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "S" Then
               MOVEMENT_COST = MOVEMENT_COST + 3
            End If
         ElseIf WIND_DIRECTION = "NE" Then
            If Direction = "NE" Then
               MOVEMENT_COST = MOVEMENT_COST + 30
            ElseIf (Direction = "N") Or (Direction = "SE") Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "S") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "SW" Then
               MOVEMENT_COST = MOVEMENT_COST + 3
            End If
         ElseIf WIND_DIRECTION = "SE" Then
            If Direction = "SE" Then
               MOVEMENT_COST = MOVEMENT_COST + 30
            ElseIf (Direction = "NE") Or (Direction = "S") Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "SW") Or (Direction = "N") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "NW" Then
               MOVEMENT_COST = MOVEMENT_COST + 3
            End If
         ElseIf WIND_DIRECTION = "S" Then
            If Direction = "S" Then
               MOVEMENT_COST = MOVEMENT_COST + 30
            ElseIf (Direction = "SE") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "NE") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "N" Then
               MOVEMENT_COST = MOVEMENT_COST + 3
            End If
         ElseIf WIND_DIRECTION = "SW" Then
            If Direction = "SW" Then
               MOVEMENT_COST = MOVEMENT_COST + 30
            ElseIf (Direction = "S") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "SE") Or (Direction = "N") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "NE" Then
               MOVEMENT_COST = MOVEMENT_COST + 3
            End If
         ElseIf WIND_DIRECTION = "NW" Then
            If Direction = "NW" Then
               MOVEMENT_COST = MOVEMENT_COST + 30
            ElseIf (Direction = "N") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 15
            ElseIf (Direction = "S") Or (Direction = "NE") Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            ElseIf Direction = "SE" Then
               MOVEMENT_COST = MOVEMENT_COST + 3
            End If
         End If
      ElseIf ROWING_POSSIBLE = "YES" Then
         MOVEMENT_COST = MOVEMENT_COST + 6
      ElseIf ROWING_POSSIBLE = "BOTH" Then
         MOVEMENT_COST = MOVEMENT_COST + 6
      Else
         MOVEMENT_LINE = MOVEMENT_LINE & " is not possible with this ship combination"
         MOVEMENT_COST = 10000
      End If
      
   ElseIf wind = "MILD" Then
      If SAILING_POSSIBLE = "YES" Then
         If WIND_DIRECTION = "N" Then
            If Direction = "N" Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "NE") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 8
            ElseIf (Direction = "SE") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 6
            ElseIf Direction = "S" Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            End If
         ElseIf WIND_DIRECTION = "NE" Then
            If Direction = "NE" Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "N") Or (Direction = "SE") Then
               MOVEMENT_COST = MOVEMENT_COST + 8
            ElseIf (Direction = "S") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 6
            ElseIf Direction = "SW" Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            End If
         ElseIf WIND_DIRECTION = "SE" Then
            If Direction = "SE" Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "NE") Or (Direction = "S") Then
               MOVEMENT_COST = MOVEMENT_COST + 8
            ElseIf (Direction = "SW") Or (Direction = "N") Then
               MOVEMENT_COST = MOVEMENT_COST + 6
            ElseIf Direction = "NW" Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            End If
         ElseIf WIND_DIRECTION = "S" Then
            If Direction = "S" Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "SE") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 8
            ElseIf (Direction = "NE") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 6
            ElseIf Direction = "N" Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            End If
         ElseIf WIND_DIRECTION = "SW" Then
            If Direction = "SW" Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "S") Or (Direction = "NW") Then
               MOVEMENT_COST = MOVEMENT_COST + 8
            ElseIf (Direction = "SE") Or (Direction = "N") Then
               MOVEMENT_COST = MOVEMENT_COST + 6
            ElseIf Direction = "NE" Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            End If
         ElseIf WIND_DIRECTION = "NW" Then
            If Direction = "NW" Then
               MOVEMENT_COST = MOVEMENT_COST + 10
            ElseIf (Direction = "N") Or (Direction = "SW") Then
               MOVEMENT_COST = MOVEMENT_COST + 8
            ElseIf (Direction = "S") Or (Direction = "NE") Then
               MOVEMENT_COST = MOVEMENT_COST + 6
            ElseIf Direction = "SE" Then
               MOVEMENT_COST = MOVEMENT_COST + 5
            End If
         End If
      ElseIf ROWING_POSSIBLE = "YES" Then
         MOVEMENT_COST = MOVEMENT_COST + 4
      ElseIf ROWING_POSSIBLE = "BOTH" Then
         MOVEMENT_COST = MOVEMENT_COST + 4
      Else
         MOVEMENT_LINE = MOVEMENT_LINE & " is not possible with this ship combination"
         MOVEMENT_COST = 10000
      End If
   End If

' DETERMINE MOVEMENT POINTS REQUIRED
' GET THE NEXT TERRAIN ( HEX MAP )
If GROUP_MOVE = "F" Then
   If Direction = "N" Then
      NEW_HEX = GET_MAP_NORTH(CURRENT_MAP)
   ElseIf Direction = "NE" Then
      NEW_HEX = GET_MAP_NORTH_EAST(CURRENT_MAP)
   ElseIf Direction = "SE" Then
      NEW_HEX = GET_MAP_SOUTH_EAST(CURRENT_MAP)
   ElseIf Direction = "S" Then
      NEW_HEX = GET_MAP_SOUTH(CURRENT_MAP)
   ElseIf Direction = "SW" Then
      NEW_HEX = GET_MAP_SOUTH_WEST(CURRENT_MAP)
   ElseIf Direction = "NW" Then
      NEW_HEX = GET_MAP_NORTH_WEST(CURRENT_MAP)
   Else
      GROUP_MOVE = "N"
      NO_MOVEMENT_REASON = "NO DIRECTION"
   End If
   hexmaptable.Seek "=", NEW_HEX
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX)
      End If
      hexmaptable.Seek "=", NEW_HEX
      NEW_TERRAIN = hexmaptable![TERRAIN]
   Else
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX)
      End If
      hexmaptable.Seek "=", NEW_HEX
      NEW_TERRAIN = hexmaptable![TERRAIN]
   End If
   NEW_HEX_N = GET_MAP_NORTH(NEW_HEX)
   hexmaptable.Seek "=", NEW_HEX_N
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX_N, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX_N
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX_N)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX_N
      End If
   End If
   N_TERRAIN = hexmaptable![TERRAIN]
   NEW_HEX_NE = GET_MAP_NORTH_EAST(NEW_HEX)
   hexmaptable.Seek "=", NEW_HEX_NE
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX_NE, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX_NE
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX_NE)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX_NE
      End If
   End If
   NE_TERRAIN = hexmaptable![TERRAIN]
   NEW_HEX_SE = GET_MAP_SOUTH_EAST(NEW_HEX)
   hexmaptable.Seek "=", NEW_HEX_SE
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX_SE, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX_SE
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX_SE)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX_SE
      End If
   End If
   SE_TERRAIN = hexmaptable![TERRAIN]
   NEW_HEX_S = GET_MAP_SOUTH(NEW_HEX)
   hexmaptable.Seek "=", NEW_HEX_S
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX_S, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX_S
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX_S)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX_S
      End If
   End If
   S_TERRAIN = hexmaptable![TERRAIN]
   NEW_HEX_SW = GET_MAP_SOUTH_WEST(NEW_HEX)
   hexmaptable.Seek "=", NEW_HEX_SW
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX_NE, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX_NE
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX_SW)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX_SW
      End If
   End If
   SW_TERRAIN = hexmaptable![TERRAIN]
   NEW_HEX_NW = GET_MAP_NORTH_WEST(NEW_HEX)
   hexmaptable.Seek "=", NEW_HEX_NW
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX_NW, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX_NW
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX_NW)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX_NW
      End If
   End If
   NW_TERRAIN = hexmaptable![TERRAIN]
   hexmaptable.Seek "=", NEW_HEX
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NEW_HEX, TERRAIN)
      hexmaptable.Seek "=", NEW_HEX
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NEW_HEX)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX
      End If
   End If
   NEW_TERRAIN = hexmaptable![TERRAIN]
   
   If FLEET = "Y" Then
      If Left(NEW_TERRAIN, 5) = "OCEAN" Or Left(NEW_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      ElseIf Left(N_TERRAIN, 5) = "OCEAN" Or Left(N_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      ElseIf Left(NE_TERRAIN, 5) = "OCEAN" Or Left(NE_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      ElseIf Left(NW_TERRAIN, 5) = "OCEAN" Or Left(NW_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      ElseIf Left(S_TERRAIN, 5) = "OCEAN" Or Left(S_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      ElseIf Left(SE_TERRAIN, 5) = "OCEAN" Or Left(SE_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      ElseIf Left(SW_TERRAIN, 5) = "OCEAN" Or Left(SW_TERRAIN, 4) = "LAKE" Then
         GROUP_MOVE = "Y"
      Else      '   ALLOW FOR RIVER TRAVEL
                '   NOT FINISHED
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NEW_HEX
         ' NEED TO FIND RIVERS
         If Mid(hexmaptable![Borders], 1, 2) = "RI" Or Mid(hexmaptable![Borders], 1, 2) = "FO" Or Mid(hexmaptable![Borders], 1, 2) = "CA" Then
            GROUP_MOVE = "Y"
         ElseIf Mid(hexmaptable![Borders], 3, 2) = "RI" Or Mid(hexmaptable![Borders], 3, 2) = "FO" Or Mid(hexmaptable![Borders], 3, 2) = "CA" Then
            GROUP_MOVE = "Y"
         ElseIf Mid(hexmaptable![Borders], 5, 2) = "RI" Or Mid(hexmaptable![Borders], 5, 2) = "FO" Or Mid(hexmaptable![Borders], 5, 2) = "CA" Then
            GROUP_MOVE = "Y"
         ElseIf Mid(hexmaptable![Borders], 7, 2) = "RI" Or Mid(hexmaptable![Borders], 7, 2) = "FO" Or Mid(hexmaptable![Borders], 7, 2) = "CA" Then
            GROUP_MOVE = "Y"
         ElseIf Mid(hexmaptable![Borders], 9, 2) = "RI" Or Mid(hexmaptable![Borders], 9, 2) = "FO" Or Mid(hexmaptable![Borders], 9, 2) = "CA" Then
            GROUP_MOVE = "Y"
         ElseIf Mid(hexmaptable![Borders], 11, 2) = "RI" Or Mid(hexmaptable![Borders], 11, 2) = "FO" Or Mid(hexmaptable![Borders], 11, 2) = "CA" Then
            GROUP_MOVE = "Y"
         Else
            GROUP_MOVE = "N"
         End If
         If GROUP_MOVE = "N" Then
            NO_MOVEMENT_REASON = "No River Adjacent to Hex to " & Direction
            NO_MOVEMENT_REASON = NO_MOVEMENT_REASON & " of HEX "
         End If
      End If
   End If

   TERRAINTABLE.MoveFirst
   TERRAINTABLE.Seek "=", NEW_TERRAIN
   ' LOOK AT FURTHER WITH REGARDS TO PASSES.

End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If (GROUP_MOVE = "F") Or (GROUP_MOVE = "Y") Then
   If MOVEMENT_COST > MOVEMENT_POINTS Then
      NO_MOVEMENT_REASON = "Not enough M.P's"
      GROUP_MOVE = "N"
      SLENGTH = 0
      SPOSTION = 100
   Else
      MOVEMENT_POINTS = MOVEMENT_POINTS - MOVEMENT_COST
      GROUP_MOVE = "Y"
   End If
End If

END_TIME = Time

   If codetrack = 1 Then
      MSG0 = "GROUP_MOVE = " & GROUP_MOVE & crlf
      MSG1 = "MOVEMENT_COST = " & MOVEMENT_COST & crlf
      MSG2 = "MOVEMENT_POINTS = " & MOVEMENT_POINTS & crlf
      Response = MsgBox((MSG0 & MSG1 & MSG2), True)
   End If

CAN_FLEET_MOVE_ERROR_CLOSE:
   Exit Sub


CAN_FLEET_MOVE_ERROR:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Resume CAN_FLEET_MOVE_ERROR_CLOSE
End If


End Sub

Sub CAN_GROUP_MOVE(CURRENT_MAP, SCOUTS)
On Error GoTo CAN_GROUP_MOVE_ERROR

Dim MOVEMENT_COST As Long
Dim NEW_HEX As String
Dim NEW_TERRAIN As String
Dim CURRENT_TERRAIN As String
Dim PASS_AVAILABLE As String
   
If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = CAN_GROUP_MOVE" & crlf
   Response = MsgBox((MSG1), True)
End If

START_TIME = Time

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

MOVEMENT_COST = 0

TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE

' INITIALISE VARIABLES

  GROUP_MOVE = "F"
  CURRENT_TERRAIN = hexmaptable![TERRAIN]
  PASS_AVAILABLE = "N"

' WEATHER OR WIND COSTS

   If (Left(WEATHER, 1) = "L") Or (Left(WEATHER, 4) = "WIND") Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf Left(WEATHER, 1) = "H" Then
      MOVEMENT_COST = MOVEMENT_COST + 2
   End If

' GET RIVERS
   RIVER_N = Mid(hexmaptable![Borders], 1, 2)
   RIVER_NE = Mid(hexmaptable![Borders], 3, 2)
   RIVER_SE = Mid(hexmaptable![Borders], 5, 2)
   RIVER_S = Mid(hexmaptable![Borders], 7, 2)
   RIVER_SW = Mid(hexmaptable![Borders], 9, 2)
   RIVER_NW = Mid(hexmaptable![Borders], 11, 2)
   If RIVER_N = "FO" Or RIVER_N = "CA" Then
      RIVER_N = "RI"
   End If
   If RIVER_NE = "FO" Or RIVER_NE = "CA" Then
      RIVER_NE = "RI"
   End If
   If RIVER_SE = "FO" Or RIVER_SE = "CA" Then
      RIVER_SE = "RI"
   End If
   If RIVER_S = "FO" Or RIVER_S = "CA" Then
      RIVER_S = "RI"
   End If
   If RIVER_SW = "FO" Or RIVER_SW = "CA" Then
      RIVER_SW = "RI"
   End If
   If RIVER_NW = "FO" Or RIVER_NW = "CA" Then
      RIVER_NW = "RI"
   End If
   FORD_N = Mid(hexmaptable![Borders], 1, 2)
   FORD_NE = Mid(hexmaptable![Borders], 3, 2)
   FORD_SE = Mid(hexmaptable![Borders], 5, 2)
   FORD_S = Mid(hexmaptable![Borders], 7, 2)
   FORD_SW = Mid(hexmaptable![Borders], 9, 2)
   FORD_NW = Mid(hexmaptable![Borders], 11, 2)
   PASS_N = Mid(hexmaptable![Borders], 1, 2)
   PASS_NE = Mid(hexmaptable![Borders], 3, 2)
   PASS_SE = Mid(hexmaptable![Borders], 5, 2)
   PASS_S = Mid(hexmaptable![Borders], 7, 2)
   PASS_SW = Mid(hexmaptable![Borders], 9, 2)
   PASS_NW = Mid(hexmaptable![Borders], 11, 2)
   NEW_HEX_N = GET_MAP_NORTH(CURRENT_MAP)
   hexmaptable.Seek "=", NEW_HEX_N
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_N)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_N
   End If
   If hexmaptable![TERRAIN] = "OCEAN" Then
      OCEAN_N = "Y"
      LAKE_N = "N"
      MOUNTAIN_N = "N"
   ElseIf hexmaptable![TERRAIN] = "LAKE" Then
      OCEAN_N = "Y"
      LAKE_N = "Y"
      MOUNTAIN_N = "N"
   ElseIf InStr(hexmaptable![TERRAIN], "MOUNTAIN") Or InStr(hexmaptable![TERRAIN], "MT") Then
      OCEAN_N = "N"
      LAKE_N = "N"
      MOUNTAIN_N = "Y"
   Else
      OCEAN_N = "N"
      LAKE_N = "N"
      MOUNTAIN_N = "N"
   End If
   NEW_HEX_NE = GET_MAP_NORTH_EAST(CURRENT_MAP)
   hexmaptable.Seek "=", NEW_HEX_NE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NE
   End If
   If hexmaptable![TERRAIN] = "OCEAN" Then
      OCEAN_NE = "Y"
      LAKE_NE = "N"
      MOUNTAIN_NE = "N"
   ElseIf hexmaptable![TERRAIN] = "LAKE" Then
      OCEAN_NE = "Y"
      LAKE_NE = "Y"
      MOUNTAIN_NE = "N"
   ElseIf InStr(hexmaptable![TERRAIN], "MOUNTAIN") Or InStr(hexmaptable![TERRAIN], "MT") Then
      OCEAN_NE = "N"
      LAKE_NE = "N"
      MOUNTAIN_NE = "Y"
   Else
      OCEAN_NE = "N"
      LAKE_NE = "N"
      MOUNTAIN_NE = "N"
   End If
   NEW_HEX_SE = GET_MAP_SOUTH_EAST(CURRENT_MAP)
   hexmaptable.Seek "=", NEW_HEX_SE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SE
   End If
   If hexmaptable![TERRAIN] = "OCEAN" Then
      OCEAN_SE = "Y"
      LAKE_SE = "N"
      MOUNTAIN_SE = "N"
   ElseIf hexmaptable![TERRAIN] = "LAKE" Then
      OCEAN_SE = "Y"
      LAKE_SE = "Y"
      MOUNTAIN_SE = "N"
   ElseIf InStr(hexmaptable![TERRAIN], "MOUNTAIN") Or InStr(hexmaptable![TERRAIN], "MT") Then
      OCEAN_SE = "N"
      LAKE_SE = "N"
      MOUNTAIN_SE = "Y"
   Else
      OCEAN_SE = "N"
      LAKE_SE = "N"
      MOUNTAIN_SE = "N"
   End If
   NEW_HEX_S = GET_MAP_SOUTH(CURRENT_MAP)
   hexmaptable.Seek "=", NEW_HEX_S
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_S)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_S
   End If
   If hexmaptable![TERRAIN] = "OCEAN" Then
      OCEAN_S = "Y"
      LAKE_S = "N"
      MOUNTAIN_S = "N"
   ElseIf hexmaptable![TERRAIN] = "LAKE" Then
      OCEAN_S = "Y"
      LAKE_S = "Y"
      MOUNTAIN_S = "N"
   ElseIf InStr(hexmaptable![TERRAIN], "MOUNTAIN") Or InStr(hexmaptable![TERRAIN], "MT") Then
      OCEAN_S = "N"
      LAKE_S = "N"
      MOUNTAIN_S = "Y"
   Else
      OCEAN_S = "N"
      LAKE_S = "N"
      MOUNTAIN_S = "N"
   End If
   NEW_HEX_SW = GET_MAP_SOUTH_WEST(CURRENT_MAP)
   hexmaptable.Seek "=", NEW_HEX_SW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SW
   End If
   If hexmaptable![TERRAIN] = "OCEAN" Then
      OCEAN_SW = "Y"
      LAKE_SW = "N"
      MOUNTAIN_SW = "N"
   ElseIf hexmaptable![TERRAIN] = "LAKE" Then
      OCEAN_SW = "Y"
      LAKE_SW = "Y"
      MOUNTAIN_SW = "N"
   ElseIf InStr(hexmaptable![TERRAIN], "MOUNTAIN") Or InStr(hexmaptable![TERRAIN], "MT") Then
      OCEAN_SW = "N"
      LAKE_SW = "N"
      MOUNTAIN_SW = "Y"
   Else
      OCEAN_SW = "N"
      LAKE_SW = "N"
      MOUNTAIN_SW = "N"
   End If
   NEW_HEX_NW = GET_MAP_NORTH_WEST(CURRENT_MAP)
   hexmaptable.Seek "=", NEW_HEX_NW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NW
   End If
   If hexmaptable![TERRAIN] = "OCEAN" Then
      OCEAN_NW = "Y"
      LAKE_NW = "N"
      MOUNTAIN_NW = "N"
   ElseIf hexmaptable![TERRAIN] = "LAKE" Then
      OCEAN_NW = "Y"
      LAKE_NW = "Y"
      MOUNTAIN_NW = "N"
   ElseIf InStr(hexmaptable![TERRAIN], "MOUNTAIN") Or InStr(hexmaptable![TERRAIN], "MT") Then
      OCEAN_NW = "N"
      LAKE_NW = "N"
      MOUNTAIN_NW = "Y"
   Else
      OCEAN_NW = "N"
      LAKE_NW = "N"
      MOUNTAIN_NW = "N"
   End If

' IN HERE DO THE CODE FOR FOLLOWING - OCEAN, LAKE, RIVER, MOUNTAIN

If Direction = "FRL" Then
   If Not RIVER_N = "RI" Then
      If RIVER_NW = "RI" Then
         Direction = "N"
      ElseIf RIVER_SW = "RI" Then
         Direction = "NW"
      ElseIf RIVER_S = "RI" Then
         Direction = "SW"
      ElseIf RIVER_SE = "RI" Then
         Direction = "S"
      Else
         Direction = "SE"
      End If
   ElseIf Not RIVER_NE = "RI" Then
      Direction = "NE"
   ElseIf Not RIVER_SE = "RI" Then
      Direction = "SE"
   ElseIf Not RIVER_S = "RI" Then
      Direction = "S"
   ElseIf Not RIVER_SW = "RI" Then
      Direction = "SW"
   ElseIf Not RIVER_NW = "RI" Then
      Direction = "NW"
   Else
      NEW_ORDERS = "STOP"
   End If
ElseIf Direction = "FRR" Then
   If Not RIVER_N = "RI" Then
      If RIVER_NE = "RI" Then
         Direction = "N"
      ElseIf RIVER_SE = "RI" Then
         Direction = "NE"
      ElseIf RIVER_S = "RI" Then
         Direction = "SE"
      ElseIf RIVER_SW = "RI" Then
         Direction = "S"
      Else
         Direction = "SW"
      End If
   ElseIf Not RIVER_NW = "RI" Then
      Direction = "NW"
   ElseIf Not RIVER_SW = "RI" Then
      Direction = "SW"
   ElseIf Not RIVER_S = "RI" Then
      Direction = "S"
   ElseIf Not RIVER_SE = "RI" Then
      Direction = "SE"
   ElseIf Not RIVER_NE = "RI" Then
      Direction = "NE"
   Else
      NEW_ORDERS = "STOP"
   End If
End If

If Direction = "FML" Then
   If MOUNTAIN_N = "N" Then
      If MOUNTAIN_NW = "Y" Then
         Direction = "N"
      ElseIf MOUNTAIN_SW = "Y" Then
         Direction = "NW"
      ElseIf MOUNTAIN_S = "Y" Then
         Direction = "SW"
      ElseIf MOUNTAIN_SE = "Y" Then
         Direction = "S"
      Else
         Direction = "SE"
      End If
   ElseIf MOUNTAIN_NE = "N" Then
      Direction = "NE"
   ElseIf MOUNTAIN_SE = "N" Then
      Direction = "SE"
   ElseIf MOUNTAIN_S = "N" Then
      Direction = "S"
   ElseIf MOUNTAIN_SW = "N" Then
      Direction = "SW"
   ElseIf MOUNTAIN_NW = "N" Then
      Direction = "NW"
   Else
      NEW_ORDERS = "STOP"
   End If
ElseIf Direction = "FMR" Then
   If MOUNTAIN_N = "N" Then
      If MOUNTAIN_NE = "Y" Then
         Direction = "N"
      ElseIf MOUNTAIN_SE = "Y" Then
         Direction = "NE"
      ElseIf MOUNTAIN_S = "Y" Then
         Direction = "SE"
      ElseIf MOUNTAIN_SW = "Y" Then
         Direction = "S"
      Else
         Direction = "SW"
      End If
   ElseIf MOUNTAIN_NW = "N" Then
      Direction = "NW"
   ElseIf MOUNTAIN_SW = "N" Then
      Direction = "SW"
   ElseIf MOUNTAIN_S = "N" Then
      Direction = "S"
   ElseIf MOUNTAIN_SE = "N" Then
      Direction = "SE"
   ElseIf MOUNTAIN_NE = "N" Then
      Direction = "NE"
   Else
      NEW_ORDERS = "STOP"
   End If
End If

If Direction = "FCL" Or Direction = "FLL" Or Direction = "FOL" Then
   If OCEAN_N = "N" Then
      If OCEAN_NW = "Y" Then
         Direction = "N"
      ElseIf OCEAN_SW = "Y" Then
         Direction = "NW"
      ElseIf OCEAN_S = "Y" Then
         Direction = "SW"
      ElseIf OCEAN_SE = "Y" Then
         Direction = "S"
      Else
         Direction = "SE"
      End If
   ElseIf OCEAN_NE = "N" Then
      Direction = "NE"
   ElseIf OCEAN_SE = "N" Then
      Direction = "SE"
   ElseIf OCEAN_S = "N" Then
      Direction = "S"
   ElseIf OCEAN_SW = "N" Then
      Direction = "SW"
   ElseIf OCEAN_NW = "N" Then
      Direction = "NW"
   Else
      NEW_ORDERS = "STOP"
   End If
ElseIf Direction = "FCR" Or Direction = "FLR" Or Direction = "FOR" Then
   If OCEAN_N = "N" Then
      If OCEAN_NE = "Y" Then
         Direction = "N"
      ElseIf OCEAN_SE = "Y" Then
         Direction = "NE"
      ElseIf OCEAN_S = "Y" Then
         Direction = "SE"
      ElseIf OCEAN_SW = "Y" Then
         Direction = "S"
      Else
         Direction = "SW"
      End If
   ElseIf OCEAN_NW = "N" Then
      Direction = "NW"
   ElseIf OCEAN_SW = "N" Then
      Direction = "SW"
   ElseIf OCEAN_S = "N" Then
      Direction = "S"
   ElseIf OCEAN_SE = "N" Then
      Direction = "SE"
   ElseIf OCEAN_NE = "N" Then
      Direction = "NE"
   Else
      NEW_ORDERS = "STOP"
   End If
End If

' CAN THE GROUP MOVE INTO THE NEXT HEX
If Direction = "N" Then
   If FORD_N = "FO" Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf RIVER_N = "RI" Then
      GROUP_MOVE = "N"
      NO_MOVEMENT_REASON = "No Ford on River to " & Direction & " of HEX"
   End If
   If PASS_N = "PA" Then
      PASS_AVAILABLE = "Y"
   End If
ElseIf Direction = "NE" Then
   If FORD_NE = "FO" Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf RIVER_NE = "RI" Then
      NO_MOVEMENT_REASON = "No Ford on River to " & Direction & " of HEX"
      GROUP_MOVE = "N"
   End If
   If PASS_NE = "PA" Then
      PASS_AVAILABLE = "Y"
   End If
ElseIf Direction = "SE" Then
   If FORD_SE = "FO" Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf RIVER_SE = "RI" Then
      NO_MOVEMENT_REASON = "No Ford on River to " & Direction & " of HEX"
      GROUP_MOVE = "N"
   End If
   If PASS_SE = "PA" Then
      PASS_AVAILABLE = "Y"
   End If
ElseIf Direction = "S" Then
   If FORD_S = "FO" Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf RIVER_S = "RI" Then
      NO_MOVEMENT_REASON = "No Ford on River to " & Direction & " of HEX"
      GROUP_MOVE = "N"
   End If
   If PASS_S = "PA" Then
      PASS_AVAILABLE = "Y"
   End If
ElseIf Direction = "SW" Then
   If FORD_SW = "FO" Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf RIVER_SW = "RI" Then
      NO_MOVEMENT_REASON = "No Ford on River to " & Direction & " of HEX"
      GROUP_MOVE = "N"
   End If
   If PASS_SW = "PA" Then
      PASS_AVAILABLE = "Y"
   End If
ElseIf Direction = "NW" Then
   If FORD_NW = "FO" Then
      MOVEMENT_COST = MOVEMENT_COST + 1
   ElseIf RIVER_NW = "RI" Then
      NO_MOVEMENT_REASON = "No Ford on River to " & Direction & " of HEX"
      GROUP_MOVE = "N"
   End If
   If PASS_NW = "PA" Then
      PASS_AVAILABLE = "Y"
   End If
End If

' DETERMINE MOVEMENT POINTS REQUIRED
' GET THE NEXT TERRAIN ( HEX MAP )
If GROUP_MOVE = "F" Then
   If Direction = "N" Then
      NEW_HEX = GET_MAP_NORTH(CURRENT_MAP)
'      MsgBox ("DIRECTION = N")
   ElseIf Direction = "NE" Then
      NEW_HEX = GET_MAP_NORTH_EAST(CURRENT_MAP)
'      MsgBox ("DIRECTION = NE")
   ElseIf Direction = "SE" Then
      NEW_HEX = GET_MAP_SOUTH_EAST(CURRENT_MAP)
'      MsgBox ("DIRECTION = SE")
   ElseIf Direction = "S" Then
      NEW_HEX = GET_MAP_SOUTH(CURRENT_MAP)
'      MsgBox ("DIRECTION = S")
   ElseIf Direction = "SW" Then
      NEW_HEX = GET_MAP_SOUTH_WEST(CURRENT_MAP)
'      MsgBox ("DIRECTION = SW")
   ElseIf Direction = "NW" Then
      NEW_HEX = GET_MAP_NORTH_WEST(CURRENT_MAP)
'      MsgBox ("DIRECTION = NW")
   Else
      GROUP_MOVE = "N"
      NO_MOVEMENT_REASON = "NO DIRECTION"
      'Should exit function
   End If
   If GROUP_MOVE = "N" Then
      'no action
   Else
      hexmaptable.Seek "=", NEW_HEX
      If hexmaptable.NoMatch Then
         Call ADD_NEW_HEX(NEW_HEX, TERRAIN)
         hexmaptable.Seek "=", NEW_HEX
         If IsNull(hexmaptable![TERRAIN]) Then
            Call UPDATE_HEX_MAP(NEW_HEX)
         End If
         hexmaptable.Seek "=", NEW_HEX
         NEW_TERRAIN = hexmaptable![TERRAIN]
      Else
         If IsNull(hexmaptable![TERRAIN]) Then
            Call UPDATE_HEX_MAP(NEW_HEX)
         End If
         hexmaptable.Seek "=", NEW_HEX
         NEW_TERRAIN = hexmaptable![TERRAIN]
     End If
   End If
   
   If (Left(NEW_TERRAIN, 5) = "OCEAN") Then
         NO_MOVEMENT_REASON = "Can't Move on Ocean to " & Direction
         NO_MOVEMENT_REASON = NO_MOVEMENT_REASON & " of HEX"
         GROUP_MOVE = "N"
   ElseIf (Left(NEW_TERRAIN, 4) = "LAKE") Then
         NO_MOVEMENT_REASON = "Can't Move on Lake to " & Direction
         NO_MOVEMENT_REASON = NO_MOVEMENT_REASON & " of HEX"
         GROUP_MOVE = "N"
   ElseIf Left(NEW_TERRAIN, 8) = "MANGROVE" Then
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "ANIMAL", "Horse"
      If Not TRIBESGOODS.NoMatch Then
         If TRIBESGOODS![ITEM_NUMBER] > 0 Then
            NO_MOVEMENT_REASON = "Horses not allowed into MANGROVE Swamp to " & Direction & " of HEX"
            GROUP_MOVE = "N"
         Else
            GROUP_MOVE = "Y"
         End If
      Else
         GROUP_MOVE = "Y"
      End If
   ElseIf (Right(NEW_TERRAIN, 9) = "MOUNTAINS") Or (Right(NEW_TERRAIN, 2) = "MT") Then
      If Not SCOUTS = "Y" Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "FINISHED", "WAGON"
         If Not TRIBESGOODS.NoMatch Then
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               NUMBER_OF_WAGONS = TRIBESGOODS![ITEM_NUMBER]
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "ANIMAL", "ELEPHANT"
               If TRIBESGOODS.NoMatch Then
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Mountains  to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               ElseIf TRIBESGOODS![ITEM_NUMBER] >= NUMBER_OF_WAGONS Then
                  GROUP_MOVE = "Y"
               Else
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Mountains  to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               End If
            Else
               NUMBER_OF_WAGONS = 0
            End If
        Else
           GROUP_MOVE = "Y"
        End If
      End If
   ElseIf Left(NEW_TERRAIN, 5) = "SWAMP" Then
      If Not SCOUTS = "Y" Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "FINISHED", "WAGON"
         If Not TRIBESGOODS.NoMatch Then
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               NUMBER_OF_WAGONS = TRIBESGOODS![ITEM_NUMBER]
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "ANIMAL", "ELEPHANT"
               If TRIBESGOODS.NoMatch Then
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Swamp  to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               ElseIf TRIBESGOODS![ITEM_NUMBER] >= NUMBER_OF_WAGONS Then
                  GROUP_MOVE = "Y"
               Else
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Swamp  to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               End If
            Else
               NUMBER_OF_WAGONS = 0
            End If
         Else
            GROUP_MOVE = "Y"
         End If
      End If
   ElseIf Left(NEW_TERRAIN, 11) = "SNOWY HILLS" Then
      If Not SCOUTS = "Y" Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "FINISHED", "WAGON"
         If Not TRIBESGOODS.NoMatch Then
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               NUMBER_OF_WAGONS = TRIBESGOODS![ITEM_NUMBER]
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "ANIMAL", "ELEPHANT"
               If TRIBESGOODS.NoMatch Then
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Snowy hills to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               ElseIf TRIBESGOODS![ITEM_NUMBER] >= NUMBER_OF_WAGONS Then
                  GROUP_MOVE = "Y"
               Else
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Snowy hills to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               End If
            Else
               NUMBER_OF_WAGONS = 0
            End If
         Else
            GROUP_MOVE = "Y"
         End If
      End If
    ElseIf Left(NEW_TERRAIN, 6) = "JUNGLE" Then
      If NEW_TERRAIN = "JUNGLE" Then
          'DO NOTHING
      ElseIf Not SCOUTS = "Y" Then
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "FINISHED", "WAGON"
         If Not TRIBESGOODS.NoMatch Then
            If TRIBESGOODS![ITEM_NUMBER] > 0 Then
               NUMBER_OF_WAGONS = TRIBESGOODS![ITEM_NUMBER]
               TRIBESGOODS.MoveFirst
               TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "ANIMAL", "ELEPHANT"
               If TRIBESGOODS.NoMatch Then
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Jungle Hill to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               ElseIf TRIBESGOODS![ITEM_NUMBER] >= NUMBER_OF_WAGONS Then
                  GROUP_MOVE = "Y"
               Else
                  NO_MOVEMENT_REASON = "Cannot Move Wagons into Jungle Hill to " & Direction & " of HEX"
                  GROUP_MOVE = "N"
               End If
            Else
               NUMBER_OF_WAGONS = 0
            End If
         Else
            GROUP_MOVE = "Y"
         End If
      End If
   End If
   
   Dim weight_buffer As Double
    weight_buffer = 1.1
   If TRIBES_WEIGHT > Walking_Capacity * weight_buffer Then
          If Not SCOUTS = "Y" Then
              NO_MOVEMENT_REASON = "Insufficient capacity to carry"
              GROUP_MOVE = "N"
              
          End If
   ElseIf TRIBES_WEIGHT > TRIBES_CAPACITY * weight_buffer Then
          If TRIBES_WEIGHT < Walking_Capacity * weight_buffer Then
              ' can move
          Else
               If Not SCOUTS = "Y" Then
                   NO_MOVEMENT_REASON = "Insufficient capacity to carry"
                   GROUP_MOVE = "N"
                   
               End If
          End If
   End If
   
   TERRAINTABLE.MoveFirst
   TERRAINTABLE.Seek "=", NEW_TERRAIN
   ' LOOK AT FURTHER WITH REGARDS TO PASSES.

      If (GROUP_MOVE = "Y") Or (GROUP_MOVE = "F") Then
         If PASS_AVAILABLE = "Y" Then
            If Left(TERRAINTABLE![TERRAIN], 3) = "LOW" Then
               MOVEMENT_COST = MOVEMENT_COST + 7
            Else
               MOVEMENT_COST = MOVEMENT_COST + 8
            End If
         Else
            MOVEMENT_COST = MOVEMENT_COST + TERRAINTABLE![MOVEMENT POINTS]
         End If
      End If
   End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If Direction = "N" Then
   If Mid(hexmaptable![ROADS], 1, 1) = "R" Then
      MOVEMENT_COST = ((MOVEMENT_COST / 6) * 5)
   ElseIf Mid(hexmaptable![ROADS], 1, 1) = "D" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.66
   ElseIf Mid(hexmaptable![ROADS], 1, 1) = "S" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.33
   End If
ElseIf Direction = "NE" Then
   If Mid(hexmaptable![ROADS], 2, 1) = "R" Then
      MOVEMENT_COST = ((MOVEMENT_COST / 6) * 5)
   ElseIf Mid(hexmaptable![ROADS], 2, 1) = "D" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.66
   ElseIf Mid(hexmaptable![ROADS], 2, 1) = "S" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.33
   End If
ElseIf Direction = "SE" Then
   If Mid(hexmaptable![ROADS], 3, 1) = "R" Then
      MOVEMENT_COST = ((MOVEMENT_COST / 6) * 5)
   ElseIf Mid(hexmaptable![ROADS], 3, 1) = "D" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.66
   ElseIf Mid(hexmaptable![ROADS], 3, 1) = "S" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.33
   End If
ElseIf Direction = "S" Then
   If Mid(hexmaptable![ROADS], 4, 1) = "R" Then
      MOVEMENT_COST = ((MOVEMENT_COST / 6) * 5)
   ElseIf Mid(hexmaptable![ROADS], 4, 1) = "D" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.66
   ElseIf Mid(hexmaptable![ROADS], 4, 1) = "S" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.33
   End If
ElseIf Direction = "SW" Then
   If Mid(hexmaptable![ROADS], 5, 1) = "R" Then
      MOVEMENT_COST = ((MOVEMENT_COST / 6) * 5)
   ElseIf Mid(hexmaptable![ROADS], 5, 1) = "D" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.66
   ElseIf Mid(hexmaptable![ROADS], 5, 1) = "S" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.33
   End If
ElseIf Direction = "NW" Then
   If Mid(hexmaptable![ROADS], 6, 1) = "R" Then
      MOVEMENT_COST = ((MOVEMENT_COST / 6) * 5)
   ElseIf Mid(hexmaptable![ROADS], 6, 1) = "D" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.66
   ElseIf Mid(hexmaptable![ROADS], 6, 1) = "S" Then
      MOVEMENT_COST = MOVEMENT_COST * 0.33
   End If
End If

If (GROUP_MOVE = "F") Or (GROUP_MOVE = "Y") Then
   If MOVEMENT_COST > MOVEMENT_POINTS Then
      NO_MOVEMENT_REASON = "Not enough M.P's to move to " & Direction & " into " & NEW_TERRAIN
      GROUP_MOVE = "N"
      SLENGTH = 0
      SPOSTION = 100
   Else
      MOVEMENT_POINTS = MOVEMENT_POINTS - MOVEMENT_COST
      GROUP_MOVE = "Y"
   End If
ElseIf GROUP_MOVE = "N" Then
   'DO NOTHING
Else
   MSG1 = "GROUP_MOVE = " & GROUP_MOVE
   MSG2 = "INVALID GROUP_MOVE"
   MsgBox (MSG1 & MSG2)
End If

TM_POS = "Can Group Move"
Movement_Trace.Edit
Movement_Trace![Direction] = Direction
Movement_Trace![Target_Hex] = NEW_HEX
Movement_Trace![Target_Terrain] = NEW_TERRAIN
Movement_Trace![CAN_GROUP_MOVE] = GROUP_MOVE
Movement_Trace![Current_Movement_Cost] = MOVEMENT_COST
Movement_Trace![NO_MOVEMENT_REASON] = NO_MOVEMENT_REASON
Movement_Trace.UPDATE

END_TIME = Time

   If codetrack = 1 Then
      MSG0 = "GROUP_MOVE = " & GROUP_MOVE & crlf
      MSG1 = "MOVEMENT_COST = " & MOVEMENT_COST & crlf
      MSG2 = "MOVEMENT_POINTS = " & MOVEMENT_POINTS & crlf
      Response = MsgBox((MSG0 & MSG1 & MSG2), True)
   End If

CAN_GROUP_MOVE_ERROR_CLOSE:
   Exit Sub


CAN_GROUP_MOVE_ERROR:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Resume CAN_GROUP_MOVE_ERROR_CLOSE
End If

End Sub

Function CHECK_TERRAIN(TERRAIN)
If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = CHECK_TERRAIN" & crlf
   Response = MsgBox((MSG1), True)
End If

If TERRAIN = "OCEAN" Then
   CHECK_TERRAIN = "N"
ElseIf TERRAIN = "BAMBOO" Then
   CHECK_TERRAIN = "N"
ElseIf TERRAIN = "BRUSH" Then
   CHECK_TERRAIN = "N"
ElseIf TERRAIN = "ROCKY" Then
   CHECK_TERRAIN = "N"
ElseIf TERRAIN = "JUNGLE" Then
   CHECK_TERRAIN = "N"
ElseIf TERRAIN = "LAKE" Then
   CHECK_TERRAIN = "N"
ElseIf InStr(TERRAIN, "SWAMP") Then
   CHECK_TERRAIN = "N"
ElseIf TERRAIN = "TUNDRA" Then
   CHECK_TERRAIN = "N"
ElseIf InStr(TERRAIN, "TUNDRA") Then
   CHECK_TERRAIN = "N"
ElseIf InStr(TERRAIN, "MOUNTAIN") Then
   CHECK_TERRAIN = "N"
ElseIf InStr(TERRAIN, "ARID") Then
   CHECK_TERRAIN = "N"
ElseIf InStr(TERRAIN, "DESERT") Then
   CHECK_TERRAIN = "N"
Else
   CHECK_TERRAIN = "Y"
End If


End Function

Sub DETERMINE_MOVEMENT_POINTS(SCOUTS)
Dim VESSEL_MOVEMENT As Long
Dim VESSEL_MOVEMENT_ROW As Long
Dim VESSEL_MOVEMENT_SAIL As Long
Dim NUMBER_OF_WAGONS As Long
Dim Number_Of_Elephants As Long
Dim Number_Of_Cattle As Long
Dim NUMBER_OF_CHARIOTS As Long
Dim NUMBER_OF_LIGHTHORSE As Long
Dim NUMBER_OF_HEAVYHORSE As Long

If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = DETERMINE_MOVEMENT_POINTS" & crlf
   Response = MsgBox((MSG1), True)
End If

VESSEL_MOVEMENT = 0
MOVEMENT_POINTS = 10000
ROWING_MOVEMENT = 0
SAILING_MOVEMENT = 0
SAILING_POSSIBLE = "MAYBE"
ROWING_POSSIBLE = "MAYBE"
SAILING_ONLY = "NO"
ROWING_ONLY = "NO"

If FLEET = "Y" Then
   SHIPSTABLE.MoveFirst

   Do Until SHIPSTABLE.EOF
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "SHIP", SHIPSTABLE![VESSEL]
      If Not TRIBESGOODS.NoMatch Then
         If TRIBESGOODS![ITEM_NUMBER] > 0 Then
            VESSEL_MOVEMENT_ROW = SHIPSTABLE![BASE_MP_ROW]
            VESSEL_MOVEMENT_SAIL = SHIPSTABLE![BASE_MP_SAIL]
            VESSEL_MOVEMENT_ROW = VESSEL_MOVEMENT_ROW + (SHIPSTABLE![NAV_MOD_ROW] * NAVIGATION_LEVEL)
            VESSEL_MOVEMENT_SAIL = VESSEL_MOVEMENT_SAIL + (SHIPSTABLE![NAV_MOD_SAIL] * NAVIGATION_LEVEL)
            VESSEL_MOVEMENT_ROW = VESSEL_MOVEMENT_ROW + (SHIPSTABLE![SEA_MOD_ROW] * SEAMANSHIP_LEVEL)
            VESSEL_MOVEMENT_SAIL = VESSEL_MOVEMENT_SAIL + (SHIPSTABLE![SEA_MOD_SAIL] * SEAMANSHIP_LEVEL)
            VESSEL_MOVEMENT_ROW = VESSEL_MOVEMENT_ROW + (SHIPSTABLE![ROW_MOD] * ROWING_LEVEL)
            VESSEL_MOVEMENT_SAIL = VESSEL_MOVEMENT_SAIL + (SHIPSTABLE![SAIL_MOD] * SAILING_LEVEL)
            SHIP_TYPE = SHIPSTABLE![VESSEL]
            If SHIPSTABLE![ROWING_POSSIBLE] = "YES" Then
               If ROWING_MOVEMENT = 0 Then
                  ROWING_MOVEMENT = VESSEL_MOVEMENT_ROW
               ElseIf VESSEL_MOVEMENT_ROW < ROWING_MOVEMENT Then
                 ROWING_MOVEMENT = VESSEL_MOVEMENT_ROW
               End If
            End If
            If SHIPSTABLE![SAILING_POSSIBLE] = "YES" Then
               If SAILING_MOVEMENT = 0 Then
                  SAILING_MOVEMENT = VESSEL_MOVEMENT_SAIL
               ElseIf VESSEL_MOVEMENT_SAIL < SAILING_MOVEMENT Then
                 SAILING_MOVEMENT = VESSEL_MOVEMENT_SAIL
               End If
            End If
            If SAILING_POSSIBLE = "MAYBE" Then
               SAILING_POSSIBLE = SHIPSTABLE![SAILING_POSSIBLE]
            End If
            If ROWING_POSSIBLE = "MAYBE" Then
               ROWING_POSSIBLE = SHIPSTABLE![ROWING_POSSIBLE]
            End If
            If SHIPSTABLE![SAILING_POSSIBLE] = "NO" _
            And SAILING_POSSIBLE = "YES" Then
                ROWING_POSSIBLE = "BOTH"
                SAILING_POSSIBLE = "BOTH"
            End If
            If SHIPSTABLE![ROWING_POSSIBLE] = "NO" _
            And ROWING_POSSIBLE = "YES" Then
                ROWING_POSSIBLE = "BOTH"
                SAILING_POSSIBLE = "BOTH"
            End If
         End If
      End If
      SHIPSTABLE.MoveNext
   Loop
   
   If codetrack = 1 Then
      MSG1 = "ROWING_MOVEMENT = " & ROWING_MOVEMENT & crlf
      MSG2 = "SAILING_MOVEMENT = " & SAILING_MOVEMENT & crlf
      MSG3 = "SHIP_TYPE = " & SHIP_TYPE & crlf
      Response = MsgBox((MSG1 & MSG2 & MSG3), True)
   End If

   If SAILING_POSSIBLE = "YES" Then
      If ROWING_POSSIBLE = "YES" Then
         MOVEMENT_POINTS = SAILING_MOVEMENT
      Else
         MOVEMENT_POINTS = SAILING_MOVEMENT
      End If
   ElseIf ROWING_POSSIBLE = "YES" Then
      MOVEMENT_POINTS = ROWING_MOVEMENT
   ElseIf ROWING_POSSIBLE = "BOTH" Then
      MOVEMENT_POINTS = ROWING_MOVEMENT
   Else
      MOVEMENT_LINE = "Movement is not possible with this ship combination"
      MOVEMENT_POINTS = 0
   End If
   
Else
   ' IS THE GROUP MOUNTED
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "FINISHED", "WAGON"
   If TRIBESGOODS.NoMatch Then
      NUMBER_OF_WAGONS = 0
   Else
      NUMBER_OF_WAGONS = TRIBESGOODS![ITEM_NUMBER]
   End If
   
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "ANIMAL", "CAMEL"
   If TRIBESGOODS.NoMatch Then
      Number_Of_Camels = 0
   Else
      Number_Of_Camels = TRIBESGOODS![ITEM_NUMBER]
   End If
   
   Number_Of_Horses = 0
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "ANIMAL", "HORSE"
   If Not TRIBESGOODS.NoMatch Then
      Number_Of_Horses = TRIBESGOODS![ITEM_NUMBER]
   End If
   
   Number_Of_Cattle = 0
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "ANIMAL", "CATTLE"
   If Not TRIBESGOODS.NoMatch Then
      Number_Of_Cattle = TRIBESGOODS![ITEM_NUMBER]
   End If
   
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "ANIMAL", "HORSE/LIGHT"
   If Not TRIBESGOODS.NoMatch Then
      NUMBER_OF_LIGHTHORSE = TRIBESGOODS![ITEM_NUMBER]
      Number_Of_Horses = Number_Of_Horses + TRIBESGOODS![ITEM_NUMBER]
   End If
   
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "ANIMAL", "HORSE/HEAVY"
   If Not TRIBESGOODS.NoMatch Then
      NUMBER_OF_HEAVYHORSE = TRIBESGOODS![ITEM_NUMBER]
      Number_Of_Horses = Number_Of_Horses + TRIBESGOODS![ITEM_NUMBER]
   End If
   
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "ANIMAL", "ELEPHANT"
   If TRIBESGOODS.NoMatch Then
      Number_Of_Elephants = 0
   Else
      Number_Of_Elephants = TRIBESGOODS![ITEM_NUMBER]
   End If
   
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", MOVE_CLAN, GOODS_TRIBE, "WAR", "CHARIOTS - LIGHT"
   If TRIBESGOODS.NoMatch Then
      NUMBER_OF_CHARIOTS = 0
   Else
      NUMBER_OF_CHARIOTS = TRIBESGOODS![ITEM_NUMBER]
   End If

   HORSES = HORSES + HORSES_USED
   If HORSES > Number_Of_Horses Then
      HORSES_USED = 0
   End If
   
   CAMELS = CAMELS + CAMELS_USED
   If CAMELS > Number_Of_Camels Then
      CAMELS_USED = 0
   End If
   
   Elephants = Elephants + ELEPHANTS_USED
   If Elephants > Number_Of_Elephants Then
      ELEPHANTS_USED = 0
   End If
   
' Check if there are enough animals to pull wagons (AlexD 12-23-2024)
   If (SCOUTS = "N") And (NUMBER_OF_WAGONS * 2 > (Number_Of_Horses + Number_Of_Cattle) + Number_Of_Elephants * 2) Then
                MOVEMENT_POINTS = 0
                NO_MOVEMENT_REASON = "Not enough animals to pull wagons"
                Exit Sub
   End If
 
   If SCOUTS = "Y" Then
      Number_Of_People_Mounted = CAMELS_USED + HORSES_USED + (ELEPHANTS_USED * 3)
   Else
      Number_Of_People_Mounted = Number_Of_Camels + Number_Of_Horses + (Number_Of_Elephants * 3)
   End If



 
   If SCOUTS = "Y" Then
      If SCOUTS_USED <= Number_Of_People_Mounted Then
         If (NUMBER_OF_LIGHTHORSE >= HORSES_USED) And (NUMBER_OF_LIGHTHORSE > 0) Then
            MOVEMENT_POINTS = 20
         Else
            MOVEMENT_POINTS = 15
         End If
      Else
         MOVEMENT_POINTS = 8
      End If
   ElseIf NUMBER_OF_CHARIOTS = 0 Then
      If NUMBER_OF_WAGONS = 0 Then
'         If Number_Of_People_Mounted >= Total_People And TRIBES_CAPACITY >= TRIBES_WEIGHT Then
         If Number_Of_People_Mounted >= Total_People Then
            If (NUMBER_OF_LIGHTHORSE >= HORSES_USED) And (NUMBER_OF_LIGHTHORSE > 0) Then
               MOVEMENT_POINTS = 35
            Else
               MOVEMENT_POINTS = 27
            End If
         Else
            MOVEMENT_POINTS = 18
         End If
      ElseIf Number_Of_Elephants >= NUMBER_OF_WAGONS Then
'         If Number_Of_People_Mounted >= Total_People And TRIBES_CAPACITY >= TRIBES_WEIGHT Then
         If Number_Of_People_Mounted >= Total_People Then
            MOVEMENT_POINTS = 27
         Else
            MOVEMENT_POINTS = 18
         End If
      Else
         MOVEMENT_POINTS = 18
      End If
   ElseIf NUMBER_OF_CHARIOTS > 0 Then
      Number_Of_People_Mounted = Number_Of_People_Mounted + (NUMBER_OF_CHARIOTS * 3)
'      If Number_Of_People_Mounted >= Total_People And TRIBES_CAPACITY >= TRIBES_WEIGHT Then
      If Number_Of_People_Mounted >= Total_People Then
         MOVEMENT_POINTS = 24
      Else
         MOVEMENT_POINTS = 18
      End If
   End If

End If

' GET MODIFIERS FOR SCOUTING AND MOVEMENT
' Picks up from the Primary Tribes modifiers in this instance
Set MODTABLE = TVDBGM.OpenRecordset("MODIFIERS")
MODTABLE.index = "PRIMARYKEY"
MODTABLE.MoveFirst
If SCOUTS = "Y" Then
   MODTABLE.Seek "=", PRIMARY_TRIBE, "SCOUT MOVEMENT"
   If Not MODTABLE.NoMatch Then
      MOVEMENT_POINTS = MOVEMENT_POINTS + MODTABLE![AMOUNT]
   Else
      MODTABLE.Seek "=", MOVE_TRIBE, "SCOUT MOVEMENT"
      If Not MODTABLE.NoMatch Then
         MOVEMENT_POINTS = MOVEMENT_POINTS + MODTABLE![AMOUNT]
      Else
         ' no modifier found
      End If
   End If
Else
   MODTABLE.Seek "=", PRIMARY_TRIBE, "TRIBE MOVEMENT"
   If Not MODTABLE.NoMatch Then
      MOVEMENT_POINTS = MOVEMENT_POINTS + MODTABLE![AMOUNT]
   Else
      MODTABLE.Seek "=", MOVE_TRIBE, "TRIBE MOVEMENT"
      If Not MODTABLE.NoMatch Then
         MOVEMENT_POINTS = MOVEMENT_POINTS + MODTABLE![AMOUNT]
      Else
         ' no modifier found
      End If
   End If
End If

End Sub

Sub GET_ALPS(TERRAIN, CURRENT_MAP)
Dim ALPS_FOUND As String

ALPS_FOUND = "NO"

If NE_TERRAIN = "ALPS" Then
   TERRAIN = TERRAIN & ", Alps NE"
   ALPS_FOUND = "YES"
End If

If SE_TERRAIN = "ALPS" Then
   If ALPS_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & ",Alps SE"
      ALPS_FOUND = "YES"
   End If
End If

If SW_TERRAIN = "ALPS" Then
   If ALPS_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & ",Alps SW"
      ALPS_FOUND = "YES"
   End If
End If

If NW_TERRAIN = "ALPS" Then
   If ALPS_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & ",Alps NW"
      ALPS_FOUND = "YES"
   End If
End If

If N_TERRAIN = "ALPS" Then
   If ALPS_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      ALPS_FOUND = TERRAIN & ",Alps N"
      ALPS_FOUND = "YES"
   End If
End If

If S_TERRAIN = "ALPS" Then
   If ALPS_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & ",Alps S"
      ALPS_FOUND = "YES"
   End If
End If


End Sub

Function GET_BEACHS(TERRAIN, CURRENT_MAP)
Dim BEACH_FOUND As String

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

BEACH_FOUND = "NO"

If Mid(hexmaptable![Borders], 1, 2) = "BE" Then
   TERRAIN = TERRAIN & " Beachs N"
   BEACH_FOUND = "YES"
End If
If Mid(hexmaptable![Borders], 3, 2) = "BE" Then
   If BEACH_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & " Beachs NE"
      BEACH_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 5, 2) = "BE" Then
   If BEACH_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & " Beachs SE"
      BEACH_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 7, 2) = "BE" Then
   If BEACH_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & " Beachs S"
      BEACH_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 9, 2) = "BE" Then
   If BEACH_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & " Beachs SW"
      BEACH_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 11, 2) = "BE" Then
   If BEACH_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & " Beachs NW"
      BEACH_FOUND = "YES"
   End If
End If

If BEACH_FOUND = "YES" Then
   TERRAIN = TERRAIN & " NW" & ","
End If


End Function

Function GET_CANALS(TERRAIN, CURRENT_MAP)
Dim CANAL_FOUND As String

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

CANAL_FOUND = "NO"

If Mid(hexmaptable![Borders], 1, 2) = "CA" Then
   TERRAIN = TERRAIN & ",Canal N"
   CANAL_FOUND = "YES"
End If
If Mid(hexmaptable![Borders], 3, 2) = "CA" Then
   If CANAL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",Canal NE"
      CANAL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 5, 2) = "CA" Then
   If CANAL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",Canal SE"
      CANAL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 7, 2) = "CA" Then
   If CANAL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",Canal S"
      CANAL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 9, 2) = "CA" Then
   If CANAL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",Canal SW"
      CANAL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 11, 2) = "CA" Then
   If CANAL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",Canal NW"
      CANAL_FOUND = "YES"
   End If
End If

End Function

Function GET_CANYONS(TERRAIN, CURRENT_MAP)
Dim CANYONS_FOUND As String

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

CANYONS_FOUND = "NO"

If Mid(hexmaptable![CANYONS], 1, 1) = "Y" Then
   TERRAIN = TERRAIN & ",Canyons N"
   CANYONS_FOUND = "YES"
End If
If Mid(hexmaptable![CANYONS], 2, 1) = "Y" Then
   If CANYONS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",Canyons NE"
      CANYONS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![CANYONS], 3, 1) = "Y" Then
   If CANYONS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",Canyons SE"
      CANYONS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![CANYONS], 4, 1) = "Y" Then
   If CANYONS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",Canyons S"
      CANYONS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![CANYONS], 5, 1) = "Y" Then
   If CANYONS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",Canyons SW"
      CANYONS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![CANYONS], 6, 1) = "Y" Then
   If CANYONS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",Canyons NW"
      CANYONS_FOUND = "YES"
   End If
End If

End Function

Sub GET_CITIES(TERRAIN, CURRENT_MAP)

   HEXMAPCITY.MoveFirst
   HEXMAPCITY.Seek "=", CURRENT_MAP

   If Not HEXMAPCITY.NoMatch Then
      If Not IsNull(HEXMAPCITY![CITY]) Then
         TERRAIN = TERRAIN & ", " & HEXMAPCITY![CITY]
         Set Printing_Switch_TABLE = TVDB.OpenRecordset("Printing_Switchs")
         Printing_Switch_TABLE.index = "PRIMARYKEY"
         Printing_Switch_TABLE.Seek "=", MOVE_CLAN
         
         If Printing_Switch_TABLE.NoMatch Then
            Printing_Switch_TABLE.AddNew
            Printing_Switch_TABLE![CLAN] = MOVE_CLAN
            Printing_Switch_TABLE![CITY] = HEXMAPCITY![CITY]
            Printing_Switch_TABLE.UPDATE
            Printing_Switch_TABLE.Close
         Else
            Printing_Switch_TABLE.Edit
            Printing_Switch_TABLE![CITY] = HEXMAPCITY![CITY]
            Printing_Switch_TABLE.UPDATE
            Printing_Switch_TABLE.Close
         End If
         
         ' update a small table
         Msg = MOVE_TRIBE & " HAS REACHED THE CITY " & HEXMAPCITY![CITY]
         MsgBox (Msg)
      End If
   End If
   
End Sub

Function GET_CLIFFS(TERRAIN, CURRENT_MAP)
Dim CLIFFS_FOUND As String

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

CLIFFS_FOUND = "NO"

If Mid(hexmaptable![Borders], 1, 2) = "CL" Then
   TERRAIN = TERRAIN & ",Cliffs N"
   CLIFFS_FOUND = "YES"
End If
If Mid(hexmaptable![Borders], 3, 2) = "CL" Then
   If CLIFFS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",Cliffs NE"
      CLIFFS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 5, 2) = "CL" Then
   If CLIFFS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",Cliffs SE"
      CLIFFS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 7, 2) = "CL" Then
   If CLIFFS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",Cliffs S"
      CLIFFS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 9, 2) = "CL" Then
   If CLIFFS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",Cliffs SW"
      CLIFFS_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 11, 2) = "CL" Then
   If CLIFFS_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",Cliffs NW"
      CLIFFS_FOUND = "YES"
   End If
End If

End Function

Function GET_FISH_AREA(TERRAIN, CURRENT_MAP)
Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If hexmaptable![FISH AREA] = "Y" Then
   TERRAIN = TERRAIN & ", Fish Area (Improved Fishing)"
End If

End Function

Function GET_FORDS(TERRAIN, CURRENT_MAP)
Dim FORD_FOUND As String

FORD_FOUND = "NO"

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If Mid(hexmaptable![Borders], 1, 2) = "FO" Then
   TERRAIN = TERRAIN & ",Ford N"
   FORD_FOUND = "YES"
End If
If Mid(hexmaptable![Borders], 3, 2) = "FO" Then
   If FORD_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",Ford NE"
      FORD_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 5, 2) = "FO" Then
   If FORD_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",Ford SE"
      FORD_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 7, 2) = "FO" Then
   If FORD_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",Ford S"
      FORD_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 9, 2) = "FO" Then
   If FORD_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",Ford SW"
      FORD_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 11, 2) = "FO" Then
   If FORD_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",Ford NW"
      FORD_FOUND = "YES"
   End If
End If

End Function


Function GET_MAP_NORTH(HEX_MAP)
      
   ORIG_DOWN_MAP_LETTER = Determine_Hex_Map_Down_Letter(HEX_MAP)
   ORIG_ACROSS_MAP_LETTER = Determine_Hex_Map_Across_Letter(HEX_MAP)
   
   ORIG_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2)
   ORIG_DOWN_NUMBER = Mid(HEX_MAP, 6, 2)
   
   WORK_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2) - 1
   WORK_DOWN_NUMBER = Mid(HEX_MAP, 6, 2) - 1

   TRANSLATED_MAP_DOWN = ORIG_DOWN_MAP_LETTER + WORK_DOWN_NUMBER
   TRANSLATED_MAP_ACROSS = ORIG_ACROSS_MAP_LETTER + WORK_ACROSS_NUMBER
   
   NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER - 1
   TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN - 1
   
   NEW_DOWN_MAP_LETTER = Determine_New_Hex_Map_Down_Letter(TRANSLATED_MAP_DOWN)
   NEW_ACROSS_MAP_LETTER = Determine_New_Hex_Map_Across_Letter(TRANSLATED_MAP_ACROSS)
   
   MAPNUMBER = NEW_DOWN_MAP_LETTER & NEW_ACROSS_MAP_LETTER

   HEXNUMBER = MAPNUMBER & " "

   If ORIG_ACROSS_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & ORIG_ACROSS_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & ORIG_ACROSS_NUMBER
   End If

   If NEW_DOWN_NUMBER < 1 Then
      NEW_DOWN_NUMBER = 21
   End If
   
   If NEW_DOWN_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_DOWN_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_DOWN_NUMBER
   End If
      
   GET_MAP_NORTH = HEXNUMBER

End Function

Function GET_MAP_NORTH_EAST(HEX_MAP)
      
   ORIG_DOWN_MAP_LETTER = Determine_Hex_Map_Down_Letter(HEX_MAP)
   ORIG_ACROSS_MAP_LETTER = Determine_Hex_Map_Across_Letter(HEX_MAP)
   
   ORIG_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2)
   ORIG_DOWN_NUMBER = Mid(HEX_MAP, 6, 2)
   
   WORK_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2) - 1
   WORK_DOWN_NUMBER = Mid(HEX_MAP, 6, 2) - 1
   
   TRANSLATED_MAP_DOWN = ORIG_DOWN_MAP_LETTER + WORK_DOWN_NUMBER
   TRANSLATED_MAP_ACROSS = ORIG_ACROSS_MAP_LETTER + WORK_ACROSS_NUMBER
   
   If ORIG_ACROSS_NUMBER Mod 2 > 0 Then
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER - 1
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER + 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN - 1
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS + 1
   Else
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER - 0
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER + 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN - 0
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS + 1

   End If

   NEW_DOWN_MAP_LETTER = Determine_New_Hex_Map_Down_Letter(TRANSLATED_MAP_DOWN)
   NEW_ACROSS_MAP_LETTER = Determine_New_Hex_Map_Across_Letter(TRANSLATED_MAP_ACROSS)

   MAPNUMBER = NEW_DOWN_MAP_LETTER & NEW_ACROSS_MAP_LETTER
   
   HEXNUMBER = MAPNUMBER & " "

   If NEW_ACROSS_NUMBER > 30 Then
      NEW_ACROSS_NUMBER = 1
   End If
      
   If NEW_ACROSS_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_ACROSS_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_ACROSS_NUMBER
   End If

   If NEW_DOWN_NUMBER < 1 Then
      NEW_DOWN_NUMBER = 21
   End If
      
   If NEW_DOWN_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_DOWN_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_DOWN_NUMBER
   End If
      
   GET_MAP_NORTH_EAST = HEXNUMBER


End Function

Function GET_MAP_NORTH_WEST(HEX_MAP)
   
   ORIG_DOWN_MAP_LETTER = Determine_Hex_Map_Down_Letter(HEX_MAP)
   ORIG_ACROSS_MAP_LETTER = Determine_Hex_Map_Across_Letter(HEX_MAP)

   ORIG_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2)
   ORIG_DOWN_NUMBER = Mid(HEX_MAP, 6, 2)
   
   WORK_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2) - 1
   WORK_DOWN_NUMBER = Mid(HEX_MAP, 6, 2) - 1
   
   TRANSLATED_MAP_DOWN = ORIG_DOWN_MAP_LETTER + WORK_DOWN_NUMBER
   TRANSLATED_MAP_ACROSS = ORIG_ACROSS_MAP_LETTER + WORK_ACROSS_NUMBER

   If ORIG_ACROSS_NUMBER = 0 Then
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER - 0
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER - 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN - 0
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS - 1
   ElseIf ORIG_ACROSS_NUMBER Mod 2 > 0 Then
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER - 1
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER - 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN - 1
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS - 1
   Else
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER - 0
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER - 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN - 0
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS - 1
   End If

   NEW_DOWN_MAP_LETTER = Determine_New_Hex_Map_Down_Letter(TRANSLATED_MAP_DOWN)
   NEW_ACROSS_MAP_LETTER = Determine_New_Hex_Map_Across_Letter(TRANSLATED_MAP_ACROSS)
   
   MAPNUMBER = NEW_DOWN_MAP_LETTER & NEW_ACROSS_MAP_LETTER
   
   HEXNUMBER = MAPNUMBER & " "

   If NEW_ACROSS_NUMBER < 1 Then
      NEW_ACROSS_NUMBER = 30
   End If
      
   If NEW_ACROSS_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_ACROSS_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_ACROSS_NUMBER
   End If

   If NEW_DOWN_NUMBER < 1 Then
      NEW_DOWN_NUMBER = 21
   End If
      
   If NEW_DOWN_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_DOWN_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_DOWN_NUMBER
   End If
      
   GET_MAP_NORTH_WEST = HEXNUMBER

End Function

Function GET_MAP_SOUTH(HEX_MAP)
      
   ORIG_DOWN_MAP_LETTER = Determine_Hex_Map_Down_Letter(HEX_MAP)
   ORIG_ACROSS_MAP_LETTER = Determine_Hex_Map_Across_Letter(HEX_MAP)
   
   ORIG_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2)
   ORIG_DOWN_NUMBER = Mid(HEX_MAP, 6, 2)
   
   WORK_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2) - 1
   WORK_DOWN_NUMBER = Mid(HEX_MAP, 6, 2) - 1
   
   TRANSLATED_MAP_DOWN = ORIG_DOWN_MAP_LETTER + WORK_DOWN_NUMBER
   TRANSLATED_MAP_ACROSS = ORIG_ACROSS_MAP_LETTER + WORK_ACROSS_NUMBER

   NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER + 1
   
   TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN + 1
   
   NEW_DOWN_MAP_LETTER = Determine_New_Hex_Map_Down_Letter(TRANSLATED_MAP_DOWN)
   NEW_ACROSS_MAP_LETTER = Determine_New_Hex_Map_Across_Letter(TRANSLATED_MAP_ACROSS)
   
   MAPNUMBER = NEW_DOWN_MAP_LETTER & NEW_ACROSS_MAP_LETTER
   
   HEXNUMBER = MAPNUMBER & " "

   If ORIG_ACROSS_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & ORIG_ACROSS_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & ORIG_ACROSS_NUMBER
   End If

   If NEW_DOWN_NUMBER > 21 Then
      NEW_DOWN_NUMBER = 1
   End If
   
   If NEW_DOWN_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_DOWN_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_DOWN_NUMBER
   End If
      
   GET_MAP_SOUTH = HEXNUMBER

End Function

Function GET_MAP_SOUTH_EAST(HEX_MAP)
      
   ORIG_DOWN_MAP_LETTER = Determine_Hex_Map_Down_Letter(HEX_MAP)
   ORIG_ACROSS_MAP_LETTER = Determine_Hex_Map_Across_Letter(HEX_MAP)

   ORIG_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2)
   ORIG_DOWN_NUMBER = Mid(HEX_MAP, 6, 2)
   
   WORK_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2) - 1
   WORK_DOWN_NUMBER = Mid(HEX_MAP, 6, 2) - 1
   
   TRANSLATED_MAP_DOWN = ORIG_DOWN_MAP_LETTER + WORK_DOWN_NUMBER
   TRANSLATED_MAP_ACROSS = ORIG_ACROSS_MAP_LETTER + WORK_ACROSS_NUMBER
   
   If ORIG_ACROSS_NUMBER Mod 2 > 0 Then
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER + 0
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER + 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN + 0
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS + 1
   Else
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER + 1
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER + 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN + 1
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS + 1

   End If
   
   NEW_DOWN_MAP_LETTER = Determine_New_Hex_Map_Down_Letter(TRANSLATED_MAP_DOWN)
   NEW_ACROSS_MAP_LETTER = Determine_New_Hex_Map_Across_Letter(TRANSLATED_MAP_ACROSS)
   
   MAPNUMBER = NEW_DOWN_MAP_LETTER & NEW_ACROSS_MAP_LETTER
   
   HEXNUMBER = MAPNUMBER & " "

   If NEW_ACROSS_NUMBER > 30 Then
      NEW_ACROSS_NUMBER = 1
   End If
      
   If NEW_ACROSS_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_ACROSS_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_ACROSS_NUMBER
   End If

   If NEW_DOWN_NUMBER > 21 Then
      NEW_DOWN_NUMBER = 1
   End If
      
   If NEW_DOWN_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_DOWN_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_DOWN_NUMBER
   End If

   GET_MAP_SOUTH_EAST = HEXNUMBER


End Function

Function GET_MAP_SOUTH_WEST(HEX_MAP)
      
   ORIG_DOWN_MAP_LETTER = Determine_Hex_Map_Down_Letter(HEX_MAP)
   ORIG_ACROSS_MAP_LETTER = Determine_Hex_Map_Across_Letter(HEX_MAP)
   
   ORIG_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2)
   ORIG_DOWN_NUMBER = Mid(HEX_MAP, 6, 2)
   
   WORK_ACROSS_NUMBER = Mid(HEX_MAP, 4, 2) - 1
   WORK_DOWN_NUMBER = Mid(HEX_MAP, 6, 2) - 1
   
   TRANSLATED_MAP_DOWN = ORIG_DOWN_MAP_LETTER + WORK_DOWN_NUMBER
   TRANSLATED_MAP_ACROSS = ORIG_ACROSS_MAP_LETTER + WORK_ACROSS_NUMBER

   If ORIG_ACROSS_NUMBER Mod 2 > 0 Then
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER + 0
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER - 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN + 0
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS - 1
   Else
      NEW_DOWN_NUMBER = ORIG_DOWN_NUMBER + 1
      NEW_ACROSS_NUMBER = ORIG_ACROSS_NUMBER - 1
      TRANSLATED_MAP_DOWN = TRANSLATED_MAP_DOWN + 1
      TRANSLATED_MAP_ACROSS = TRANSLATED_MAP_ACROSS - 1

   End If
   
   NEW_DOWN_MAP_LETTER = Determine_New_Hex_Map_Down_Letter(TRANSLATED_MAP_DOWN)
   NEW_ACROSS_MAP_LETTER = Determine_New_Hex_Map_Across_Letter(TRANSLATED_MAP_ACROSS)
   
   MAPNUMBER = NEW_DOWN_MAP_LETTER & NEW_ACROSS_MAP_LETTER
   
   HEXNUMBER = MAPNUMBER & " "

   If NEW_ACROSS_NUMBER < 1 Then
      NEW_ACROSS_NUMBER = 30
   End If
      
   If NEW_ACROSS_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_ACROSS_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_ACROSS_NUMBER
   End If

   If NEW_DOWN_NUMBER > 21 Then
      NEW_DOWN_NUMBER = 1
   End If
      
   If NEW_DOWN_NUMBER < 10 Then
      HEXNUMBER = HEXNUMBER & "0" & NEW_DOWN_NUMBER
   Else
      HEXNUMBER = HEXNUMBER & NEW_DOWN_NUMBER
   End If
      
   GET_MAP_SOUTH_WEST = HEXNUMBER

End Function

Function GET_MOUNTAINS(TERRAIN, CURRENT_MAP)
Dim Lcm_FOUND As String
Dim Lsm_FOUND As String
Dim Hsm_FOUND As String
Dim LVm_FOUND As String
Dim LJm_FOUND As String

Lcm_FOUND = "NO"
Lsm_FOUND = "NO"
Hsm_FOUND = "NO"
LVm_FOUND = "NO"
LJm_FOUND = "NO"

If (NE_TERRAIN = "LOW CONIFER MOUNTAINS") Or (NE_TERRAIN = "LOW CONIFER MT") Then
   TERRAIN = TERRAIN & " Lcm NE"
   Lcm_FOUND = "YES"
End If

If (SE_TERRAIN = "LOW CONIFER MOUNTAINS") Or (SE_TERRAIN = "LOW CONIFER MT") Then
   If Lcm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " Lcm SE"
      Lcm_FOUND = "YES"
   End If
End If

If (SW_TERRAIN = "LOW CONIFER MOUNTAINS") Or (SW_TERRAIN = "LOW CONIFER MT") Then
   If Lcm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " Lcm SW"
      Lcm_FOUND = "YES"
   End If
End If

If (NW_TERRAIN = "LOW CONIFER MOUNTAINS") Or (NW_TERRAIN = "LOW CONIFER MT") Then
   If Lcm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " Lcm NW"
      Lcm_FOUND = "YES"
   End If
End If

If (N_TERRAIN = "LOW CONIFER MOUNTAINS") Or (N_TERRAIN = "LOW CONIFER MT") Then
   If Lcm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " Lcm N"
      Lcm_FOUND = "YES"
   End If
End If

If (S_TERRAIN = "LOW CONIFER MOUNTAINS") Or (S_TERRAIN = "LOW CONIFER MT") Then
   If Lcm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " Lcm S"
      Lcm_FOUND = "YES"
   End If
End If

If Lcm_FOUND = "YES" Then
   TERRAIN = TERRAIN & ","
End If

If (NE_TERRAIN = "LOW SNOWY MOUNTAINS") Or (NE_TERRAIN = "LOW SNOWY MT") Then
   TERRAIN = TERRAIN & " Lsm NE"
   Lsm_FOUND = "YES"
End If

If (SE_TERRAIN = "LOW SNOWY MOUNTAINS") Or (SE_TERRAIN = "LOW SNOWY MT") Then
   If Lsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " Lsm SE"
      Lsm_FOUND = "YES"
   End If
End If

If (SW_TERRAIN = "LOW SNOWY MOUNTAINS") Or (SW_TERRAIN = "LOW SNOWY MT") Then
   If Lsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " Lsm SW"
      Lsm_FOUND = "YES"
   End If
End If

If (NW_TERRAIN = "LOW SNOWY MOUNTAINS") Or (NW_TERRAIN = "LOW SNOWY MT") Then
   If Lsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " Lsm NW"
      Lsm_FOUND = "YES"
   End If
End If

If (N_TERRAIN = "LOW SNOWY MOUNTAINS") Or (N_TERRAIN = "LOW SNOWY MT") Then
   If Lsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " Lsm N"
      Lsm_FOUND = "YES"
   End If
End If

If (S_TERRAIN = "LOW SNOWY MOUNTAINS") Or (S_TERRAIN = "LOW SNOWY MT") Then
   If Lsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " Lsm S"
      Lsm_FOUND = "YES"
   End If
End If

If Lsm_FOUND = "YES" Then
   TERRAIN = TERRAIN & ","
End If

If (NE_TERRAIN = "HIGH SNOWY MOUNTAINS") Or (NE_TERRAIN = "HIGH SNOWY MT") Then
   TERRAIN = TERRAIN & " Hsm NE"
   Hsm_FOUND = "YES"
End If

If (SE_TERRAIN = "HIGH SNOWY MOUNTAINS") Or (SE_TERRAIN = "HIGH SNOWY MT") Then
   If Hsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " Hsm SE"
      Hsm_FOUND = "YES"
   End If
End If

If (SW_TERRAIN = "HIGH SNOWY MOUNTAINS") Or (SW_TERRAIN = "HIGH SNOWY MT") Then
   If Hsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " Hsm SW"
      Hsm_FOUND = "YES"
   End If
End If

If (NW_TERRAIN = "HIGH SNOWY MOUNTAINS") Or (NW_TERRAIN = "HIGH SNOWY MT") Then
   If Hsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " Hsm NW"
      Hsm_FOUND = "YES"
   End If
End If

If (N_TERRAIN = "HIGH SNOWY MOUNTAINS") Or (N_TERRAIN = "HIGH SNOWY MT") Then
   If Hsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " Hsm N"
      Hsm_FOUND = "YES"
   End If
End If

If (S_TERRAIN = "HIGH SNOWY MOUNTAINS") Or (S_TERRAIN = "HIGH SNOWY MT") Then
   If Hsm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " Hsm S"
      Hsm_FOUND = "YES"
   End If
End If

If Hsm_FOUND = "YES" Then
   TERRAIN = TERRAIN & ","
End If

If (NE_TERRAIN = "LOW VOLCANO MOUNTAINS") Or (NE_TERRAIN = "LOW VOLCANO MT") Then
   TERRAIN = TERRAIN & " LVm NE"
   LVm_FOUND = "YES"
End If

If (SE_TERRAIN = "LOW VOLCANO MOUNTAINS") Or (SE_TERRAIN = "LOW VOLCANO MT") Then
   If LVm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " LVm SE"
      LVm_FOUND = "YES"
   End If
End If

If (SW_TERRAIN = "LOW VOLCANO MOUNTAINS") Or (SW_TERRAIN = "LOW VOLCANO MT") Then
   If LVm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " LVm SW"
      LVm_FOUND = "YES"
   End If
End If

If (NW_TERRAIN = "LOW VOLCANO MOUNTAINS") Or (NW_TERRAIN = "LOW VOLCANO MT") Then
   If LVm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " LVm NW"
      LVm_FOUND = "YES"
   End If
End If

If (N_TERRAIN = "LOW VOLCANO MOUNTAINS") Or (N_TERRAIN = "LOW VOLCANO MT") Then
   If LVm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " LVm N"
      LVm_FOUND = "YES"
   End If
End If

If (S_TERRAIN = "LOW VOLCANO MOUNTAINS") Or (S_TERRAIN = "LOW VOLCANO MT") Then
   If LVm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " LVm S"
      LVm_FOUND = "YES"
   End If
End If

If LVm_FOUND = "YES" Then
   TERRAIN = TERRAIN & ","
End If

If (NE_TERRAIN = "LOW JUNGLE MOUNTAINS") Or (NE_TERRAIN = "LOW JUNGLE MT") Then
   TERRAIN = TERRAIN & " LJm NE"
   LJm_FOUND = "YES"
End If

If (SE_TERRAIN = "LOW JUNGLE MOUNTAINS") Or (SE_TERRAIN = "LOW JUNGLE MT") Then
   If LJm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " LJm SE"
      LJm_FOUND = "YES"
   End If
End If

If (SW_TERRAIN = "LOW JUNGLE MOUNTAINS") Or (SW_TERRAIN = "LOW JUNGLE MT") Then
   If LJm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " LJm SW"
      LJm_FOUND = "YES"
   End If
End If

If (NW_TERRAIN = "LOW JUNGLE MOUNTAINS") Or (NW_TERRAIN = "LOW JUNGLE MT") Then
   If LJm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " LJm NW"
      LJm_FOUND = "YES"
   End If
End If

If (N_TERRAIN = "LOW JUNGLE MOUNTAINS") Or (N_TERRAIN = "LOW JUNGLE MT") Then
   If LJm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " LJm N"
      LJm_FOUND = "YES"
   End If
End If

If (S_TERRAIN = "LOW JUNGLE MOUNTAINS") Or (S_TERRAIN = "LOW JUNGLE MT") Then
   If LJm_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " LJm S"
      LJm_FOUND = "YES"
   End If
End If
   
If LJm_FOUND = "YES" Then
   TERRAIN = TERRAIN & ","
End If

End Function

Sub GET_NEXT_SCOUT_MOVE()

RESTART:

   stext = "MOVEMENT" & CStr(cnt1)
   If IsNull(SCOUT_MOVEMENT_TABLE(stext).Value) Then
      Direction = "EMPTY"
      ORIG_Direction = "EMPTY"
   Else
      Direction = SCOUT_MOVEMENT_TABLE(stext).Value
      ORIG_Direction = SCOUT_MOVEMENT_TABLE(stext).Value
   End If

End Sub

Function GET_PASSES(TERRAIN, CURRENT_MAP)
Dim PASS_FOUND As String

PASS_FOUND = "NO"

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

         If Mid(hexmaptable![Borders], 1, 2) = "PA" Then
            TERRAIN = TERRAIN & ",Pass N"
            PASS_FOUND = "YES"
         End If
         If Mid(hexmaptable![Borders], 3, 2) = "PA" Then
            If PASS_FOUND = "YES" Then
               TERRAIN = TERRAIN & " NE"
            Else
               TERRAIN = TERRAIN & ",Pass NE"
               PASS_FOUND = "YES"
            End If
         End If
         If Mid(hexmaptable![Borders], 5, 2) = "PA" Then
            If PASS_FOUND = "YES" Then
               TERRAIN = TERRAIN & " SE"
            Else
               TERRAIN = TERRAIN & ",Pass SE"
               PASS_FOUND = "YES"
            End If
         End If
         If Mid(hexmaptable![Borders], 7, 2) = "PA" Then
            If PASS_FOUND = "YES" Then
               TERRAIN = TERRAIN & " S"
            Else
               TERRAIN = TERRAIN & ",Pass S"
               PASS_FOUND = "YES"
            End If
         End If
         If Mid(hexmaptable![Borders], 9, 2) = "PA" Then
            If PASS_FOUND = "YES" Then
               TERRAIN = TERRAIN & " SW"
            Else
               TERRAIN = TERRAIN & ",Pass SW"
               PASS_FOUND = "YES"
            End If
         End If
         If Mid(hexmaptable![Borders], 11, 2) = "PA" Then
            If PASS_FOUND = "YES" Then
               TERRAIN = TERRAIN & " NW"
            Else
               TERRAIN = TERRAIN & ",Pass NW"
               PASS_FOUND = "YES"
            End If
         End If

End Function

Function GET_RIVERS(TERRAIN, CURRENT_MAP)
Dim RIVER_FOUND As String

RIVER_FOUND = "NO"

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If Mid(hexmaptable![Borders], 1, 2) = "RI" Then
   TERRAIN = TERRAIN & ",River N"
   RIVER_FOUND = "YES"
End If
If Mid(hexmaptable![Borders], 3, 2) = "RI" Then
   If RIVER_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",River NE"
      RIVER_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 5, 2) = "RI" Then
   If RIVER_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",River SE"
      RIVER_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 7, 2) = "RI" Then
   If RIVER_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",River S"
      RIVER_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 9, 2) = "RI" Then
   If RIVER_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",River SW"
      RIVER_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![Borders], 11, 2) = "RI" Then
   If RIVER_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",River NW"
      RIVER_FOUND = "YES"
   End If
End If

End Function

Function GET_ROADS(TERRAIN, CURRENT_MAP)
   
Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If Mid(hexmaptable![ROADS], 1, 1) = "R" Then
   TERRAIN = TERRAIN & " Rune Road N"
ElseIf Mid(hexmaptable![ROADS], 1, 1) = "D" Then
   TERRAIN = TERRAIN & " Dirt Road N"
ElseIf Mid(hexmaptable![ROADS], 1, 1) = "S" Then
   TERRAIN = TERRAIN & " Stone Road N"
End If

If Mid(hexmaptable![ROADS], 2, 1) = "R" Then
   If InStr(TERRAIN, "RUNE ROAD") Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & " Rune Road NE"
   End If
ElseIf Mid(hexmaptable![ROADS], 2, 1) = "D" Then
   If InStr(TERRAIN, "DIRT ROAD") Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & " Dirt Road NE"
   End If
ElseIf Mid(hexmaptable![ROADS], 2, 1) = "S" Then
   If InStr(TERRAIN, "STONE ROAD") Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & " Stone Road NE"
   End If
End If

If Mid(hexmaptable![ROADS], 3, 1) = "R" Then
   If InStr(TERRAIN, "RUNE ROAD") Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & " Rune Road SE"
   End If
ElseIf Mid(hexmaptable![ROADS], 3, 1) = "D" Then
   If InStr(TERRAIN, "DIRT ROAD") Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & " Dirt Road SE"
   End If
ElseIf Mid(hexmaptable![ROADS], 3, 1) = "S" Then
   If InStr(TERRAIN, "STONE ROAD") Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & " Stone Road SE"
   End If
End If

If Mid(hexmaptable![ROADS], 4, 1) = "R" Then
   If InStr(TERRAIN, "RUNE ROAD") Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & " Rune Road S"
   End If
ElseIf Mid(hexmaptable![ROADS], 4, 1) = "D" Then
   If InStr(TERRAIN, "DIRT ROAD") Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & " Dirt Road S"
   End If
ElseIf Mid(hexmaptable![ROADS], 4, 1) = "S" Then
   If InStr(TERRAIN, "STONE ROAD") Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & " Stone Road S"
   End If
End If

If Mid(hexmaptable![ROADS], 5, 1) = "R" Then
   If InStr(TERRAIN, "RUNE ROAD") Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & " Rune Road SW"
   End If
ElseIf Mid(hexmaptable![ROADS], 5, 1) = "D" Then
   If InStr(TERRAIN, "DIRT ROAD") Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & " Dirt Road SW"
   End If
ElseIf Mid(hexmaptable![ROADS], 5, 1) = "S" Then
   If InStr(TERRAIN, "STONE ROAD") Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & " Stone Road SW"
   End If
End If

If Mid(hexmaptable![ROADS], 6, 1) = "R" Then
   If InStr(TERRAIN, "RUNE ROAD") Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & " Rune Road NW"
   End If
ElseIf Mid(hexmaptable![ROADS], 6, 1) = "D" Then
   If InStr(TERRAIN, "DIRT ROAD") Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & " Dirt Road NW"
   End If
ElseIf Mid(hexmaptable![ROADS], 6, 1) = "S" Then
   If InStr(TERRAIN, "STONE ROAD") Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & " Stone Road NW"
   End If
End If

If InStr(TERRAIN, "STONE ROAD") Or InStr(TERRAIN, "DIRT ROAD") Or InStr(TERRAIN, "RUNE ROAD") Then
   TERRAIN = TERRAIN & ","
End If

End Function

Function GET_SALMON_RUN(TERRAIN, CURRENT_MAP)

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If hexmaptable![SALMON RUN] = "Y" Then
   TERRAIN = TERRAIN & ", Salmon Run (Improved Fishing)"
End If

End Function

Function GET_STREAMS(TERRAIN, CURRENT_MAP)
Dim STREAM_FOUND As String

STREAM_FOUND = "NO"

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If Mid(hexmaptable![STREAMS], 1, 1) = "Y" Then
   TERRAIN = TERRAIN & ",Streams N"
   STREAM_FOUND = "YES"
End If
If Mid(hexmaptable![STREAMS], 2, 1) = "Y" Then
   If STREAM_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",Streams NE"
      STREAM_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![STREAMS], 3, 1) = "Y" Then
   If STREAM_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",Streams SE"
      STREAM_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![STREAMS], 4, 1) = "Y" Then
   If STREAM_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",Streams S"
      STREAM_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![STREAMS], 5, 1) = "Y" Then
   If STREAM_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",Streams SW"
      STREAM_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![STREAMS], 6, 1) = "Y" Then
   If STREAM_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",Streams NW"
      STREAM_FOUND = "YES"
   End If
End If

End Function

Function GET_SURROUNDING_DATA(STARTING_HEX)
' THIS SUBFUNCTION TAKES LESS THAN 2 SECONDS
Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", STARTING_HEX

hexmaptable.Edit
N_HEX = GET_MAP_NORTH(STARTING_HEX)
NE_HEX = GET_MAP_NORTH_EAST(STARTING_HEX)
SE_HEX = GET_MAP_SOUTH_EAST(STARTING_HEX)
S_HEX = GET_MAP_SOUTH(STARTING_HEX)
SW_HEX = GET_MAP_SOUTH_WEST(STARTING_HEX)
NW_HEX = GET_MAP_NORTH_WEST(STARTING_HEX)

hexmaptable.MoveFirst
hexmaptable.Seek "=", NE_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NE_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NE_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NE_HEX
   End If
   NE_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NE_HEX
   End If
   NE_TERRAIN = hexmaptable![TERRAIN]
End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", SE_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(SE_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", SE_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SE_HEX
   End If
   SE_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SE_HEX
   End If
   SE_TERRAIN = hexmaptable![TERRAIN]
End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", N_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(N_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", N_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(N_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", N_HEX
   End If
   N_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(N_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", N_HEX
   End If
   N_TERRAIN = hexmaptable![TERRAIN]
End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", S_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(S_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", S_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(S_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", S_HEX
   End If
   S_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(S_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", S_HEX
   End If
   S_TERRAIN = hexmaptable![TERRAIN]
End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", SW_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(SW_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", SW_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SW_HEX
   End If
   SW_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SW_HEX
   End If
   SW_TERRAIN = hexmaptable![TERRAIN]
End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", NW_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NW_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NW_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NW_HEX
   End If
   NW_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NW_HEX
   End If
   NW_TERRAIN = hexmaptable![TERRAIN]
End If
   
End Function

Sub GET_SURROUNDING_FLEET(STARTING_HEX)
If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = GET_SURROUNDING_FLEET" & crlf
   Response = MsgBox((MSG1), True)
End If

START_TIME = Time

Dim MININGMOD As Long

TERRAIN_SURROUNDING_FLEET = "+9"

hexmaptable.MoveFirst
hexmaptable.Seek "=", STARTING_HEX

hexmaptable.Edit
FIRST_N_HEX = GET_MAP_NORTH(STARTING_HEX)
FIRST_NE_HEX = GET_MAP_NORTH_EAST(STARTING_HEX)
FIRST_SE_HEX = GET_MAP_SOUTH_EAST(STARTING_HEX)
FIRST_S_HEX = GET_MAP_SOUTH(STARTING_HEX)
FIRST_SW_HEX = GET_MAP_SOUTH_WEST(STARTING_HEX)
FIRST_NW_HEX = GET_MAP_NORTH_WEST(STARTING_HEX)

If codetrack = 1 Then
   MSG1 = "FIRST_NE_HEX = " & FIRST_NE_HEX & crlf
   MSG2 = "FIRST_SE_HEX = " & FIRST_SE_HEX & crlf
   MSG3 = "FIRST_N_HEX = " & FIRST_N_HEX & crlf
   MSG4 = "FIRST_S_HEX = " & FIRST_S_HEX & crlf
   MSG5 = "FIRST_SW_HEX = " & FIRST_SW_HEX & crlf
   MSG6 = "FIRST_NW_HEX = " & FIRST_NW_HEX & crlf
   Response = MsgBox((MSG1 & MSG2 & MSG3 & MSG4 & MSG5 & MSG6), True)
End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", FIRST_NE_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(FIRST_NE_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", FIRST_NE_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_NE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_NE_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_NE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_NE_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "NE " & SHORTTERRAIN & ", "

hexmaptable.MoveFirst
hexmaptable.Seek "=", FIRST_SE_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(FIRST_SE_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", FIRST_SE_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_SE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_SE_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_SE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_SE_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "SE " & SHORTTERRAIN & ", "

hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", FIRST_N_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(FIRST_N_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", FIRST_N_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_N_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_N_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_N_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_N_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "N " & SHORTTERRAIN & ", "

hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", FIRST_S_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(FIRST_S_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", FIRST_S_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_S_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_S_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_S_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_S_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "S " & SHORTTERRAIN & ", "

hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", FIRST_SW_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(FIRST_SW_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", FIRST_SW_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_SW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_SW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_SW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_SW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "SW " & SHORTTERRAIN & ", "

hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", FIRST_NW_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(FIRST_NW_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", FIRST_NW_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_NW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_NW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(FIRST_NW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", FIRST_NW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "NW " & SHORTTERRAIN & ", "

'GET SURROUNDING TERRAIN ONLY NOW.
TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "+0+9"

'GET TERRAIN FOR HEX NORTH/NORTH
'IF SPYGLASSES DISPLAY TERRAIN ELSE JUST HAVE Land AND Direction
NEW_HEX_NN = GET_MAP_NORTH(FIRST_N_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_NN

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_NN, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_NN
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NN)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NN
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NN)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NN
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "N/N " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - N/N,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - N/N,"
End If

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, FIRST_NW_HEX, "N")

'If Not TRIBESINHEX = "EMPTY" Then
'   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & " " & TRIBESINHEX & ", "
'End If
NEW_HEX_NNE = GET_MAP_NORTH_EAST(FIRST_N_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_NNE

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_NNE, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_NNE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NNE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NNE
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NNE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NNE
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "N/NE " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - N/NE,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - N/NE,"
End If

NEW_HEX_NNW = GET_MAP_NORTH_WEST(FIRST_N_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_NNW

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_NNW, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_NNW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NNW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NNW
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NNW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NNW
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "N/NW " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - N/NW,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - N/NW,"
End If

NEW_HEX_NENE = GET_MAP_NORTH_EAST(FIRST_NE_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_NENE

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_NENE, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_NENE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NENE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NENE
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NENE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NENE
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "NE/NE " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - NE/NE,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - NE/NE,"
End If

NEW_HEX_NESE = GET_MAP_SOUTH_EAST(FIRST_NE_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_NESE

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_NESE, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_NESE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NESE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NESE
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NESE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NESE
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "NE/SE " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - NE/SE,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - NE/SE,"
End If

NEW_HEX_SESE = GET_MAP_SOUTH_EAST(FIRST_SE_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_SESE

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_SESE, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_SESE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SESE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SESE
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SESE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SESE
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "SE/SE " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - SE/SE,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - SE/SE,"
End If

NEW_HEX_SSE = GET_MAP_SOUTH_EAST(FIRST_S_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_SSE

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_SSE, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_SSE
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SSE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SSE
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SSE)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SSE
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "S/SE " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - S/SE,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - S/SE,"
End If

NEW_HEX_SS = GET_MAP_SOUTH(FIRST_S_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_SS

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_SS, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_SS
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SS)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SS
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SS)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SS
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "S/S " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - S/S,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - S/S,"
End If

NEW_HEX_SSW = GET_MAP_SOUTH_WEST(FIRST_S_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_SSW

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_SSW, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_SSW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SSW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SSW
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SSW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SSW
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "S/SW " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - S/SW,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - S/SW,"
End If

NEW_HEX_SWSW = GET_MAP_SOUTH_WEST(FIRST_SW_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_SWSW

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_SWSW, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_SWSW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SWSW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SWSW
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SWSW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SWSW
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "SW/SW " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - SW/SW,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - SW/SW,"
End If

NEW_HEX_SWNW = GET_MAP_NORTH_WEST(FIRST_SW_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_SWNW

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_SWNW, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_SWNW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SWNW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SWNW
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_SWNW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_SWNW
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "SW/NW " & SHORTTERRAIN & ", "
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - SW/NW,"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - SW/NW,"
End If

NEW_HEX_NWNW = GET_MAP_NORTH_WEST(FIRST_NW_HEX)
hexmaptable.index = "PRIMARYKEY"
hexmaptable.Seek "=", NEW_HEX_NWNW

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NEW_HEX_NWNW, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NEW_HEX_NWNW
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NWNW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NWNW
   End If
   TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NEW_HEX_NWNW)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NEW_HEX_NWNW
   End If
   TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

If SPYGLASSES = "Y" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "NW/NW " & SHORTTERRAIN & "+0"
ElseIf Not SHORTTERRAIN = "O" And Not SHORTTERRAIN = "L" Then
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Land - NW/NW, +0"
Else
   TERRAIN_SURROUNDING_FLEET = TERRAIN_SURROUNDING_FLEET & "Sight Water - NW/NW, +0"
End If

TERRAIN = TERRAIN_SURROUNDING_FLEET

END_TIME = Time

If codetrack = 1 Then
   MSG0 = "SURROUND FLEET " & crlf
   MSG1 = "START_TIME = " & START_TIME & crlf
   MSG2 = "END_TIME = " & END_TIME & crlf
   MSG3 = "TERRAIN = " & TERRAIN & crlf
   MSG4 = "SURROUNDING_TERRAIN = " & SURROUNDING_TERRAIN & crlf
   Response = MsgBox((MSG0 & MSG1 & MSG2 & MSG3 & MSG4), True)
End If

End Sub

Sub GET_SURROUNDING_TERRAIN(STARTING_HEX)
If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = GET_SURROUNDING_TERRAIN" & crlf
   Response = MsgBox((MSG1), True)
End If

START_TIME = Time

Dim MININGMOD As Long

SURROUNDING_TERRAIN = "EMPTY"

' TRIBE MOVEMENT
hexmaptable.MoveFirst
hexmaptable.Seek "=", STARTING_HEX

hexmaptable.Edit
N_HEX = GET_MAP_NORTH(STARTING_HEX)
NE_HEX = GET_MAP_NORTH_EAST(STARTING_HEX)
SE_HEX = GET_MAP_SOUTH_EAST(STARTING_HEX)
S_HEX = GET_MAP_SOUTH(STARTING_HEX)
SW_HEX = GET_MAP_SOUTH_WEST(STARTING_HEX)
NW_HEX = GET_MAP_NORTH_WEST(STARTING_HEX)

SURROUNDING_TERRAIN = "+9"

count = 1

   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NE_HEX
   
   If hexmaptable.NoMatch Then
      Call ADD_NEW_HEX(NE_HEX, TERRAIN)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NE_HEX
      If IsNull(hexmaptable![TERRAIN]) Then
         Call UPDATE_HEX_MAP(NE_HEX)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", NE_HEX
      End If
      TERRAIN = hexmaptable![TERRAIN]
      NE_TERRAIN = hexmaptable![TERRAIN]
   Else
      If IsNull(hexmaptable![TERRAIN]) Then
        Call UPDATE_HEX_MAP(NE_HEX)
        hexmaptable.MoveFirst
        hexmaptable.Seek "=", NE_HEX
     End If
      TERRAIN = hexmaptable![TERRAIN]
     NE_TERRAIN = hexmaptable![TERRAIN]
  End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "NE " & SHORTTERRAIN & ", "

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, NE_HEX, "N")

'If Not TRIBESINHEX = "EMPTY" Then
'   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & " " & TRIBESINHEX & ", "
'End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", SE_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(SE_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", SE_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SE_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   SE_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SE_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SE_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   SE_TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "SE " & SHORTTERRAIN & ", "

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, SE_HEX, "N")

'If Not TRIBESINHEX = "EMPTY" Then
'   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & " " & TRIBESINHEX & ", "
'End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", N_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(N_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", N_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(N_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", N_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   N_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(N_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", N_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   N_TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "N " & SHORTTERRAIN & ", "

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, N_HEX, "N")

'If Not TRIBESINHEX = "EMPTY" Then
'   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & " " & TRIBESINHEX & ", "
'End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", S_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(S_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", S_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(S_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", S_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   S_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(S_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", S_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   S_TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "S " & SHORTTERRAIN & ", "

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, S_HEX, "N")

'If Not TRIBESINHEX = "EMPTY" Then
'   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & " " & TRIBESINHEX & ", "
'End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", SW_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(SW_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", SW_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   SW_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(SW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", SW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   SW_TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "SW " & SHORTTERRAIN & ", "

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, SW_HEX, "N")

'If Not TRIBESINHEX = "EMPTY" Then
'   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & " " & TRIBESINHEX & ", "
'End If

hexmaptable.MoveFirst
hexmaptable.Seek "=", NW_HEX

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(NW_HEX, TERRAIN)
   hexmaptable.MoveFirst
   hexmaptable.Seek "=", NW_HEX
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   NW_TERRAIN = hexmaptable![TERRAIN]
Else
   If IsNull(hexmaptable![TERRAIN]) Then
      Call UPDATE_HEX_MAP(NW_HEX)
      hexmaptable.MoveFirst
      hexmaptable.Seek "=", NW_HEX
   End If
   TERRAIN = hexmaptable![TERRAIN]
   NW_TERRAIN = hexmaptable![TERRAIN]
End If

Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "NW " & SHORTTERRAIN

'TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, NW_HEX, "N")

'If TRIBESINHEX = "EMPTY" Then
   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & "+0"
'Else
'   SURROUNDING_TERRAIN = SURROUNDING_TERRAIN & " " & TRIBESINHEX & "+0"
'End If

TERRAIN = SURROUNDING_TERRAIN

END_TIME = Time

If codetrack = 1 Then
   MSG0 = "SURROUNDING TERRAIN" & crlf
   MSG1 = "START_TIME = " & START_TIME & crlf
   MSG2 = "END_TIME = " & END_TIME & crlf
   Response = MsgBox((MSG0 & MSG1 & MSG2), True)
End If

End Sub

Sub GET_TERRAIN(Direction, TERRAIN, CURRENT_MAP)
If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = GET_TERRAIN" & crlf
   Response = MsgBox((MSG1), True)
End If

START_TIME = Time

Dim MININGMOD As Long

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(CURRENT_MAP, TERRAIN)
   hexmaptable.Seek "=", CURRENT_MAP
ElseIf IsNull(hexmaptable![TERRAIN]) Then
   Msg = "WHAT IS THE TERRAIN OF HEX " & CURRENT_MAP & "? "
   TERRAIN = InputBox(Msg, TERRAIN, "XXXXXX")
   If Not TERRAIN = "XXXXXX" Then
      hexmaptable.Edit
      hexmaptable![TERRAIN] = TERRAIN
      hexmaptable.UPDATE
   End If
End If

On Error Resume Next
   
   'MSG = "GET TERRAIN"
   'RESPONSE = MsgBox(MSG, True)

If Direction = "NE" Then
   CURRENT_MAP = GET_MAP_NORTH_EAST(CURRENT_MAP)
ElseIf Direction = "SE" Then
   CURRENT_MAP = GET_MAP_SOUTH_EAST(CURRENT_MAP)
ElseIf Direction = "SW" Then
   CURRENT_MAP = GET_MAP_SOUTH_WEST(CURRENT_MAP)
ElseIf Direction = "NW" Then
   CURRENT_MAP = GET_MAP_NORTH_WEST(CURRENT_MAP)
ElseIf Direction = "N" Then
   CURRENT_MAP = GET_MAP_NORTH(CURRENT_MAP)
ElseIf Direction = "S" Then
   CURRENT_MAP = GET_MAP_SOUTH(CURRENT_MAP)
End If

hexmaptable.Seek "=", CURRENT_MAP

If hexmaptable.NoMatch Then
   Call ADD_NEW_HEX(CURRENT_MAP, TERRAIN)
   hexmaptable.Seek "=", CURRENT_MAP
End If

If Not IsNull(hexmaptable![TERRAIN]) Then
   CURRENT_TERRAIN = hexmaptable![TERRAIN]
   TERRAIN = hexmaptable![TERRAIN]
   Call SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)
   SHORTTERRAIN = SHORTTERRAIN & ","
   Call GET_SURROUNDING_DATA(CURRENT_MAP)
   Call GET_MOUNTAINS(TERRAIN, CURRENT_MAP)
   Call Check_Surrounding_Terrain(TERRAIN, CURRENT_MAP, "ALPS", "Alps")
   If FLEET = "N" Then
      Call Check_Surrounding_Terrain(TERRAIN, CURRENT_MAP, "OCEAN", "O")
   End If
   Call Check_Surrounding_Terrain(TERRAIN, CURRENT_MAP, "LAKE", "L")
   Call GET_RIVERS(TERRAIN, CURRENT_MAP)
   Call GET_FORDS(TERRAIN, CURRENT_MAP)
   Call GET_PASSES(TERRAIN, CURRENT_MAP)
   Call GET_BEACHS(TERRAIN, CURRENT_MAP)
   Call GET_CLIFFS(TERRAIN, CURRENT_MAP)
   Call GET_ROADS(TERRAIN, CURRENT_MAP)
   Call GET_CANALS(TERRAIN, CURRENT_MAP)
   Call GET_CANYONS(TERRAIN, CURRENT_MAP)
   Call GET_STREAMS(TERRAIN, CURRENT_MAP)
   Call GET_WATERFALLS(TERRAIN, CURRENT_MAP)
   Call GET_CITIES(TERRAIN, CURRENT_MAP)
   Call GET_QUARRIES(TERRAIN, CURRENT_HEX_MAP)
   Call GET_SPRINGS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_SALMON_RUN(TERRAIN, CURRENT_MAP)
   Call GET_WHALING_AREA(TERRAIN, CURRENT_MAP)
   Call GET_FISH_AREA(TERRAIN, CURRENT_MAP)
'   TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
   TRIBESINHEX = "EMPTY"
   If TRIBESINHEX = "EMPTY" Then
      TERRAIN = SHORTTERRAIN & " " & TERRAIN
   Else
      TERRAIN = SHORTTERRAIN & " " & TERRAIN & "+9" & TRIBESINHEX & "+0"
   End If
End If

NE_TERRAIN = ""
N_TERRAIN = ""
NW_TERRAIN = ""
SE_TERRAIN = ""
SW_TERRAIN = ""
S_TERRAIN = ""

END_TIME = Time

If codetrack = 1 Then
   MSG0 = "START_TIME = " & START_TIME & crlf
   MSG1 = "END_TIME = " & END_TIME & crlf
   Response = MsgBox((MSG0 & MSG1), True)
End If

End Sub

Sub GET_WATERFALLS(TERRAIN, CURRENT_MAP)
Dim WATERFALL_FOUND As String

WATERFALL_FOUND = "NO"

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If Mid(hexmaptable![WATERFALLS], 1, 1) = "Y" Then
   TERRAIN = TERRAIN & ",Waterfalls N"
   WATERFALL_FOUND = "YES"
End If
If Mid(hexmaptable![WATERFALLS], 2, 1) = "Y" Then
   If WATERFALL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NE"
   Else
      TERRAIN = TERRAIN & ",Waterfalls NE"
      WATERFALL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![WATERFALLS], 3, 1) = "Y" Then
   If WATERFALL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SE"
   Else
      TERRAIN = TERRAIN & ",Waterfalls SE"
      WATERFALL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![WATERFALLS], 4, 1) = "Y" Then
   If WATERFALL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " S"
   Else
      TERRAIN = TERRAIN & ",Waterfalls S"
      WATERFALL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![WATERFALLS], 5, 1) = "Y" Then
   If WATERFALL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " SW"
   Else
      TERRAIN = TERRAIN & ",Waterfalls SW"
      WATERFALL_FOUND = "YES"
   End If
End If
If Mid(hexmaptable![WATERFALLS], 6, 1) = "Y" Then
   If WATERFALL_FOUND = "YES" Then
      TERRAIN = TERRAIN & " NW"
   Else
      TERRAIN = TERRAIN & ",Waterfalls NW"
      WATERFALL_FOUND = "YES"
   End If
End If


End Sub

Sub GET_WHALING_AREA(TERRAIN, CURRENT_MAP)

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If hexmaptable![WHALE AREA] = "Y" Then
   TERRAIN = TERRAIN & ", Whaling Area (Improved Whaling)"
End If

End Sub

Function MOVE_ROAMING_HERDS()
Dim ROAMING_HERD_HEX As String
Dim HEX_N As String
Dim HEX_NE As String
Dim HEX_SE As String
Dim HEX_S As String
Dim HEX_SW As String
Dim HEX_NW As String
Dim MOVE_HERD As String

DoCmd.Hourglass True

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)
   
' TRIBE MOVEMENT
Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst

Do Until hexmaptable.EOF
   If hexmaptable![ROAMING HERD] = "Y" Then
      ROAMING_HERD_HEX = hexmaptable![MAP]
      HEX_N = GET_MAP_NORTH(hexmaptable![MAP])
      HEX_NE = GET_MAP_NORTH_EAST(hexmaptable![MAP])
      HEX_SE = GET_MAP_SOUTH_EAST(hexmaptable![MAP])
      HEX_S = GET_MAP_SOUTH(hexmaptable![MAP])
      HEX_SW = GET_MAP_SOUTH_WEST(hexmaptable![MAP])
      HEX_NW = GET_MAP_NORTH_WEST(hexmaptable![MAP])
      RIVER_N = Mid(hexmaptable![Borders], 1, 2)
      RIVER_NE = Mid(hexmaptable![Borders], 3, 2)
      RIVER_SE = Mid(hexmaptable![Borders], 5, 2)
      RIVER_S = Mid(hexmaptable![Borders], 7, 2)
      RIVER_SW = Mid(hexmaptable![Borders], 9, 2)
      RIVER_NW = Mid(hexmaptable![Borders], 11, 2)
      PASS_N = Mid(hexmaptable![Borders], 1, 2)
      PASS_NE = Mid(hexmaptable![Borders], 3, 2)
      PASS_SE = Mid(hexmaptable![Borders], 5, 2)
      PASS_S = Mid(hexmaptable![Borders], 7, 2)
      PASS_SW = Mid(hexmaptable![Borders], 9, 2)
      PASS_NW = Mid(hexmaptable![Borders], 11, 2)
      
      hexmaptable.Edit
      hexmaptable![ROAMING HERD] = "N"
      hexmaptable.UPDATE

      MOVED = "N"
      Do Until MOVED = "Y"
         DICE1 = DICE_ROLL("AAA", "AAA")
         If DICE1 <= 17 Then
            If RIVER_N = "NN" Or PASS_N = "NN" Then
               hexmaptable.Seek "=", HEX_N
               MOVE_HERD = CHECK_TERRAIN(hexmaptable![TERRAIN])
               If MOVE_HERD = "Y" Then
                  hexmaptable.Edit
                  hexmaptable![ROAMING HERD] = "Y"
                  hexmaptable.UPDATE
                 MOVED = "Y"
               End If
            End If
         ElseIf DICE1 <= 34 Then
            If RIVER_NE = "NN" Or PASS_NE = "NN" Then
               hexmaptable.Seek "=", HEX_NE
               MOVE_HERD = CHECK_TERRAIN(hexmaptable![TERRAIN])
               If MOVE_HERD = "Y" Then
                  hexmaptable.Edit
                  hexmaptable![ROAMING HERD] = "Y"
                  hexmaptable.UPDATE
                 MOVED = "Y"
               End If
            End If
         ElseIf DICE1 <= 51 Then
            If RIVER_SE = "NN" Or PASS_SE = "NN" Then
               hexmaptable.Seek "=", HEX_SE
               MOVE_HERD = CHECK_TERRAIN(hexmaptable![TERRAIN])
               If MOVE_HERD = "Y" Then
                  hexmaptable.Edit
                  hexmaptable![ROAMING HERD] = "Y"
                  hexmaptable.UPDATE
                 MOVED = "Y"
               End If
            End If
         ElseIf DICE1 <= 68 Then
            If RIVER_S = "NN" Or PASS_S = "NN" Then
               hexmaptable.Seek "=", HEX_S
               MOVE_HERD = CHECK_TERRAIN(hexmaptable![TERRAIN])
               If MOVE_HERD = "Y" Then
                  hexmaptable.Edit
                  hexmaptable![ROAMING HERD] = "Y"
                  hexmaptable.UPDATE
                 MOVED = "Y"
               End If
            End If
         ElseIf DICE1 <= 85 Then
            If RIVER_SW = "NN" Or PASS_SW = "NN" Then
               hexmaptable.Seek "=", HEX_SW
               MOVE_HERD = CHECK_TERRAIN(hexmaptable![TERRAIN])
               If MOVE_HERD = "Y" Then
                  hexmaptable.Edit
                  hexmaptable![ROAMING HERD] = "Y"
                  hexmaptable.UPDATE
                 MOVED = "Y"
               End If
            End If
         ElseIf DICE1 <= 100 Then
            If RIVER_NW = "NN" Or PASS_NW = "NN" Then
               hexmaptable.Seek "=", HEX_NW
               MOVE_HERD = CHECK_TERRAIN(hexmaptable![TERRAIN])
               If MOVE_HERD = "Y" Then
                  hexmaptable.Edit
                  hexmaptable![ROAMING HERD] = "Y"
                  hexmaptable.UPDATE
                 MOVED = "Y"
               End If
            End If
         End If

      Loop
   
   End If

   hexmaptable.MoveNext
Loop

hexmaptable.Close


DoCmd.Hourglass False


End Function

Function SCOUT_MOVEMENT(SCREEN As String)
On Error GoTo SCOUT_MOVEMENT_EXIT
Dim SLENGTH As Long
Dim Mycontrol As Control
Dim DICE_TRIBE As Long
Dim Scout_Movement_Allowed(8) As String
Dim Scout_Direction(8) As String
Dim Number_Of_Scouts(8) As Integer
Dim Number_Of_Horses(8) As Integer
Dim Number_Of_Elephants(8) As Integer
Dim Number_Of_Camels(8) As Integer
Dim SCOUTS_MISSION(8) As String
Dim ACTIVITY_LABEL As String

codetrack = 0
crlf = Chr(13) & Chr(10)

TM_POS = "START"

START_TIME = Time
Scouting = "YES"

USE_SCREEN = SCREEN
DoCmd.Hourglass True

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESINFO.MoveFirst
TRIBESINFO.index = "PRIMARYKEY"

Set SCOUT_MOVEMENT_TABLE = TVDBGM.OpenRecordset("SCOUT_MOVEMENT")
SCOUT_MOVEMENT_TABLE.index = "PRIMARYKEY"
SCOUT_MOVEMENT_TABLE.MoveFirst

Set Scout_Result = TVDB.OpenRecordset("Scouting_Results")
Scout_Result.index = "PRIMARYKEY"

Set Movement_Trace = TVDB.OpenRecordset("Movement_Trace")
Movement_Trace.MoveFirst
Movement_Trace.index = "PRIMARYKEY"

'Load screen data into table before processing table
If USE_SCREEN = "Y" Then
   Set MOVEFORM = Forms![SCOUT MOVEMENT]
   MOVE_TRIBE = MOVEFORM![TRIBE NAME]
   MOVE_CLAN = "0" & Mid(MOVE_TRIBE, 2, 3)
   ' this loads the table with the information from the screen
   cnt1 = 1
   SCOUT_MOVEMENT_TABLE.Seek "=", MOVE_TRIBE
   If SCOUT_MOVEMENT_TABLE.EOF Then
      'ignore
   Else
      Do
        SCOUT_MOVEMENT_TABLE.Delete
        SCOUT_MOVEMENT_TABLE.MoveNext
        If Not SCOUT_MOVEMENT_TABLE![TRIBE] = MOVE_TRIBE Then
           Exit Do
        End If
             
      Loop
   End If
   Do While cnt1 < 9
      SCOUT_MOVEMENT_TABLE.AddNew
      SCOUT_MOVEMENT_TABLE![TRIBE] = MOVE_TRIBE
      stext1 = "Scout" & CStr(cnt1) & "Move01"
      stext2 = "Scout" & CStr(cnt1) & "Move02"
      stext3 = "Scout" & CStr(cnt1) & "Move03"
      stext4 = "Scout" & CStr(cnt1) & "Move04"
      stext5 = "Scout" & CStr(cnt1) & "Move05"
      stext6 = "Scout" & CStr(cnt1) & "Move06"
      stext7 = "Scout" & CStr(cnt1) & "Move07"
      stext8 = "Scout" & CStr(cnt1) & "Move08"
      SCOUT_MOVEMENT_TABLE![Movement1] = MOVEFORM(stext1).Value
      SCOUT_MOVEMENT_TABLE![Movement2] = MOVEFORM(stext2).Value
      SCOUT_MOVEMENT_TABLE![Movement3] = MOVEFORM(stext3).Value
      SCOUT_MOVEMENT_TABLE![Movement4] = MOVEFORM(stext4).Value
      SCOUT_MOVEMENT_TABLE![Movement5] = MOVEFORM(stext5).Value
      SCOUT_MOVEMENT_TABLE![Movement6] = MOVEFORM(stext6).Value
      SCOUT_MOVEMENT_TABLE![Movement7] = MOVEFORM(stext7).Value
      SCOUT_MOVEMENT_TABLE![Movement8] = MOVEFORM(stext8).Value
      stext = "SCOUTS" & CStr(cnt1)
      SCOUT_MOVEMENT_TABLE![No_of_Scouts] = MOVEFORM(stext).Value
      stext = "HORSES" & CStr(cnt1)
      SCOUT_MOVEMENT_TABLE![No_of_Horses] = MOVEFORM(stext).Value
      stext = "Elephants" & CStr(cnt1)
      SCOUT_MOVEMENT_TABLE![No_of_Elephants] = MOVEFORM(stext).Value
      stext = "Camels" & CStr(cnt1)
      SCOUT_MOVEMENT_TABLE![No_of_Camels] = MOVEFORM(stext).Value
      stext = "MISSION" & CStr(cnt1)
      SCOUT_MOVEMENT_TABLE![MISSION] = MOVEFORM(stext).Value
      SCOUT_MOVEMENT_TABLE![PROCESSED] = "N"
      SCOUT_MOVEMENT_TABLE.UPDATE
      cnt1 = cnt1 + 1
      If cnt1 > 8 Then
         Exit Do
      End If
   Loop
End If

'if USE_SCREEN = "Y" then need to start at the right record and at the end of the records, exit
If USE_SCREEN = "Y" Then
   SCOUT_MOVEMENT_TABLE.Seek "=", MOVE_TRIBE
Else
   ' reset the SCOUT_MOVEMENT_TABLE back to the start
   SCOUT_MOVEMENT_TABLE.MoveFirst
End If

TM_POS = "START OF TABLE LOOP"
Do Until SCOUT_MOVEMENT_TABLE.EOF

   If USE_SCREEN = "Y" Then
      If SCOUT_MOVEMENT_TABLE![TRIBE] <> MOVE_TRIBE Then
         GoTo End_Loop
      End If
   End If
   If SCOUT_MOVEMENT_TABLE![PROCESSED] = "Y" Then
      GoTo End_Loop
   End If
   If SCOUT_MOVEMENT_TABLE![No_of_Scouts] = 0 Or IsNull(SCOUT_MOVEMENT_TABLE![No_of_Scouts]) Then
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![PROCESSED] = "Y"
      SCOUT_MOVEMENT_TABLE.UPDATE
      GoTo End_Loop
   End If
   MOVE_TRIBE = SCOUT_MOVEMENT_TABLE![TRIBE]
   MOVE_CLAN = "0" & Mid(MOVE_TRIBE, 2, 3)

   SKILL_MOVE_TRIBE = Unit_Check("TRIBE", MOVE_TRIBE)
   DICE_TRIBE = Unit_Check("DICE", MOVE_TRIBE)

   ' TRIBE MOVEMENT

   TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_CLAN

   If IsNull(TRIBESINFO![TRUCES]) Then
      Truced_Clans = "EMPTY"
   Else
      Truced_Clans = TRIBESINFO![TRUCES]
   End If

   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE

   If TRIBESINFO.NoMatch Then
      ' not a valid tribe therefore the loop if screen
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![PROCESSED] = "Y"
      SCOUT_MOVEMENT_TABLE.UPDATE
      GoTo End_Loop
   End If

   TCLANNUMBER = MOVE_CLAN
   CURRENT_MAP = TRIBESINFO![CURRENT HEX]

   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      GOODS_TRIBE = MOVE_TRIBE
   End If

   Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_Goods")
   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.MoveFirst

   Set SCOUTING_TABLE = TVDB.OpenRecordset("SCOUTING_FINDS")

   Set Globaltable = TVDBGM.OpenRecordset("GLOBAL")
   Globaltable.index = "PRIMARYKEY"
   Globaltable.MoveFirst

   TURN_CURRENT = Globaltable![CURRENT TURN]

   Set TERRAINTABLE = TVDB.OpenRecordset("VALID_TERRAIN")
   TERRAINTABLE.index = "PRIMARYKEY"

   Set HEXMAPCITY = TVDBGM.OpenRecordset("HEX_MAP_CITY")
   HEXMAPCITY.index = "PRIMARYKEY"

   Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
   HEXMAPMINERALS.index = "PRIMARYKEY"

   Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
   HEXMAPCONST.index = "MAP"

   Call Obtain_Skill_Levels

   Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
   hexmaptable.MoveFirst
   hexmaptable.index = "PRIMARYKEY"
   hexmaptable.Seek "=", CURRENT_MAP

   TM_POS = "DETERMINE WEATHER COST"
   ' DETERMINE WEATHER COST

   If hexmaptable![WEATHER_ZONE] = "GREEN" Then
      WEATHER = Globaltable![Zone1]
      wind = Globaltable![Wind1]
      WIND_DIRECTION = Globaltable![DIRECTION1]
   ElseIf hexmaptable![WEATHER_ZONE] = "RED" Then
      WEATHER = Globaltable![Zone2]
      wind = Globaltable![Wind2]
      WIND_DIRECTION = Globaltable![DIRECTION2]
   ElseIf hexmaptable![WEATHER_ZONE] = "ORANGE" Then
      WEATHER = Globaltable![Zone3]
      wind = Globaltable![Wind3]
      WIND_DIRECTION = Globaltable![DIRECTION3]
   ElseIf hexmaptable![WEATHER_ZONE] = "YELLOW" Then
      WEATHER = Globaltable![Zone4]
      wind = Globaltable![Wind4]
      WIND_DIRECTION = Globaltable![DIRECTION4]
   ElseIf hexmaptable![WEATHER_ZONE] = "BLUE" Then
      WEATHER = Globaltable![Zone5]
      wind = Globaltable![Wind5]
      WIND_DIRECTION = Globaltable![DIRECTION5]
   ElseIf hexmaptable![WEATHER_ZONE] = "BROWN" Then
      WEATHER = Globaltable![Zone6]
      wind = Globaltable![Wind6]
      WIND_DIRECTION = Globaltable![DIRECTION6]
   End If

   If codetrack = 1 Then
      MSG1 = "WEATHER = " & WEATHER & crlf
      MSG2 = "wind = " & wind & crlf
      MSG3 = "WIND_DIRECTION = " & WIND_DIRECTION & crlf
      Response = MsgBox((MSG1 & MSG2 & MSG3), True)
   End If
   
   Globaltable.Close
   FLEET = "N"
 
   END_TIME = Time

   If codetrack = 2 Then
      TOTAL_TIME = END_TIME - START_TIME
      MSG1 = "START TIME = " & START_TIME & crlf
      MSG2 = "END_TIME = " & END_TIME & crlf
      MSG3 = "TOTAL TIME FOR SETUP = " & TOTAL_TIME & crlf
      Response = MsgBox((MSG1 & MSG2 & MSG3), True)
   End If

   WHICH_SCOUT = "SCOUT 1"

   MOVEMENT_ITERATIONS = 1

   Movement_Trace.MoveFirst
   Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, WHICH_SCOUT, MOVEMENT_ITERATIONS

   If Movement_Trace.NoMatch Then
      Movement_Trace.AddNew
      Movement_Trace![TRIBE] = MOVE_TRIBE
      Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
      Movement_Trace![Tribe_Or_Scout] = WHICH_SCOUT
      Movement_Trace![Movement_Number] = MOVEMENT_ITERATIONS
      Movement_Trace![STARTING_HEX] = CURRENT_MAP
      Movement_Trace![Target_Hex] = "??"
      Movement_Trace![Target_Terrain] = "??"
      Movement_Trace![Direction] = "??"
      Movement_Trace![WEATHER] = WEATHER
      Movement_Trace![wind] = wind
      Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
      Movement_Trace![Current_Movement_Points] = 0
      Movement_Trace![Current_Movement_Cost] = 0
      Movement_Trace![CAN_GROUP_MOVE] = "U"
      Movement_Trace![NO_MOVEMENT_REASON] = "??"
      Movement_Trace![SCOUTS] = 0
      Movement_Trace![HORSES] = 0
      Movement_Trace![Elephants] = 0
      Movement_Trace.UPDATE
      Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, WHICH_SCOUT, MOVEMENT_ITERATIONS
   Else
      ' can only happen on a repeat move or if the player backtracks
      Movement_Trace.Delete
      Movement_Trace.AddNew
      Movement_Trace![TRIBE] = MOVE_TRIBE
      Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
      Movement_Trace![Tribe_Or_Scout] = WHICH_SCOUT
      Movement_Trace![Movement_Number] = MOVEMENT_ITERATIONS
      Movement_Trace![STARTING_HEX] = CURRENT_MAP
      Movement_Trace![Target_Hex] = "??"
      Movement_Trace![Target_Terrain] = "??"
      Movement_Trace![Direction] = "??"
      Movement_Trace![WEATHER] = WEATHER
      Movement_Trace![wind] = wind
      Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
      Movement_Trace![Current_Movement_Points] = 0
      Movement_Trace![Current_Movement_Cost] = 0
      Movement_Trace![CAN_GROUP_MOVE] = "U"
      Movement_Trace![NO_MOVEMENT_REASON] = "??"
      Movement_Trace![SCOUTS] = 0
      Movement_Trace![HORSES] = 0
      Movement_Trace![Elephants] = 0
      Movement_Trace.UPDATE
      Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, WHICH_SCOUT, MOVEMENT_ITERATIONS
   End If
   ' SCOUT MOVEMENT

   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "FINISHED", "SPYGLASS"
   If TRIBESGOODS.NoMatch Then
      SPYGLASSES = "N"
   ElseIf TRIBESGOODS![ITEM_NUMBER] > 0 Then
         SPYGLASSES = "Y"
   Else
      SPYGLASSES = "N"
   End If
  
   LineI = 1


   TM_POS = "VARIABLE LOAD"

   'Here is where I should load all of the variable, regardless of source
   ' got everything except for movement 1 through 8
   ' throw this into a load routine
 
   'clear variables
   cnt1 = 1
   Do While cnt1 < 9
      Scout_Movement_Allowed(cnt1) = "NO"
      Number_Of_Scouts(cnt1) = 0
      Number_Of_Horses(cnt1) = 0
      Number_Of_Elephants(cnt1) = 0
      Number_Of_Camels(cnt1) = 0
      SCOUTS_MISSION(cnt1) = "EMPTY"
      cnt2 = 1
      Do While cnt2 < 9
         Scouting_Movement(cnt1, cnt2) = "EMPTY"
         cnt2 = cnt2 + 1
      Loop
      cnt1 = cnt1 + 1
   Loop

   cnt1 = 1

   ' load variables from table
   Do While cnt1 < 9
     
      If SCOUT_MOVEMENT_TABLE![Movement1] = "EMPTY" Then
         Scout_Movement_Allowed(cnt1) = "NO"
         Exit Do
      Else
         Scout_Movement_Allowed(cnt1) = "YES"
         If IsNull(SCOUT_MOVEMENT_TABLE![No_of_Scouts]) Then
            Number_Of_Scouts(cnt1) = 0
            Scout_Movement_Allowed(cnt1) = "NO"
         Else
            Number_Of_Scouts(cnt1) = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
         End If
         If IsNull(SCOUT_MOVEMENT_TABLE![No_of_Horses]) Then
            Number_Of_Horses(cnt1) = 0
         Else
            Number_Of_Horses(cnt1) = SCOUT_MOVEMENT_TABLE![No_of_Horses]
         End If
         If IsNull(SCOUT_MOVEMENT_TABLE![No_of_Elephants]) Then
            Number_Of_Elephants(cnt1) = 0
         Else
            Number_Of_Elephants(cnt1) = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
         End If
         If IsNull(SCOUT_MOVEMENT_TABLE![No_of_Camels]) Then
            Number_Of_Camels(cnt1) = 0
         Else
            Number_Of_Camels(cnt1) = SCOUT_MOVEMENT_TABLE![No_of_Camels]
         End If
         If IsNull(SCOUT_MOVEMENT_TABLE![MISSION]) Then
            SCOUTS_MISSION(cnt1) = "EMPTY"
         Else
            SCOUTS_MISSION(cnt1) = SCOUT_MOVEMENT_TABLE![MISSION]
         End If
         cnt2 = 1
         Do While cnt2 < 9
            stext = "MOVEMENT" & CStr(cnt2)
   
            If IsNull(SCOUT_MOVEMENT_TABLE(stext).Value) Then
               Scouting_Movement(cnt1, cnt2) = "EMPTY"
            Else
               Scouting_Movement(cnt1, cnt2) = SCOUT_MOVEMENT_TABLE(stext).Value
            End If
            cnt2 = cnt2 + 1
         Loop
   
         cnt1 = cnt1 + 1
         If cnt1 > 8 Then
            Exit Do
         End If
      End If
   Loop
    
HORSES = 0
Elephants = 0
SCOUT_NUMBER = 0

TM_POS = "START SCOUT LOOP"
Do
  START_TIME = Time

  LineI = 1

  SCOUT_NUMBER = SCOUT_NUMBER + 1
  WHICH_SCOUT = "SCOUT" & Str(SCOUT_NUMBER)
  MINERALSINHEX = ""
  CURRENT_MAP = TRIBESINFO![CURRENT HEX]
  SCOUTS_USED = Number_Of_Scouts(SCOUT_NUMBER)
  HORSES_USED = Number_Of_Horses(SCOUT_NUMBER)
  ELEPHANTS_USED = Number_Of_Elephants(SCOUT_NUMBER)
  CAMELS_USED = Number_Of_Camels(SCOUT_NUMBER)
       
 
  If Scout_Movement_Allowed(SCOUT_NUMBER) = "YES" Then
     ' GET MOVEMENT POINTS
     Call DETERMINE_MOVEMENT_POINTS("Y")
     'get the relevant movement
     
'     SCOUTMOVEMENT = Scout_Direction(SCOUT_NUMBER)
     MOVEMENT_LINE = "Scout" + Str(SCOUT_NUMBER) + ":Scout "
     ACTIVITY_LABEL = "SCOUT " + Str(SCOUT_NUMBER) + " MOVEMENT"

     NO_MOVEMENT_REASON = ""
     NEW_ORDERS = ""
     cnt1 = 1
     MOVEMENT_ITERATIONS = 0
     TM_POS = "START SCOUT MOVEMENT LOOP"
     Do While cnt1 < 9
        MOVEMENT_ITERATIONS = MOVEMENT_ITERATIONS + 1
        Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, WHICH_SCOUT, MOVEMENT_ITERATIONS

        If Movement_Trace.NoMatch Then
           Movement_Trace.AddNew
           Movement_Trace![TRIBE] = MOVE_TRIBE
           Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
           Movement_Trace![Tribe_Or_Scout] = WHICH_SCOUT
           Movement_Trace![Movement_Number] = MOVEMENT_ITERATIONS
           Movement_Trace![STARTING_HEX] = CURRENT_MAP
           Movement_Trace![Target_Hex] = "??"
           Movement_Trace![Target_Terrain] = "??"
           Movement_Trace![Direction] = "??"
           Movement_Trace![WEATHER] = WEATHER
           Movement_Trace![wind] = wind
           Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
           Movement_Trace![Current_Movement_Points] = MOVEMENT_POINTS
           Movement_Trace![Current_Movement_Cost] = 0
           Movement_Trace![CAN_GROUP_MOVE] = "U"
           Movement_Trace![NO_MOVEMENT_REASON] = "??"
           Movement_Trace![SCOUTS] = SCOUTS_USED
           Movement_Trace![HORSES] = HORSES_USED
           Movement_Trace![Elephants] = ELEPHANTS_USED
           Movement_Trace.UPDATE
           Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, WHICH_SCOUT, MOVEMENT_ITERATIONS
        Else
           ' can only happen on a repeat move or if the player backtracks
           Movement_Trace.Delete
           Movement_Trace.AddNew
           Movement_Trace![TRIBE] = MOVE_TRIBE
           Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
           Movement_Trace![Tribe_Or_Scout] = WHICH_SCOUT
           Movement_Trace![Movement_Number] = MOVEMENT_ITERATIONS
           Movement_Trace![STARTING_HEX] = CURRENT_MAP
           Movement_Trace![Target_Hex] = "??"
           Movement_Trace![Target_Terrain] = "??"
           Movement_Trace![Direction] = "??"
           Movement_Trace![WEATHER] = WEATHER
           Movement_Trace![wind] = wind
           Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
           Movement_Trace![Current_Movement_Points] = MOVEMENT_POINTS
           Movement_Trace![Current_Movement_Cost] = 0
           Movement_Trace![CAN_GROUP_MOVE] = "U"
           Movement_Trace![NO_MOVEMENT_REASON] = "??"
           Movement_Trace![SCOUTS] = SCOUTS_USED
           Movement_Trace![HORSES] = HORSES_USED
           Movement_Trace![Elephants] = ELEPHANTS_USED
           Movement_Trace.UPDATE
           Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, WHICH_SCOUT, MOVEMENT_ITERATIONS
        End If
        If Not IsNull(NEW_ORDERS) And Not (NEW_ORDERS = "") Then
           Direction = NEW_ORDERS
           ORIG_Direction = NEW_ORDERS
           NEW_ORDERS = ""
        Else
           'Direction = Scouting_Movement(SCOUT_NUMBER, cnt1)
           'ORIG_Direction = Scouting_Movement(SCOUT_NUMBER, cnt1)
           Call GET_NEXT_SCOUT_MOVE
        End If
        If Direction = "NL" Then
           Direction = "N"
        ElseIf Direction = "NEL" Then
           Direction = "NE"
        ElseIf Direction = "SEL" Then
           Direction = "SE"
        ElseIf Direction = "SL" Then
           Direction = "S"
        ElseIf Direction = "SWL" Then
           Direction = "SW"
        ElseIf Direction = "NWL" Then
           Direction = "NW"
        End If
        
        If Direction = "STILL" Then
           GROUP_MOVE = "Y"
        ElseIf Direction = "STOP" Then
           Exit Do
        ElseIf Direction = "EMPTY" Then
           If MOVEMENT_ITERATIONS = 1 Then
              MOVEMENT_LINE = "EMPTY"
           End If
           Exit Do
        ElseIf Direction = "HALT" Then
           Exit Do
        Else
           Call CAN_GROUP_MOVE(CURRENT_MAP, "Y")
        End If
        If GROUP_MOVE = "Y" Then
           If Direction = "STILL" Then
              cnt1 = 10
           Else
              Call GET_TERRAIN(Direction, TERRAIN, CURRENT_MAP)
           End If
           If codetrack = 1 Then
              MSG0 = "SCOUT " + SCOUT_NUMBER + ", 1ST CHECK FOR MOVEMENT LINE > 150 " & crlf
              MSG1 = "DIRECTION = " & Direction & crlf
              MSG2 = "TERRAIN = " & TERRAIN & crlf
              MSG3 = "MOVEMENT_LINE = " & MOVEMENT_LINE & crlf
              Response = MsgBox((MSG0 & MSG1 & MSG2 & MSG3), True)
           End If
           ' FIND MINERALS WITH EVERY MOVEMENT AT ANYTIME
           Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
           HEXMAPMINERALS.index = "PRIMARYKEY"
           HEXMAPMINERALS.MoveFirst
           HEXMAPMINERALS.Seek "=", CURRENT_MAP
           If Not HEXMAPMINERALS.NoMatch Then
               If Not IsNull(HEXMAPMINERALS![ORE_TYPE]) Then
                   MINERALSINHEX = "Find " & HEXMAPMINERALS![ORE_TYPE]
               End If
               Call Get_Research_Data(TCLANNUMBER, TTRIBENUMBER, "Geologists")
               If RESEARCH_FOUND = "Y" Then
                   If Not IsNull(HEXMAPMINERALS![SECOND_ORE]) Then
                       MINERALSINHEX = MINERALSINHEX & ", " & HEXMAPMINERALS![ORE_TYPE]
                   End If
                   If Not IsNull(HEXMAPMINERALS![THIRD_ORE]) Then
                       MINERALSINHEX = MINERALSINHEX & ", " & HEXMAPMINERALS![THIRD_ORE]
                   End If
                   If Not IsNull(HEXMAPMINERALS![FORTH_ORE]) Then
                       MINERALSINHEX = MINERALSINHEX & ", " & HEXMAPMINERALS![FORTH_ORE]
                   End If
               End If
           End If
           If Left(MINERALSINHEX, 4) = "Find" Then
              MOVEMENT_LINE = MOVEMENT_LINE & Direction & "-" & TERRAIN & "," & MINERALSINHEX & "\"
           Else
              MOVEMENT_LINE = MOVEMENT_LINE & Direction & "-" & TERRAIN & "\"
           End If
           MINERALSINHEX = ""
           TERRAIN = ""
           Direction = ""
           If FLEET = "Y" Or (Right(CURRENT_TERRAIN, 9) = "MOUNTAINS") Or (Right(CURRENT_TERRAIN, 2) = "MT") Then
              If SPYGLASSES = "Y" Then
                 GET_SURROUNDING_TERRAIN (CURRENT_MAP)
                 hexmaptable.MoveFirst
                 hexmaptable.Seek "=", CURRENT_MAP
              End If
           End If

           If codetrack = 1 Then
              MSG0 = "SCOUT" + Str(SCOUT_NUMBER) + ", 2ND CHECK FOR MOVEMENT LINE > 150 " & crlf
              MSG1 = "DIRECTION = " & Direction & crlf
              MSG2 = "TERRAIN = " & TERRAIN & crlf
              MSG3 = "MOVEMENT_LINE = " & MOVEMENT_LINE & crlf
              Response = MsgBox((MSG0 & MSG1 & MSG2 & MSG3), True)
           End If
           
           TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
    
           If Not TRIBESINHEX = "EMPTY" Then
              MOVEMENT_LINE = MOVEMENT_LINE & Direction & "-" & TERRAIN & ", " & TRIBESINHEX & "\"
              Scout_Result.AddNew
              Scout_Result![TRIBE] = MOVE_TRIBE
              Scout_Result![SCOUT] = SCOUT_NUMBER
              Scout_Result![MISSION] = "None"
              Scout_Result![FOUND] = TRIBESINHEX
              Scout_Result![Results] = "Were found in the hex"
              Scout_Result.UPDATE
           End If
           MINERALSINHEX = ""
           TERRAIN = ""
           Direction = ""
        ElseIf MOVEMENT_POINTS >= 3 Then
           ' no instructions
           NEW_ORDERS = "STOP"
           If NEW_ORDERS = "STOP" Then
              MOVEMENT_LINE = MOVEMENT_LINE & "," & NO_MOVEMENT_REASON & ","
              cnt1 = 10
           Else
              Direction = NEW_ORDERS
              ORIG_Direction = NEW_ORDERS
           End If
        Else
          ' group cant move and MP's are lower than 3
          MOVEMENT_LINE = MOVEMENT_LINE & "," & NO_MOVEMENT_REASON & ","
          cnt1 = 10
        End If
        'MSG = "TERRAIN = " & TERRAIN
        'RESPONSE = MsgBox(MSG, True)
        If ORIG_Direction = "FRR" Or ORIG_Direction = "FRL" _
        Or ORIG_Direction = "FMR" Or ORIG_Direction = "FML" _
        Or ORIG_Direction = "FLR" Or ORIG_Direction = "FLL" _
        Or ORIG_Direction = "FOR" Or ORIG_Direction = "FOL" _
        Or ORIG_Direction = "FCR" Or ORIG_Direction = "FCL" _
        Or ORIG_Direction = "FL" Or ORIG_Direction = "FO" _
        Or ORIG_Direction = "NL" Or ORIG_Direction = "NEL" _
        Or ORIG_Direction = "SEL" Or ORIG_Direction = "SL" _
        Or ORIG_Direction = "SWL" Or ORIG_Direction = "NWL" Then
          ' DO NOTHING
        Else
           cnt1 = cnt1 + 1
        End If
        If cnt1 > 8 Then
           Exit Do
        End If
     Loop
     'TRIBESINHEX = WHO_IS_IN_HEX(CLAN, tribe, CURRENT_MAP, N)
     'If Not TRIBESINHEX = "EMPTY" Then
     '   MOVEMENT_LINE = MOVEMENT_LINE & ", " & TRIBESINHEX
     'End If

     ' CHECK FOR SCOUTING FIND

     If MOVEMENT_ITERATIONS = 1 And Direction = "EMPTY" Then
        'do nothing
     Else
        SCOUT_MISSION = SCOUTS_MISSION(SCOUT_NUMBER)
        Call CALC_SCOUTING_FINDS
     End If
   Else
      MOVEMENT_LINE = "EMPTY"
   End If

   ACTIVITY_LABEL = "SCOUT " + Str(SCOUT_NUMBER) + " MOVEMENT"
   
   If Left(MOVEMENT_LINE, 5) = "EMPTY" Then
      'dont write output
   Else
      If Len(MOVEMENT_LINE) > 0 Then
         Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, ACTIVITY_LABEL, LineI, MOVEMENT_LINE, "No")
      End If
   End If
   
   END_TIME = Time

   If codetrack = 2 Then
      TOTAL_TIME = END_TIME - START_TIME
      MSG1 = "START TIME = " & START_TIME & crlf
      MSG2 = "END_TIME = " & END_TIME & crlf
      MSG3 = "TOTAL TIME FOR SCOUT1 = " & TOTAL_TIME & crlf
      Response = MsgBox((MSG1 & MSG2 & MSG3), True)
   End If
   
   'update scout as processed
   If SCOUT_MOVEMENT_TABLE.NoMatch Then
      Exit Do
   Else
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![PROCESSED] = "Y"
      SCOUT_MOVEMENT_TABLE.UPDATE
      SCOUT_MOVEMENT_TABLE.MoveNext
   End If
   If Not (SCOUT_MOVEMENT_TABLE![TRIBE] = MOVE_TRIBE) Then
      SCOUT_MOVEMENT_TABLE.MovePrevious
      Exit Do
   End If
   
   If SCOUT_NUMBER >= 8 Then
      Exit Do
   End If
   ' end of tribe scouting loop
Loop

' end of table loop
End_Loop:
SCOUT_MOVEMENT_TABLE.MoveNext
If SCOUT_MOVEMENT_TABLE.EOF Then
   Exit Do
End If
   
Loop

TM_POS = "CLOSE FILES ETC"
TRIBESINFO.Close
Movement_Trace.Close
SCOUT_MOVEMENT_TABLE.Close

If SCREEN = "Y" Then
   DoCmd.Close A_FORM, "SCOUT MOVEMENT"
   DoCmd.OpenForm "SCOUT MOVEMENT"
End If
   

SCOUT_MOVEMENT_EXIT_CLOSE:
   DoCmd.Hourglass False
   Exit Function


SCOUT_MOVEMENT_EXIT:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Msg = "Error Occurred in section " & TM_POS
   MsgBox (Msg)
   Resume SCOUT_MOVEMENT_EXIT_CLOSE
End If

End Function

Sub tt9213874kjhf()
    Call SHORTEN_TERRAIN("", "LOW GLACIER MOUNTAINS")
End Sub

Sub SHORTEN_TERRAIN(SHORTTERRAIN, TERRAIN)

SHORTTERRAIN = Nz(ELookup("TERR_SHORT", _
                        "VALID_TERRAIN", _
                        "TERRAIN = '" & TERRAIN & "'"), "")

'If TERRAIN = "ARID" Then
'   SHORTTERRAIN = "AR"
'ElseIf TERRAIN = "BAMBOO" Then
'   SHORTTERRAIN = "BAM"
'ElseIf TERRAIN = "BRUSH FLAT" Then
'   SHORTTERRAIN = "BR"
'ElseIf TERRAIN = "BRUSH HILLS" Then
'   SHORTTERRAIN = "BH"
'ElseIf TERRAIN = "CONIFER" Then
'   SHORTTERRAIN = "CF"
'ElseIf TERRAIN = "DECIDUOUS" Then
'   SHORTTERRAIN = "DF"
'ElseIf TERRAIN = "DECIDUOUS FLAT" Then
'   SHORTTERRAIN = "DF"
'ElseIf TERRAIN = "CONIFER HILLS" Then
'   SHORTTERRAIN = "CH"
'ElseIf TERRAIN = "DECIDUOUS FOREST" Then
'   SHORTTERRAIN = "DF"
'ElseIf TERRAIN = "DECIDUOUS HILLS" Then
'   SHORTTERRAIN = "DH"
'ElseIf TERRAIN = "HARDWOOD FOREST" Then
'   SHORTTERRAIN = "HF"
'ElseIf TERRAIN = "HARDWOOD HILLS" Then
'   SHORTTERRAIN = "HH"
'ElseIf TERRAIN = "JUNGLE" Then
'   SHORTTERRAIN = "JG"
'ElseIf TERRAIN = "JUNGLE HILLS" Then
'   SHORTTERRAIN = "JH"
'ElseIf TERRAIN = "ROCKY HILLS" Then
'   SHORTTERRAIN = "RH"
'ElseIf TERRAIN = "SWAMP" Then
'   SHORTTERRAIN = "SW"
'ElseIf TERRAIN = "PRAIRIE" Then
'   SHORTTERRAIN = "PR"
'ElseIf TERRAIN = "TUNDRA" Then
'   SHORTTERRAIN = "TU"
'ElseIf TERRAIN = "GRASSY HILLS" Then
'   SHORTTERRAIN = "GH"
'ElseIf TERRAIN = "DESERT" Then
'   SHORTTERRAIN = "DE"
'ElseIf TERRAIN = "SNOWY HILLS" Then
'   SHORTTERRAIN = "SH"
'ElseIf TERRAIN = "HIGH SNOWY MOUNTAINS" Then
'   SHORTTERRAIN = "HM"
'ElseIf TERRAIN = "HIGH SNOWY MT" Then
'   SHORTTERRAIN = "HM"
'ElseIf TERRAIN = "LOW CONIFER MOUNTAINS" Then
'   SHORTTERRAIN = "LCM"
'ElseIf TERRAIN = "LOW CONIFER MT" Then
'   SHORTTERRAIN = "LCM"
'ElseIf TERRAIN = "LOW SNOWY MOUNTAINS" Then
'   SHORTTERRAIN = "LSM"
'ElseIf TERRAIN = "LOW SNOWY MT" Then
'   SHORTTERRAIN = "LSM"
'ElseIf TERRAIN = "LOW VOLCANO MOUNTAINS" Then
'   SHORTTERRAIN = "LVM"
'ElseIf TERRAIN = "LOW VOLCANO MT" Then
'   SHORTTERRAIN = "LVM"
'ElseIf TERRAIN = "LOW JUNGLE MOUNTAINS" Then
'   SHORTTERRAIN = "LJM"
'ElseIf TERRAIN = "LOW JUNGLE MT" Then
'   SHORTTERRAIN = "LJM"
'ElseIf TERRAIN = "OCEAN" Then
'   SHORTTERRAIN = "O"
'ElseIf TERRAIN = "LAKE" Then
'   SHORTTERRAIN = "L"
'ElseIf TERRAIN = "MANGROVE SWAMPS" Then
'   SHORTTERRAIN = "MS"
'ElseIf TERRAIN = "MANGROVE SWAMP" Then
'   SHORTTERRAIN = "MS"
'ElseIf TERRAIN = "POLAR ICE" Then
'   SHORTTERRAIN = "PI"
'Else
'   Msg = "TERRAIN NOT IN SHORTEN_TERRAIN SUB - " & TERRAIN
'   Response = MsgBox(Msg, True)
'End If

TERRAIN = ""

End Sub

Function Tribe_Movement(SCREEN As String)
On Error GoTo TRIBE_MOVEMENT_EXIT
Dim Pass_One As Boolean
Dim Multiple_Follows As Boolean
Dim Initial_Tribe As String

TM_POS = "START"

codetrack = 0
crlf = Chr(13) & Chr(10)



USE_SCREEN = SCREEN

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set TRIBESINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESINFO.MoveFirst
TRIBESINFO.index = "PRIMARYKEY"

Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXMAPCONST.index = "CLANCONST"

Set ConstructionTable = TVDBGM.OpenRecordset("UNDER_CONSTRUCTION")
ConstructionTable.index = "TRIBE"

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.MoveFirst
hexmaptable.index = "PRIMARYKEY"

Set TERRAINTABLE = TVDB.OpenRecordset("VALID_TERRAIN")
TERRAINTABLE.index = "PRIMARYKEY"

Set Movement_Trace = TVDB.OpenRecordset("Movement_Trace")
Movement_Trace.MoveFirst
Movement_Trace.index = "PRIMARYKEY"

Set TRIBE_MOVEMENT_TABLE = TVDBGM.OpenRecordset("PROCESS_TRIBE_MOVEMENT")
TRIBE_MOVEMENT_TABLE.index = "PRIMARYKEY"
TRIBE_MOVEMENT_TABLE.MoveFirst

Pass_One = True
Multiple_Follows = False

Movement_Table_Start:
Do Until TRIBE_MOVEMENT_TABLE.EOF
   MOVEMENT_LINE = ""
   NO_MOVEMENT_REASON = ""
   
If USE_SCREEN = "Y" Then
   MOVE_CLAN = "0" & Mid(Forms![TRIBE MOVEMENT]![TRIBE NAME], 2, 3)
   MOVE_TRIBE = Forms![TRIBE MOVEMENT]![TRIBE NAME]
   Set MOVEFORM = Forms![TRIBE MOVEMENT]
   MOVEMENT_ORDERS(1) = MOVEFORM![Movement01]
   If MOVEMENT_ORDERS(1) = "FOLLOW" Then
      'check that it starts with a number
      'check if its a hex - might be in as a 6 or 7 character format (with or without the space)
      'check if its a city -
      CLAN = "0" & Mid(MOVEFORM![Follow_Tribe], 2, 3)
      TRIBE = MOVEFORM![Follow_Tribe]
   End If
   Pass_One = False
   Multiple_Follows = False
Else
   If TRIBE_MOVEMENT_TABLE![PROCESSED] = "Y" Or IsNull(TRIBE_MOVEMENT_TABLE![TRIBE]) Then
      GoTo End_Loop
   End If
   MOVE_TRIBE = TRIBE_MOVEMENT_TABLE![TRIBE]
   MOVE_CLAN = "0" & Mid(MOVE_TRIBE, 2, 3)
   If Not IsNull(TRIBE_MOVEMENT_TABLE![Follow_Tribe]) Then
      Follow_Tribe = TRIBE_MOVEMENT_TABLE![Follow_Tribe]
   End If
   If TRIBE_MOVEMENT_TABLE![MOVEMENT_1] = "FOLLOW" Then
      TRIBE_MOVEMENT_TABLE.index = "SECONDARYKEY"
      TRIBE_MOVEMENT_TABLE.MoveFirst
      TRIBE_MOVEMENT_TABLE.Seek "=", Follow_Tribe
      If TRIBE_MOVEMENT_TABLE.NoMatch Then
         TRIBE_MOVEMENT_TABLE.index = "PRIMARYKEY"
         TRIBE_MOVEMENT_TABLE.Seek "=", MOVE_TRIBE
      ElseIf Multiple_Follows Then
         TRIBE_MOVEMENT_TABLE.index = "PRIMARYKEY"
         TRIBE_MOVEMENT_TABLE.Seek "=", MOVE_TRIBE
      Else
         Multiple_Follows = True
         TRIBE_MOVEMENT_TABLE.index = "PRIMARYKEY"
         TRIBE_MOVEMENT_TABLE.Seek "=", MOVE_TRIBE
         GoTo End_Loop
      End If
      CLAN = "0" & Mid(TRIBE_MOVEMENT_TABLE![Follow_Tribe], 2, 3)
      TRIBE = TRIBE_MOVEMENT_TABLE![Follow_Tribe]
      If Pass_One Then
         GoTo End_Loop
      End If
   End If
   If IsNull(TRIBE_MOVEMENT_TABLE![MOVEMENT_1]) Then
      If IsNull(TRIBE_MOVEMENT_TABLE![Follow_Tribe]) Then
         MOVEMENT_LINE = "No movement details provided"
         Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, "TRIBE MOVEMENT", 1, MOVEMENT_LINE, "No")
         TRIBE_MOVEMENT_TABLE.Edit
         TRIBE_MOVEMENT_TABLE![PROCESSED] = "Y"
         TRIBE_MOVEMENT_TABLE.UPDATE
         GoTo End_Loop
      Else
         TRIBE_MOVEMENT_TABLE.Edit
         TRIBE_MOVEMENT_TABLE![MOVEMENT_1] = "FOLLOW"
         TRIBE_MOVEMENT_TABLE.UPDATE
         CLAN = "0" & Mid(TRIBE_MOVEMENT_TABLE![Follow_Tribe], 2, 3)
         TRIBE = TRIBE_MOVEMENT_TABLE![Follow_Tribe]
         TRIBESINFO.MoveFirst
         TRIBESINFO.Seek "=", CLAN, TRIBE
         If TRIBESINFO.NoMatch Then
            MOVEMENT_LINE = "Incorrect movement details provided"
            Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, "TRIBE MOVEMENT", 1, MOVEMENT_LINE, "No")
            GoTo End_Loop
         End If
         If Pass_One Then
            GoTo End_Loop
         End If
      End If
   End If
   If TRIBE_MOVEMENT_TABLE![MOVEMENT_1] = "EMPTY" And IsNull(TRIBE_MOVEMENT_TABLE![Follow_Tribe]) Then
      MOVEMENT_LINE = "No movement details provided"
      Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, "TRIBE MOVEMENT", 1, MOVEMENT_LINE, "No")
      TRIBE_MOVEMENT_TABLE.Edit
      TRIBE_MOVEMENT_TABLE![PROCESSED] = "Y"
      TRIBE_MOVEMENT_TABLE.UPDATE
      GoTo End_Loop
   End If
   If TRIBE_MOVEMENT_TABLE![MOVEMENT_1] = "GOTO" And IsNull(TRIBE_MOVEMENT_TABLE![HEX]) Then
      MOVEMENT_LINE = "No movement details provided"
      Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, "TRIBE MOVEMENT", 1, MOVEMENT_LINE, "No")
      TRIBE_MOVEMENT_TABLE.Edit
      TRIBE_MOVEMENT_TABLE![PROCESSED] = "Y"
      TRIBE_MOVEMENT_TABLE.UPDATE
      GoTo End_Loop
   End If
      
End If

PRIMARY_TRIBE = Left(MOVE_TRIBE, 4)

SKILL_MOVE_TRIBE = Unit_Check("TRIBE", MOVE_TRIBE)
DICE_TRIBE = Unit_Check("DICE", MOVE_TRIBE)

Scouting = "NO"
PREVIOUS_Direction = "NONE"

DoCmd.Hourglass True

' TRIBE MOVEMENT

   
TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE
If TRIBESINFO.NoMatch Then
   ' not a valid tribe therefore the loop if screen
   If USE_SCREEN = "N" Then
      TRIBE_MOVEMENT_TABLE.Edit
      TRIBE_MOVEMENT_TABLE![PROCESSED] = "Y"
      TRIBE_MOVEMENT_TABLE.UPDATE
      GoTo End_Loop
   Else
   ' send message, exit function
   Msg = "Clan " & MOVE_CLAN & " Tribe " & MOVE_TRIBE
   Msg = Msg & " do not exist, exit movement"
   MsgBox (Msg)
   Resume TRIBE_MOVEMENT_EXIT_CLOSE
   End If
End If

TRIBESINFO.Edit

TRIBESINFO![Previous_Hex] = TRIBESINFO![CURRENT HEX]
TRIBESINFO.UPDATE
TRIBESINFO.Edit

If IsNull(TRIBESINFO![TRUCES]) Then
   Truced_Clans = "Empty"
Else
   Truced_Clans = TRIBESINFO![TRUCES]
End If

Total_People = TRIBESINFO!WARRIORS + TRIBESINFO!ACTIVES + TRIBESINFO!INACTIVES + TRIBESINFO!SLAVE
TCLANNUMBER = MOVE_CLAN
TM_POS = "GET WEIGHT"

TRIBES_WEIGHT = TRIBESINFO![WEIGHT]
TM_POS = "GET CAPACITY"
TRIBES_CAPACITY = TRIBESINFO![CAPACITY]
Walking_Capacity = TRIBESINFO![Walking_Capacity]

TM_POS = "Movement Info"

If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
   GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
Else
   GOODS_TRIBE = MOVE_TRIBE
End If

Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_GOODS")
TRIBESGOODS.index = "PRIMARYKEY"

Call Obtain_Skill_Levels

Set SHIPSTABLE = TVDB.OpenRecordset("VALID_SHIPS")
SHIPSTABLE.index = "PRIMARYKEY"

CURRENT_MAP = TRIBESINFO![CURRENT HEX]

Set Globaltable = TVDBGM.OpenRecordset("GLOBAL")
Globaltable.index = "PRIMARYKEY"
Globaltable.MoveFirst

hexmaptable.Seek "=", CURRENT_MAP

' DETERMINE WEATHER COST

If hexmaptable![WEATHER_ZONE] = "GREEN" Then
   WEATHER = Globaltable![Zone1]
   wind = Globaltable![Wind1]
   WIND_DIRECTION = Globaltable![DIRECTION1]
ElseIf hexmaptable![WEATHER_ZONE] = "RED" Then
   WEATHER = Globaltable![Zone2]
   wind = Globaltable![Wind2]
   WIND_DIRECTION = Globaltable![DIRECTION2]
ElseIf hexmaptable![WEATHER_ZONE] = "ORANGE" Then
   WEATHER = Globaltable![Zone3]
   wind = Globaltable![Wind3]
   WIND_DIRECTION = Globaltable![DIRECTION3]
ElseIf hexmaptable![WEATHER_ZONE] = "YELLOW" Then
   WEATHER = Globaltable![Zone4]
   wind = Globaltable![Wind4]
   WIND_DIRECTION = Globaltable![DIRECTION4]
ElseIf hexmaptable![WEATHER_ZONE] = "BLUE" Then
   WEATHER = Globaltable![Zone5]
   wind = Globaltable![Wind5]
   WIND_DIRECTION = Globaltable![DIRECTION5]
ElseIf hexmaptable![WEATHER_ZONE] = "BROWN" Then
   WEATHER = Globaltable![Zone6]
   wind = Globaltable![Wind6]
   WIND_DIRECTION = Globaltable![DIRECTION6]
End If
   
If codetrack = 1 Then
   MSG1 = "WEATHER = " & WEATHER & crlf
   MSG2 = "wind = " & wind & crlf
   MSG3 = "WIND_DIRECTION = " & WIND_DIRECTION & crlf
   Response = MsgBox((MSG1 & MSG2 & MSG3), True)
End If
   
Globaltable.Close

TM_POS = "First Movement Trace call"

Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, "TRIBE", 1

If Movement_Trace.NoMatch Then
   Movement_Trace.AddNew
   Movement_Trace![TRIBE] = MOVE_TRIBE
   Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
   Movement_Trace![Tribe_Or_Scout] = "TRIBE"
   Movement_Trace![Movement_Number] = 1
   Movement_Trace![STARTING_HEX] = CURRENT_MAP
   Movement_Trace![Target_Hex] = "??"
   Movement_Trace![Target_Terrain] = "??"
   Movement_Trace![Direction] = "??"
   Movement_Trace![WEATHER] = WEATHER
   Movement_Trace![wind] = wind
   Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
   Movement_Trace![Current_Movement_Points] = 0
   Movement_Trace![Current_Movement_Cost] = 0
   Movement_Trace![CAN_GROUP_MOVE] = "U"
   Movement_Trace![NO_MOVEMENT_REASON] = "??"
   Movement_Trace![SCOUTS] = 0
   Movement_Trace![HORSES] = 0
   Movement_Trace![Elephants] = 0
   Movement_Trace.UPDATE
   Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, "TRIBE", 1
Else
   ' can only happen on a repeat move or if the player backtracks
   Movement_Trace.Delete
   Movement_Trace.AddNew
   Movement_Trace![TRIBE] = MOVE_TRIBE
   Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
   Movement_Trace![Tribe_Or_Scout] = "TRIBE"
   Movement_Trace![Movement_Number] = 1
   Movement_Trace![STARTING_HEX] = CURRENT_MAP
   Movement_Trace![Target_Hex] = "??"
   Movement_Trace![Target_Terrain] = "??"
   Movement_Trace![Direction] = "??"
   Movement_Trace![WEATHER] = WEATHER
   Movement_Trace![wind] = wind
   Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
   Movement_Trace![Current_Movement_Points] = 0
   Movement_Trace![Current_Movement_Cost] = 0
   Movement_Trace![CAN_GROUP_MOVE] = "U"
   Movement_Trace![NO_MOVEMENT_REASON] = "??"
   Movement_Trace![SCOUTS] = 0
   Movement_Trace![HORSES] = 0
   Movement_Trace![Elephants] = 0
   Movement_Trace.UPDATE
   Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, "TRIBE", 1
End If
   
Set HEXMAPCITY = TVDBGM.OpenRecordset("HEX_MAP_CITY")
HEXMAPCITY.index = "PRIMARYKEY"

Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
HEXMAPMINERALS.index = "PRIMARYKEY"

Set HEXMAPPOLITICS = TVDBGM.OpenRecordset("HEX_MAP_politics")
HEXMAPPOLITICS.index = "PRIMARYKEY"

count = 1

TM_POS = "Hex_Map Construction update"
Do
 If Not HEXMAPCONST.EOF Then
    HEXMAPCONST.MoveFirst
 End If
 HEXMAPCONST.index = "CLANCONST"
 HEXMAPCONST.Seek "=", CURRENT_MAP, MOVE_CLAN, MOVE_TRIBE
 
 If HEXMAPCONST.NoMatch Then
     Exit Do
 Else
    TEMP_TRIBE = "AAA" & count
    HEXMAPCONST.Seek "=", CURRENT_MAP, TEMP_TRIBE, TEMP_TRIBE
    If HEXMAPCONST.NoMatch Then
       HEXMAPCONST.Seek "=", CURRENT_MAP, MOVE_CLAN, MOVE_TRIBE
       HEXMAPCONST.Edit
       HEXMAPCONST![CLAN] = TEMP_TRIBE
       HEXMAPCONST![TRIBE] = TEMP_TRIBE
       HEXMAPCONST.UPDATE
    Else
       count = count + 1
    End If
 End If

Loop

count = 1

TM_POS = "Clean Construction Table"
Do
 If Not ConstructionTable.EOF Then
    ConstructionTable.MoveFirst
 End If
 ConstructionTable.Seek "=", MOVE_TRIBE
 
 If ConstructionTable.NoMatch Then
     Exit Do
 Else
     ConstructionTable.Delete
 End If

Loop

LineI = 1

HORSES = 0
Elephants = 0

' DETERMINE IF A FLEET
  If TRIBESINFO![Village] = "FLEET" Then
     FLEET = "Y"
  ElseIf InStr(MOVE_TRIBE, "F") > 0 Then
     FLEET = "Y"
     TRIBESINFO.Edit
     TRIBESINFO![Village] = "FLEET"
     TRIBESINFO.UPDATE
  Else
     FLEET = "N"
  End If

TM_POS = "Check for spy glasses"

' DETERMINE IF HAVE SPYGLASSES
  Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_GOODS")
  TRIBESGOODS.MoveFirst
  TRIBESGOODS.index = "PRIMARYKEY"
  TRIBESGOODS.Seek "=", MOVE_CLAN, MOVE_TRIBE, "FINISHED", "SPYGLASS"
  If TRIBESGOODS.NoMatch Then
     SPYGLASSES = "N"
  ElseIf TRIBESGOODS![ITEM_NUMBER] > 0 Then
        SPYGLASSES = "Y"
  Else
     SPYGLASSES = "N"
  End If

TM_POS = "Get Movement points"

' GET MOVEMENT POINTS
  Call DETERMINE_MOVEMENT_POINTS("N")

  Movement_Trace.Edit
  Movement_Trace![Current_Movement_Points] = MOVEMENT_POINTS
  Movement_Trace.UPDATE

If codetrack = 1 Then
   MSG0 = "TRIBE MOVEMENT " & crlf
   MSG1 = "FLEET = " & FLEET & crlf
   MSG2 = "SPYGLASSES = " & SPYGLASSES & crlf
   MSG3 = "MOVEMENT_POINTS = " & MOVEMENT_POINTS & crlf
   Response = MsgBox((MSG0 & MSG1 & MSG2 & MSG3), True)
End If

TM_POS = "About to start movement"

If MOVEMENT_ORDERS(1) = "EMPTY" Then
   MOVEMENT_LINE = "EMPTY"
   Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, "TRIBE MOVEMENT", LineI, MOVEMENT_LINE, "No")
Else
   MOVEMENT_LINE = "Tribe Movement: Move "
   NEW_ORDERS = ""
   MOVEMENT_COUNT = 1
   MOVEMENT_ITERATIONS = 0
   Do Until MOVEMENT_COUNT >= 35
      TM_POS = "Movement Section"
      MOVEMENT_ITERATIONS = MOVEMENT_ITERATIONS + 1
      If MOVEMENT_ITERATIONS > 35 Then
         Exit Do
      End If
      Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, "TRIBE", MOVEMENT_ITERATIONS

      If Movement_Trace.NoMatch Then
         Movement_Trace.AddNew
         Movement_Trace![TRIBE] = MOVE_TRIBE
         Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
         Movement_Trace![Tribe_Or_Scout] = "TRIBE"
         Movement_Trace![Movement_Number] = MOVEMENT_ITERATIONS
         Movement_Trace![STARTING_HEX] = CURRENT_MAP
         Movement_Trace![Target_Hex] = "??"
         Movement_Trace![Target_Terrain] = "??"
         Movement_Trace![Direction] = Direction
         Movement_Trace![WEATHER] = WEATHER
         Movement_Trace![wind] = wind
         Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
         Movement_Trace![Current_Movement_Points] = MOVEMENT_POINTS
         Movement_Trace![Current_Movement_Cost] = 0
         Movement_Trace![CAN_GROUP_MOVE] = "U"
         Movement_Trace![NO_MOVEMENT_REASON] = "??"
         Movement_Trace![SCOUTS] = SCOUTS_USED
         Movement_Trace![HORSES] = 0
         Movement_Trace![Elephants] = 0
         Movement_Trace.UPDATE
         Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, "TRIBE", MOVEMENT_ITERATIONS
      Else
         ' can only happen on a repeat move or if the player backtracks
         Movement_Trace.Delete
         Movement_Trace.AddNew
         Movement_Trace![TRIBE] = MOVE_TRIBE
         Movement_Trace![GOODS_TRIBE] = GOODS_TRIBE
         Movement_Trace![Tribe_Or_Scout] = "TRIBE"
         Movement_Trace![Movement_Number] = MOVEMENT_ITERATIONS
         Movement_Trace![STARTING_HEX] = CURRENT_MAP
         Movement_Trace![Target_Hex] = "??"
         Movement_Trace![Target_Terrain] = "??"
         Movement_Trace![Direction] = Direction
         Movement_Trace![WEATHER] = WEATHER
         Movement_Trace![wind] = wind
         Movement_Trace![WIND_DIRECTION] = WIND_DIRECTION
         Movement_Trace![Current_Movement_Points] = MOVEMENT_POINTS
         Movement_Trace![Current_Movement_Cost] = 0
         Movement_Trace![CAN_GROUP_MOVE] = "U"
         Movement_Trace![NO_MOVEMENT_REASON] = "??"
         Movement_Trace![SCOUTS] = SCOUTS_USED
         Movement_Trace![HORSES] = 0
         Movement_Trace![Elephants] = 0
         Movement_Trace.UPDATE
         Movement_Trace.Seek "=", MOVE_TRIBE, GOODS_TRIBE, "TRIBE", MOVEMENT_ITERATIONS
      End If
      If Not IsNull(NEW_ORDERS) And Not (NEW_ORDERS = "") Then
         Direction = NEW_ORDERS
         ORIG_Direction = NEW_ORDERS
         NEW_ORDERS = ""
      Else
         Call GET_NEXT_TRIBE_MOVE
      End If
      If Direction = "NL" Then
         Direction = "N"
      ElseIf Direction = "NEL" Then
         Direction = "NE"
      ElseIf Direction = "SEL" Then
         Direction = "SE"
      ElseIf Direction = "SL" Then
         Direction = "S"
      ElseIf Direction = "SWL" Then
         Direction = "SW"
      ElseIf Direction = "NWL" Then
         Direction = "NW"
      End If
      If Direction = "FOLLOW" Then
         TRIBESINFO.MoveFirst
         TRIBESINFO.Seek "=", CLAN, TRIBE
         If TRIBESINFO.NoMatch Then
            GoTo End_Loop
         End If
         CURRENT_MAP = TRIBESINFO![CURRENT HEX]
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", CURRENT_MAP
         TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE
         TRIBESINFO.Edit
         TRIBESINFO![CURRENT HEX] = CURRENT_MAP
         TRIBESINFO![CURRENT TERRAIN] = hexmaptable![TERRAIN]
         TRIBESINFO.UPDATE
         Call Tribe_Checking("Update_Hex", MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP)
         MOVEMENT_LINE = "Tribe Follows " & TRIBE
         MOVEMENT_POINTS = 0
         GROUP_MOVE = "N"
         Exit Do
      ElseIf Direction = "GOTO" Then
         CURRENT_MAP = Right(TRIBE_MOVEMENT_TABLE![HEX], 7)
         hexmaptable.MoveFirst
         hexmaptable.Seek "=", CURRENT_MAP
         TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE
         If TRIBESINFO.NoMatch Then
            GoTo End_Loop
         End If
         TRIBESINFO.Edit
         TRIBESINFO![CURRENT HEX] = CURRENT_MAP
         TRIBESINFO![CURRENT TERRAIN] = hexmaptable![TERRAIN]
         TRIBESINFO.UPDATE
         Call Tribe_Checking("Update_Hex", MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP)
         MOVEMENT_LINE = "Tribe Goes to " & CURRENT_MAP
         MOVEMENT_POINTS = 0
         GROUP_MOVE = "N"
         Exit Do
      ElseIf Direction = "STILL" Then
         GROUP_MOVE = "Y"
      ElseIf Direction = "STOP" Then
         Exit Do
      ElseIf Direction = "EMPTY" Then
         Exit Do
      ElseIf Direction = "HALT" Then
         Exit Do
      Else
         If FLEET = "Y" Then
             If TRIBES_WEIGHT > TRIBES_CAPACITY Then
                  NO_MOVEMENT_REASON = "Insufficient capacity to carry "
                  GROUP_MOVE = "N"
             ElseIf TRIBES_WEIGHT > TRIBES_CAPACITY Then
                    If TRIBES_WEIGHT < Walking_Capacity Then
                        GROUP_MOVE = "Y"
                    Else
                          NO_MOVEMENT_REASON = "Insufficient capacity to carry "
                          GROUP_MOVE = "N"
                    End If
             Else
                    GROUP_MOVE = "Y"
             End If
             If GROUP_MOVE = "Y" Then
                 Call CAN_FLEET_MOVE(CURRENT_MAP, "N")
             End If
         Else
             If TRIBES_WEIGHT > Walking_Capacity Then
                  NO_MOVEMENT_REASON = "Insufficient capacity to carry "
                  GROUP_MOVE = "N"
             ElseIf NO_MOVEMENT_REASON = "Not enough animals to pull wagons" Then
                    GROUP_MOVE = "N"
                    MOVEMENT_LINE = "Tribe Movement: Not enough animals to pull wagons. Movement is not possible."
                    Movement_Trace.Edit
                    Movement_Trace![NO_MOVEMENT_REASON] = NO_MOVEMENT_REASON
                    Movement_Trace.UPDATE
                    Msg = "The Clan was " & MOVE_CLAN & " The Tribe was " & MOVE_TRIBE
                    Msg = Msg & Chr(13) & Chr(10) & " Not enough animals to pull wagons."
                    MsgBox (Msg)
                    Exit Do
             ElseIf TRIBES_WEIGHT > TRIBES_CAPACITY Then
                    If TRIBES_WEIGHT < Walking_Capacity Then
                        GROUP_MOVE = "Y"
                    Else
                          NO_MOVEMENT_REASON = "Insufficient capacity to carry "
                          GROUP_MOVE = "N"
                    End If
             Else
                    GROUP_MOVE = "Y"
             End If
             If GROUP_MOVE = "Y" Then
                 Call CAN_GROUP_MOVE(CURRENT_MAP, "N")
             End If
         End If
      End If
      If codetrack = 1 Then
         MSG1 = "GROUP_MOVE = " & GROUP_MOVE & crlf
         MSG2 = "DIRECTION = " & Direction & crlf
         Response = MsgBox((MSG1 & MSG2), True)
      End If
      If GROUP_MOVE = "Y" Then
         If Not Direction = "STILL" Then
            Call GET_TERRAIN(Direction, TERRAIN, CURRENT_MAP)
         End If
         If Len(TERRAIN) > 0 Then
            MOVEMENT_LINE = MOVEMENT_LINE & "^B" & Direction & "-" & TERRAIN & "^B"
            PREVIOUS_Direction = Direction
            TERRAIN = ""
            Direction = ""
         End If
         If Direction = "STILL" Then
            If SPYGLASSES = "Y" Then
               GET_SURROUNDING_TERRAIN (CURRENT_MAP)
               hexmaptable.MoveFirst
               hexmaptable.Seek "=", CURRENT_MAP
            End If
         ElseIf (Right(CURRENT_TERRAIN, 9) = "MOUNTAINS") Or (Right(CURRENT_TERRAIN, 2) = "MT") Then
            If SPYGLASSES = "Y" Then
               GET_SURROUNDING_TERRAIN (CURRENT_MAP)
               hexmaptable.MoveFirst
               hexmaptable.Seek "=", CURRENT_MAP
            End If
         ElseIf FLEET = "Y" Then
            GET_SURROUNDING_FLEET (CURRENT_MAP)
            hexmaptable.MoveFirst
            hexmaptable.Seek "=", CURRENT_MAP
         End If
         If codetrack = 1 Then
            MSG0 = "TRIBE MOVEMENT " & crlf
            MSG1 = "TERRAIN = " & TERRAIN & crlf
            MSG2 = "DIRECTION = " & Direction & crlf
            Response = MsgBox((MSG0 & MSG1 & MSG2), True)
         End If
         If Len(TERRAIN) > 0 Then
            ' TRIBEINHEX TO BE POPULATED ONCE THIS IS BATCHED.
            TRIBESINHEX = "EMPTY"
            ' TRIBESINHEX = WHO_IS_IN_HEX(MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP, "N")
            If Not TRIBESINHEX = "EMPTY" Then
               MOVEMENT_LINE = MOVEMENT_LINE & Direction & "-" & TERRAIN & ", " & TRIBESINHEX & "\"
               MSG0 = "Tribe Finds " & TRIBESINHEX & crlf
               Response = MsgBox(MSG0, True)
            Else
               MOVEMENT_LINE = MOVEMENT_LINE & Direction & "-" & TERRAIN & "\"
            End If
         Else
            MOVEMENT_LINE = MOVEMENT_LINE & "\"
         End If
         If MOVEMENT_POINTS < 3 Then
            TM_POS = "Movement Points less than 3"
            Movement_Trace.Edit
            Movement_Trace![NO_MOVEMENT_REASON] = "3 or less MP's"
            Movement_Trace.UPDATE
            Exit Do
         End If
         TERRAIN = ""
         Direction = ""
         If codetrack = 1 Then
            MSG0 = "TRIBE MOVEMENT " & crlf
            MSG1 = "MOVEMENT_LINE = " & MOVEMENT_LINE & crlf
            Response = MsgBox((MSG0 & MSG1), True)
         End If
      ElseIf MOVEMENT_POINTS > 3 Then
         SPOSITION = 1
         NEW_ORDERS = "STOP"
         If NEW_ORDERS = "STOP" Then
            If NO_MOVEMENT_REASON = "Insufficient capacity to carry " Then
                MOVEMENT_LINE = "Tribe Movement: Move failed due to " & NO_MOVEMENT_REASON
            Else
                MOVEMENT_LINE = MOVEMENT_LINE & NO_MOVEMENT_REASON
            End If
            TM_POS = "Alternative movement required"
            Movement_Trace.Edit
            Movement_Trace![NO_MOVEMENT_REASON] = NO_MOVEMENT_REASON
            Movement_Trace.UPDATE
            Exit Do
         Else
            PREVIOUS_Direction = Direction
            Direction = NEW_ORDERS
            ORIG_Direction = NEW_ORDERS
         End If
      Else
         TM_POS = "Cant Move"
         Movement_Trace.Edit
         Movement_Trace![NO_MOVEMENT_REASON] = "Group cant move and MP's are 3 or less"
         Movement_Trace.UPDATE
         Exit Do
      End If
      'MSG = "TERRAIN = " & TERRAIN
      'RESPONSE = MsgBox(MSG, True)
      If ORIG_Direction = "FRR" Or ORIG_Direction = "FRL" _
      Or ORIG_Direction = "FMR" Or ORIG_Direction = "FML" _
      Or ORIG_Direction = "FLR" Or ORIG_Direction = "FLL" _
      Or ORIG_Direction = "FOR" Or ORIG_Direction = "FOL" _
      Or ORIG_Direction = "FCR" Or ORIG_Direction = "FCL" _
      Or ORIG_Direction = "FL" Or ORIG_Direction = "FO" _
      Or ORIG_Direction = "NL" Or ORIG_Direction = "NEL" _
      Or ORIG_Direction = "SEL" Or ORIG_Direction = "SL" _
      Or ORIG_Direction = "SWL" Or ORIG_Direction = "NWL" Then
         'do nothing
      ElseIf ORIG_Direction = "STILL" Then
         Exit Do
      Else
         MOVEMENT_COUNT = MOVEMENT_COUNT + 1
      End If
   Loop
   TM_POS = "End of Movement Section"

   TRIBESINFO.Edit
   TRIBESINFO![CURRENT HEX] = CURRENT_MAP
   TRIBESINFO.UPDATE
   
   Call Tribe_Checking("Update_Hex", MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP)

End If

TM_POS = "Politics Section"

hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP
  
Set HEXMAPPOLITICS = TVDBGM.OpenRecordset("HEX_MAP_politics")
HEXMAPPOLITICS.index = "PRIMARYKEY"
HEXMAPPOLITICS.MoveFirst
HEXMAPPOLITICS.Seek "=", CURRENT_MAP

If HEXMAPPOLITICS.NoMatch Then
   ' lets not fuss
ElseIf HEXMAPPOLITICS![PL_CLAN] = MOVE_CLAN Then
   If HEXMAPPOLITICS![PL_TRIBE] = MOVE_TRIBE Then
      HEXMAPPOLITICS.Edit
      HEXMAPPOLITICS![PL_CLAN] = "MOVE_CLAN"
      HEXMAPPOLITICS![PL_TRIBE] = "MOVE_TRIBE"
      HEXMAPPOLITICS![PACIFICATION_LEVEL] = 0
      HEXMAPPOLITICS![POPULATION] = 0
      HEXMAPPOLITICS.UPDATE
   End If
End If
   
  Set TERRAINTABLE = TVDB.OpenRecordset("VALID_TERRAIN")
  TERRAINTABLE.index = "PRIMARYKEY"
  TERRAINTABLE.MoveFirst
  TERRAINTABLE.Seek "=", hexmaptable![TERRAIN]

  TRIBESINFO.Edit
  TRIBESINFO![CURRENT HEX] = CURRENT_MAP
  TRIBESINFO![CURRENT TERRAIN] = hexmaptable![TERRAIN]
  TRIBESINFO.UPDATE
  Call Tribe_Checking("Update_Hex", MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP)

  TM_POS = "Write out Movement"
  
  If Len(MOVEMENT_LINE) > 0 Then
     Call WRITE_TURN_ACTIVITY(MOVE_CLAN, MOVE_TRIBE, "TRIBE MOVEMENT", LineI, MOVEMENT_LINE, "No")
  End If

count = 1
TEMP_TRIBE = "AAA" & count

TM_POS = "Update Hexmapconstruction"
HEXMAPCONST.index = "MAP"
HEXMAPCONST.MoveFirst
HEXMAPCONST.Seek "=", CURRENT_MAP
   
If HEXMAPCONST.NoMatch Then
   STOP_LOOP = "YES"
Else
   ' loop through hex taking ownership of vacant buildings
   Do While HEXMAPCONST![MAP] = CURRENT_MAP
      If Left(HEXMAPCONST![CLAN], 3) = "AAA" Then
         HEXMAPCONST.Edit
         HEXMAPCONST![CLAN] = MOVE_CLAN
         HEXMAPCONST![TRIBE] = MOVE_TRIBE
         HEXMAPCONST.UPDATE
         If Err = 3022 Then
            
         End If
      End If
      HEXMAPCONST.MoveNext
      If HEXMAPCONST.EOF Then
         Exit Do
      End If
   Loop
End If
TM_POS = "Update Movement Table"

'update scout as processed
TRIBE_MOVEMENT_TABLE.Edit
TRIBE_MOVEMENT_TABLE![PROCESSED] = "Y"
TRIBE_MOVEMENT_TABLE.UPDATE

If SCREEN = "Y" Then
   Exit Do
End If

End_Loop:
TRIBE_MOVEMENT_TABLE.MoveNext


If TRIBE_MOVEMENT_TABLE![TRIBE] = "9999" Then
   Exit Do
End If

Loop

' checks if its the first pass, if true then reset to start of table
' processing all follow movement orders as pass_two
If Pass_One Then
   Pass_One = False
   TRIBE_MOVEMENT_TABLE.MoveFirst
   GoTo Movement_Table_Start
End If

' checks if its the second pass, and multiple follows were found
If Not Pass_One Then
   If Multiple_Follows Then
      Multiple_Follows = False
      TRIBE_MOVEMENT_TABLE.MoveFirst
      GoTo Movement_Table_Start
   End If
End If

TM_POS = "Movement Loop Ended"
HEXMAPCONST.Close
hexmaptable.Close
TERRAINTABLE.Close
Movement_Trace.Close

If MOVE_TRIBE = "" Then
   'ignore
Else
   'DUMP GOODS WHEN HAVE MOVED.
   Num_Goods = GET_TRIBES_GOOD_QUANTITY(MOVE_CLAN, GOODS_TRIBE, "FISHTRAP")
   If Num_Goods > 0 Then
      Call UPDATE_TRIBES_GOODS_TABLES(MOVE_CLAN, GOODS_TRIBE, "FISHTRAP", "SUBTRACT", Num_Goods)
   End If
    
   Call Tribe_Checking("Update_Hex", MOVE_CLAN, MOVE_TRIBE, CURRENT_MAP)
   
   Msg = "FINAL HEX = " & CURRENT_MAP
   Response = MsgBox(Msg, 0)

   DoCmd.Close A_FORM, "TRIBE MOVEMENT"
   DoCmd.OpenForm "TRIBE MOVEMENT"
End If

TRIBE_MOVEMENT_EXIT_CLOSE:
   DoCmd.Hourglass False
   Exit Function


TRIBE_MOVEMENT_EXIT:
If (Err = 3021) Or (Err = 3022) Or Err = 6 Then
   Resume Next

ElseIf Err = 94 Then
   If TM_POS = "GET WEIGHT" Then
      Call Determine_Weights(MOVE_CLAN, MOVE_TRIBE)
      TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE
      TRIBESINFO.Edit
      TRIBES_WEIGHT = TRIBESINFO![WEIGHT]
      Resume Next
   ElseIf TM_POS = "GET CAPACITY" Then
      Call Determine_Capacities("GROUP", MOVE_CLAN, MOVE_TRIBE)
      TRIBESINFO.Seek "=", MOVE_CLAN, MOVE_TRIBE
      TRIBESINFO.Edit
      TRIBES_CAPACITY = TRIBESINFO![CAPACITY]
      Walking_Capacity = TRIBESINFO![Walking_Capacity]
      Resume Next
   End If
ElseIf Err = 3163 Then
   'FIELD TO SHORT
   ' WHICH BLOODY FIELD
   Msg = "Error # " & Err & " " & Error$
   Msg = Msg & " Error Occurred in section" & TM_POS
   MsgBox (Msg)
   Resume TRIBE_MOVEMENT_EXIT_CLOSE
   
Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Msg = "Error Occurred in section " & TM_POS
   MsgBox (Msg)
   Resume TRIBE_MOVEMENT_EXIT_CLOSE
End If

End Function

Sub UPDATE_HEX_MAP(CURRENT_MAP)
Dim FORMARG As String

If codetrack = 1 Then
   MSG1 = "SUB FUNCTION = ADD NEW HEX" & crlf
   Response = MsgBox((MSG1), True)
End If

FORMARG = "[MAP] = """ & CURRENT_MAP & """"

DoCmd.Hourglass False

DoCmd.OpenForm "HEX_MAP", , , FORMARG, A_EDIT, A_DIALOG

DoCmd.Hourglass True


End Sub

Sub UPDATE_TRIBES_TABLES(ITEM, MOVE_TYPE, MOVE_QUANTITY)
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
On Error GoTo ERR_TABLES

VALID_GOODS:
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "PRIMARYKEY"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", ITEM
       
If VALIDGOODS.NoMatch Then
   Msg = "ITEM NOT FOUND = " & ITEM
   Response = MsgBox(Msg, True)
   Msg = "AMOUNT TO ADD = " & MOVE_QUANTITY
   Response = MsgBox(Msg, True)
   Exit Sub
End If
   
Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, VALIDGOODS![TABLE], ITEM
If MOVE_TYPE = "ADD" Then
   If TRIBESGOODS.NoMatch Then
      TRIBESGOODS.AddNew
      TRIBESGOODS![CLAN] = TCLANNUMBER
      TRIBESGOODS![TRIBE] = GOODS_TRIBE
      TRIBESGOODS![ITEM_TYPE] = VALIDGOODS![TABLE]
      TRIBESGOODS![ITEM] = ITEM
      TRIBESGOODS![ITEM_NUMBER] = MOVE_QUANTITY
      TRIBESGOODS.UPDATE
   Else
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + MOVE_QUANTITY
      TRIBESGOODS.UPDATE
   End If
Else
   TRIBESGOODS.Edit
   TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - MOVE_QUANTITY
   TRIBESGOODS.UPDATE
   If TRIBESGOODS![ITEM_NUMBER] <= 0 Then
      TRIBESGOODS.Delete
   End If
End If
       
ERR_close:
   Exit Sub

ERR_TABLES:
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  Resume ERR_close


End Sub

Public Function Check_Surrounding_Terrain(TERRAIN, CURRENT_MAP, TERRAIN_TO_FIND, SHORT_TERRAIN)
Dim TERRAIN_FOUND As String

TERRAIN_FOUND = "NO"

If NE_TERRAIN = TERRAIN_TO_FIND Then
   TERRAIN = TERRAIN & ", " & SHORT_TERRAIN & " NE"
   TERRAIN_FOUND = "YES"
End If

If SE_TERRAIN = TERRAIN_TO_FIND Then
   If TERRAIN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & ", " & SHORT_TERRAIN & " SE"
      TERRAIN_FOUND = "YES"
   End If
End If

If SW_TERRAIN = TERRAIN_TO_FIND Then
   If TERRAIN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & ", " & SHORT_TERRAIN & " SW"
      TERRAIN_FOUND = "YES"
   End If
End If

If NW_TERRAIN = TERRAIN_TO_FIND Then
   If TERRAIN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & ", " & SHORT_TERRAIN & " NW"
      TERRAIN_FOUND = "YES"
   End If
End If

If N_TERRAIN = TERRAIN_TO_FIND Then
   If TERRAIN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & ", " & SHORT_TERRAIN & " N"
      TERRAIN_FOUND = "YES"
   End If
End If

If S_TERRAIN = TERRAIN_TO_FIND Then
   If TERRAIN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & ", " & SHORT_TERRAIN & " S"
      TERRAIN_FOUND = "YES"
   End If
End If

End Function
Function PROCESS_SCOUTING_LOSSES()
Dim QUANTITY As Long
Dim UPDATE As String

CLAN = Forms![SCOUTING_LOSSES]![CLANNUMBER]
TRIBE = Forms![SCOUTING_LOSSES]![TRIBENUMBER]
UPDATE = "SUBTRACT"

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 1]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 1]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 1]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 2]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 2]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 2]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 3]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 3]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 3]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 4]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 4]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 4]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 5]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 5]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 5]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 6]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 6]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 6]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 7]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 7]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 7]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 8]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 8]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 8]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 9]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 9]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 9]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 10]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 10]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 10]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 11]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 11]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 11]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 12]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 12]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 12]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 13]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 13]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 13]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

If Not IsNull(Forms![SCOUTING_LOSSES]![ITEM 14]) Then
   ITEM = Forms![SCOUTING_LOSSES]![ITEM 14]
   QUANTITY = Forms![SCOUTING_LOSSES]![AMT 14]
   Call UPDATE_TRIBES_GOODS_TABLES(CLAN, TRIBE, ITEM, UPDATE, QUANTITY)
End If

Forms![SCOUTING_LOSSES]![AMT 1] = ""
Forms![SCOUTING_LOSSES]![AMT 2] = ""
Forms![SCOUTING_LOSSES]![AMT 3] = ""
Forms![SCOUTING_LOSSES]![AMT 4] = ""
Forms![SCOUTING_LOSSES]![AMT 5] = ""
Forms![SCOUTING_LOSSES]![AMT 6] = ""
Forms![SCOUTING_LOSSES]![AMT 7] = ""
Forms![SCOUTING_LOSSES]![AMT 8] = ""
Forms![SCOUTING_LOSSES]![AMT 9] = ""
Forms![SCOUTING_LOSSES]![AMT 10] = ""
Forms![SCOUTING_LOSSES]![AMT 11] = ""
Forms![SCOUTING_LOSSES]![AMT 12] = ""
Forms![SCOUTING_LOSSES]![AMT 13] = ""
Forms![SCOUTING_LOSSES]![AMT 14] = ""
Forms![SCOUTING_LOSSES]![ITEM 1] = ""
Forms![SCOUTING_LOSSES]![ITEM 2] = ""
Forms![SCOUTING_LOSSES]![ITEM 3] = ""
Forms![SCOUTING_LOSSES]![ITEM 4] = ""
Forms![SCOUTING_LOSSES]![ITEM 5] = ""
Forms![SCOUTING_LOSSES]![ITEM 6] = ""
Forms![SCOUTING_LOSSES]![ITEM 7] = ""
Forms![SCOUTING_LOSSES]![ITEM 8] = ""
Forms![SCOUTING_LOSSES]![ITEM 9] = ""
Forms![SCOUTING_LOSSES]![ITEM 10] = ""
Forms![SCOUTING_LOSSES]![ITEM 11] = ""
Forms![SCOUTING_LOSSES]![ITEM 12] = ""
Forms![SCOUTING_LOSSES]![ITEM 13] = ""
Forms![SCOUTING_LOSSES]![ITEM 14] = ""

DoCmd.Hourglass False


End Function



Public Function GET_QUARRIES(TERRAIN, CURRENT_MAP)
Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

If hexmaptable![QUARRYING] = "Y" Then
   TERRAIN = TERRAIN & " Quarry Hex,"
End If

End Function

Public Function GET_SPRINGS(TERRAIN, CURRENT_MAP)
Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", CURRENT_MAP

   If hexmaptable![SPRINGS] = "Y" Then
      TERRAIN = TERRAIN & " Springs in hex,"
   End If


End Function

Function GET_OCEANS(TERRAIN, CURRENT_MAP)
Dim OCEAN_FOUND As String

OCEAN_FOUND = "NO"

If (NE_TERRAIN = "OCEAN") Then
   TERRAIN = TERRAIN & " O NE"
   OCEAN_FOUND = "YES"
End If

If SE_TERRAIN = "OCEAN" Then
   If OCEAN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " O SE"
      OCEAN_FOUND = "YES"
   End If
End If

If SW_TERRAIN = "OCEAN" Then
   If OCEAN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " O SW"
      OCEAN_FOUND = "YES"
   End If
End If

If NW_TERRAIN = "OCEAN" Then
   If OCEAN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " O NW"
      OCEAN_FOUND = "YES"
   End If
End If

If N_TERRAIN = "OCEAN" Then
   If OCEAN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " O N"
      OCEAN_FOUND = "YES"
   End If
End If

If S_TERRAIN = "OCEAN" Then
   If OCEAN_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " O S"
      OCEAN_FOUND = "YES"
   End If
End If

If OCEAN_FOUND = "YES" Then
   TERRAIN = TERRAIN & ","
End If

End Function

Public Function GET_LAKES(TERRAIN, CURRENT_MAP)
Dim LAKE_FOUND As String

LAKE_FOUND = "NO"

If (NE_TERRAIN = "LAKE") Then
   TERRAIN = TERRAIN & " L NE"
   LAKE_FOUND = "YES"
End If

If SE_TERRAIN = "LAKE" Then
   If LAKE_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SE"
   Else
      TERRAIN = TERRAIN & " L SE"
      LAKE_FOUND = "YES"
   End If
End If

If SW_TERRAIN = "LAKE" Then
   If LAKE_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", SW"
   Else
      TERRAIN = TERRAIN & " L SW"
      LAKE_FOUND = "YES"
   End If
End If

If NW_TERRAIN = "LAKE" Then
   If LAKE_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", NW"
   Else
      TERRAIN = TERRAIN & " L NW"
      LAKE_FOUND = "YES"
   End If
End If

If N_TERRAIN = "LAKE" Then
   If LAKE_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", N"
   Else
      TERRAIN = TERRAIN & " L N"
      LAKE_FOUND = "YES"
   End If
End If

If S_TERRAIN = "LAKE" Then
   If LAKE_FOUND = "YES" Then
      TERRAIN = TERRAIN & ", S"
   Else
      TERRAIN = TERRAIN & " L S"
      LAKE_FOUND = "YES"
   End If
End If

End Function


Public Function GET_NEXT_TRIBE_MOVE()
      
If USE_SCREEN = "Y" Then
   If MOVEMENT_COUNT < 10 Then
      stext = "Movement0" & CStr(MOVEMENT_COUNT)
   Else
      stext = "Movement" & CStr(MOVEMENT_COUNT)
   End If
   sValue = MOVEFORM(stext).Value
   
   If IsNull(sValue) Then
      Direction = "STOP"
      ORIG_Direction = "STOP"
   Else
      Direction = MOVEFORM(stext).Value
      ORIG_Direction = MOVEFORM(stext).Value
   End If
Else
   stext = "MOVEMENT_" & CStr(MOVEMENT_COUNT)
   If IsNull(TRIBE_MOVEMENT_TABLE(stext).Value) Then
      Direction = "EMPTY"
      ORIG_Direction = "EMPTY"
   Else
      Direction = TRIBE_MOVEMENT_TABLE(stext).Value
      ORIG_Direction = TRIBE_MOVEMENT_TABLE(stext).Value
   End If
   
End If

End Function


Public Function Get_HexMAP_and_Terrain_of_a_hex(STARTING_HEX, DIRECTION1, DIRECTION2, DIRECTION3, DIRECTION4, DIRECTION5, DIRECTION6, DIRECTION7, DIRECTION8)
On Error GoTo GET_HEXMAP_EXIT
Dim DIRECTION1_HEX As String
Dim DIRECTION2_HEX As String
Dim DIRECTION3_HEX As String
Dim DIRECTION4_HEX As String
Dim DIRECTION5_HEX As String
Dim DIRECTION6_HEX As String
Dim DIRECTION7_HEX As String
Dim DIRECTION8_HEX As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
         
hexmaptable.Seek "=", STARTING_HEX
CURRENT_TERRAIN = hexmaptable![TERRAIN]

If DIRECTION1 = "N" Then
    DIRECTION1_HEX = GET_MAP_NORTH(STARTING_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION1_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION1 = "NE" Then
    DIRECTION1_HEX = GET_MAP_NORTH_EAST(STARTING_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION1_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION1 = "SE" Then
    DIRECTION1_HEX = GET_MAP_SOUTH_EAST(STARTING_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION1_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION1 = "S" Then
    DIRECTION1_HEX = GET_MAP_SOUTH(STARTING_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION1_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION1 = "SW" Then
    DIRECTION1_HEX = GET_MAP_SOUTH_WEST(STARTING_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION1_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION1 = "NW" Then
    DIRECTION1_HEX = GET_MAP_NORTH_WEST(STARTING_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION1_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION2 = "N" Then
    DIRECTION2_HEX = GET_MAP_NORTH(DIRECTION1_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION2_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION2 = "NE" Then
    DIRECTION2_HEX = GET_MAP_NORTH_EAST(DIRECTION1_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION2_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION2 = "SE" Then
    DIRECTION2_HEX = GET_MAP_SOUTH_EAST(DIRECTION1_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION2_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION2 = "S" Then
    DIRECTION2_HEX = GET_MAP_SOUTH(DIRECTION1_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION2_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION2 = "SW" Then
    DIRECTION2_HEX = GET_MAP_SOUTH_WEST(DIRECTION1_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION2_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION2 = "NW" Then
    DIRECTION2_HEX = GET_MAP_NORTH_WEST(DIRECTION1_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION2_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION3 = "N" Then
    DIRECTION3_HEX = GET_MAP_NORTH(DIRECTION2_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION3_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION3 = "NE" Then
    DIRECTION3_HEX = GET_MAP_NORTH_EAST(DIRECTION2_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION3_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION3 = "SE" Then
    DIRECTION3_HEX = GET_MAP_SOUTH_EAST(DIRECTION2_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION3_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION3 = "S" Then
    DIRECTION3_HEX = GET_MAP_SOUTH(DIRECTION2_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION3_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION3 = "SW" Then
    DIRECTION3_HEX = GET_MAP_SOUTH_WEST(DIRECTION2_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION3_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION3 = "NW" Then
    DIRECTION3_HEX = GET_MAP_NORTH_WEST(DIRECTION2_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION3_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION4 = "N" Then
    DIRECTION4_HEX = GET_MAP_NORTH(DIRECTION3_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION4_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION4 = "NE" Then
    DIRECTION4_HEX = GET_MAP_NORTH_EAST(DIRECTION3_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION4_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION4 = "SE" Then
    DIRECTION4_HEX = GET_MAP_SOUTH_EAST(DIRECTION3_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION4_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION4 = "S" Then
    DIRECTION4_HEX = GET_MAP_SOUTH(DIRECTION3_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION4_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION4 = "SW" Then
    DIRECTION4_HEX = GET_MAP_SOUTH_WEST(DIRECTION3_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION4_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION4 = "NW" Then
    DIRECTION4_HEX = GET_MAP_NORTH_WEST(DIRECTION3_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION4_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION5 = "N" Then
    DIRECTION5_HEX = GET_MAP_NORTH(DIRECTION4_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION5_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION5 = "NE" Then
    DIRECTION5_HEX = GET_MAP_NORTH_EAST(DIRECTION4_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION5_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION5 = "SE" Then
    DIRECTION5_HEX = GET_MAP_SOUTH_EAST(DIRECTION4_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION5_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION5 = "S" Then
    DIRECTION5_HEX = GET_MAP_SOUTH(DIRECTION4_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION5_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION5 = "SW" Then
    DIRECTION5_HEX = GET_MAP_SOUTH_WEST(DIRECTION4_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION5_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION5 = "NW" Then
    DIRECTION5_HEX = GET_MAP_NORTH_WEST(DIRECTION4_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION5_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION6 = "N" Then
    DIRECTION6_HEX = GET_MAP_NORTH(DIRECTION5_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION6_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION6 = "NE" Then
    DIRECTION6_HEX = GET_MAP_NORTH_EAST(DIRECTION5_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION6_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION6 = "SE" Then
    DIRECTION6_HEX = GET_MAP_SOUTH_EAST(DIRECTION5_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION6_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION6 = "S" Then
    DIRECTION6_HEX = GET_MAP_SOUTH(DIRECTION5_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION6_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION6 = "SW" Then
    DIRECTION6_HEX = GET_MAP_SOUTH_WEST(DIRECTION5_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION6_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION6 = "NW" Then
    DIRECTION6_HEX = GET_MAP_NORTH_WEST(DIRECTION5_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION6_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION7 = "N" Then
    DIRECTION7_HEX = GET_MAP_NORTH(DIRECTION6_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION7_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION7 = "NE" Then
    DIRECTION7_HEX = GET_MAP_NORTH_EAST(DIRECTION6_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION7_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION7 = "SE" Then
    DIRECTION7_HEX = GET_MAP_SOUTH_EAST(DIRECTION6_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION7_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION7 = "S" Then
    DIRECTION7_HEX = GET_MAP_SOUTH(DIRECTION6_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION7_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION7 = "SW" Then
    DIRECTION7_HEX = GET_MAP_SOUTH_WEST(DIRECTION6_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION7_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION7 = "NW" Then
    DIRECTION7_HEX = GET_MAP_NORTH_WEST(DIRECTION6_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION7_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

If DIRECTION8 = "N" Then
    DIRECTION8_HEX = GET_MAP_NORTH(DIRECTION7_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION8_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION8 = "NE" Then
    DIRECTION8_HEX = GET_MAP_NORTH_EAST(DIRECTION7_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION8_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION8 = "SE" Then
    DIRECTION8_HEX = GET_MAP_SOUTH_EAST(DIRECTION7_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION8_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION8 = "S" Then
    DIRECTION8_HEX = GET_MAP_SOUTH(DIRECTION7_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION8_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION8 = "SW" Then
    DIRECTION8_HEX = GET_MAP_SOUTH_WEST(DIRECTION7_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION8_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
ElseIf DIRECTION8 = "NW" Then
    DIRECTION8_HEX = GET_MAP_NORTH_WEST(DIRECTION7_HEX)
    hexmaptable.MoveFirst
    hexmaptable.Seek "=", DIRECTION8_HEX
    CURRENT_TERRAIN = hexmaptable![TERRAIN]
End If

CURRENT_HEX = hexmaptable![MAP]

GET_HEXMAP_EXIT_CLOSE:
   DoCmd.Hourglass False
   Exit Function


GET_HEXMAP_EXIT:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Msg = "Error Occurred in section GET_HEXMAP and Terrain"
   MsgBox (Msg)
   Resume GET_HEXMAP_EXIT_CLOSE
End If

End Function
Public Function Populate_Tribe_Movement_Tsble()

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![TRIBE MOVEMENT]

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", MYFORM![MAP]

If hexmaptable.NoMatch Then
   MYFORM![North_Border] = "None"
   MYFORM![North_East_Border] = "None"
   MYFORM![South_East_Border] = "None"
   MYFORM![South_Border] = "None"
   MYFORM![South_West_Border] = "None"
   MYFORM![North_West_Border] = "None"
   MYFORM![ROAD(N)] = "N"
   MYFORM![ROAD(NE)] = "N"
   MYFORM![ROAD(SE)] = "N"
   MYFORM![ROAD(S)] = "N"
   MYFORM![ROAD(SW)] = "N"
   MYFORM![ROAD(NW)] = "N"

End If


End Function

Function Check_Truced(CLAN, TRUCED, FOUND)
Dim String_Length
Dim count As Long
Dim count_2 As Long
Dim start As Long
Dim Truced_Clan(30) As String
Dim FOUND_UNIT(30) As String
Dim Interim_Found As String
Dim New_Found As String
Dim FIND As Integer
Dim Found_Clan As String
Dim Search_Clan As String

If TRUCED = "EMPTY" Then
   GoTo EXIT_POINT
End If

Search_Clan = "0" & Mid(CLAN, 2, 3)
count = 1
Do
   Truced_Clan(count) = "EMPTY"
   FOUND_UNIT(count) = "EMPTY"
   count = count + 1
   If count > 30 Then
      Exit Do
   End If
Loop

' check in FOUND string for anyone contained in TRUCED
' in truced it is always 0330, 0330, etc

' fill array with truced clans
count = 1
start = 1
String_Length = Len(TRUCED)
Do Until start > String_Length
   Truced_Clan(count) = Mid(TRUCED, start, 4)
   start = start + 6
   count = count + 1
   If count > 30 Then
      Exit Do
   End If
Loop

' Fill array with found units

count = 1
start = 1
Interim_Found = FOUND
String_Length = Len(Interim_Found)
Do
   BRACKET = InStr(Interim_Found, ",")
   If BRACKET = 0 Then
      FOUND_UNIT(count) = Interim_Found
      Exit Do
   Else
      FOUND_UNIT(count) = Mid(Interim_Found, 1, BRACKET - 1)
   End If
   Interim_Found = Mid(Interim_Found, BRACKET + 2, String_Length)
   String_Length = Len(Interim_Found)
   If String_Length < 4 Then
      Exit Do
   End If
   count = count + 1
   If count > 30 Then
      Exit Do
   End If
Loop

' now check each unit if it matches the clan
count = 1
Do Until count > 30
   If Search_Clan = "0" & Mid(FOUND_UNIT(count), 2, 3) Then
      FOUND_UNIT(count) = "EMPTY"
   End If
   count = count + 1
   If count > 30 Then
      Exit Do
   End If
Loop

' now check each truced clan against the found_unit array
count = 1
Do
   count_2 = 1
   Do
      If Truced_Clan(count) = "0" & Mid(FOUND_UNIT(count_2), 2, 3) Then
         FOUND_UNIT(count_2) = "EMPTY"
      End If
      count_2 = count_2 + 1
      If count_2 > 30 Then
         Exit Do
      End If
   Loop
   count = count + 1
   If count > 30 Then
      Exit Do
   End If
Loop

Interim_Found = ""

count = 1
Do
   If FOUND_UNIT(count) <> "EMPTY" Then
      If Interim_Found = "" Then
         Interim_Found = FOUND_UNIT(count)
      Else
         Interim_Found = Interim_Found & ", " & FOUND_UNIT(count)
      End If
   End If
   count = count + 1
   If count > 30 Then
      Exit Do
   End If
Loop

Check_Truced = Interim_Found

EXIT_POINT:

End Function
