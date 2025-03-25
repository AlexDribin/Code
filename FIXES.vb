Attribute VB_Name = "FIXES"
Option Compare Database   'Use database order for string comparisons
Option Explicit




Public Function FIX_VALID_GOODS()
On Error GoTo FIX_VALID_GOODS_ERROR
TRIBE_STATUS = "Fix Valid Goods"

Function_Name = "FIX_VALID_GOODS"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing valid goods"
Forms![TRIBEVIBES].Repaint
    
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set ActivitiesTable = TVDBGM.OpenRecordset("VALID_GOODS")
ActivitiesTable.index = "PRIMARYKEY"

ActivitiesTable.Seek "=", "SWORD STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Sword/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "SPEAR STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Spear/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "SHIELD STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Shield/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "HELM STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Helm/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "ARAB LANCE STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Arablance/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "AXE STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Axe/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "BREASTPLATE STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "B/Plate/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "FLUTED PLATE STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "F/Plate/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "FLUTED PLATE"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Fluted Plate"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "CHAIN STEEL"

If ActivitiesTable.NoMatch Then
   ' nothing to do
Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = "Chain/Stl"
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Seek "=", "TRAWLER"

If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![GOODS] = "TRAWLER"
   ActivitiesTable![TABLE] = "SHIP"
   ActivitiesTable![SHORTNAME] = "Trawler"
   ActivitiesTable![RATING] = 0
   ActivitiesTable![BASE SELL PRICE] = 0
   ActivitiesTable![BASE BUY PRICE] = 0
   ActivitiesTable![SPRING] = 50
   ActivitiesTable![SUMMER] = 60
   ActivitiesTable![AUTUMN] = 70
   ActivitiesTable![WINTER] = 80
   ActivitiesTable![WEIGHT] = 0
   ActivitiesTable![CARRIES] = 0
   ActivitiesTable.UPDATE
End If
FIX_VALID_GOODS_ERROR_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   ActivitiesTable.Close
   Exit Function


FIX_VALID_GOODS_ERROR:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume FIX_VALID_GOODS_ERROR_CLOSE
End If

End Function

Public Function Fix_Completed_Research()
On Error GoTo error_fix_completed_research
TRIBE_STATUS = "Fix Completed Research"

Function_Name = "Fix_Completed_Research"
Function_Section = "Main"

Dim TRIBE As String
Dim TOPIC As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Fixing Research Table"
Forms![TRIBEVIBES].Repaint
    
Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")

Do Until COMPRESTAB.EOF
   If COMPRESTAB![TOPIC] = "Charring Specialists" Then
      TRIBE = COMPRESTAB![TRIBE]
      TOPIC = "Improved Charcoal Making"
      COMPRESTAB.Delete
      COMPRESTAB.AddNew
      COMPRESTAB![TRIBE] = TRIBE
      COMPRESTAB![TOPIC] = TOPIC
      COMPRESTAB.UPDATE
   End If
   COMPRESTAB.MoveNext
Loop

error_fix_completed_research_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   COMPRESTAB.Close
   
   Exit Function

error_fix_completed_research:
If (Err = 3167) Or Err = 3022 Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_fix_completed_research_CLOSE
End If


End Function

Public Function Fix_Research()
On Error GoTo error_Fix_Research
TRIBE_STATUS = "Fix Research"

Function_Name = "Fix_Research"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing Research"
Forms![TRIBEVIBES].Repaint
    
Dim TRIBE As String
Dim TOPIC As String
Dim DL_LEVEL As Long
Dim DL_LEVEL_REQD As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Fixing Research Table"
Forms![TRIBEVIBES].Repaint
    
Set COMPRESTAB = TVDB.OpenRecordset("RESEARCH")
COMPRESTAB.index = "PRIMARYKEY"
COMPRESTAB.MoveFirst
COMPRESTAB.Seek "=", "ARCHERY", "Snipers Archers nominating targets"

If Not COMPRESTAB.NoMatch Then
   COMPRESTAB.Delete
End If


error_Fix_Research_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint

COMPRESTAB.Close

   Exit Function

error_Fix_Research:
If (Err = 3167) Or Err = 3022 Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_Fix_Research_CLOSE
End If

End Function


Public Function Fix_Modifiers()
On Error GoTo error_Fix_Modifiers
TRIBE_STATUS = "FIx Modifiers"

Function_Name = "Fix_Modifiers"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing Modifiers"
Forms![TRIBEVIBES].Repaint
    

Dim trtab As Recordset
Dim crtab As Recordset
Dim POSITION As Long
Dim WORDLEN As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TribesModifiers = TVDBGM.OpenRecordset("MODIFIERS")
TribesModifiers.index = "PRIMARYKEY"
Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "PRIMARYKEY"
Set trtab = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
trtab.index = "tribe"
Set crtab = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
crtab.index = "PRIMARYKEY"
crtab.MoveFirst

Do
      trtab.Seek "=", crtab![TRIBE]
      POSITION = InStr(crtab![TOPIC], "(")
      If POSITION > 0 Then
         WORDLEN = POSITION - 1
      Else
         WORDLEN = Len(crtab![TOPIC])
      End If
If Mid(crtab![TOPIC], 1, WORDLEN) = "GOVERNMENT LEVEL 1" Then
    trtab.Edit
    trtab![GOVT LEVEL] = 1
    trtab.UPDATE
ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "GOVERNMENT LEVEL 2" Then
    trtab.Edit
    trtab![GOVT LEVEL] = 2
    trtab.UPDATE
ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "GOVERNMENT LEVEL 3" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 3
         trtab.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "GOVERNMENT LEVEL 4" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 4
         trtab.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "GOVERNMENT LEVEL 5" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 5
         trtab.UPDATE
      End If
                  
      If Mid(crtab![TOPIC], 1, WORDLEN) = "TRAPPERS" Then
         TribesModifiers.MoveFirst
         TribesModifiers.Seek "=", crtab![TRIBE], "TRAPS"
         If TribesModifiers.NoMatch Then
            TribesModifiers.AddNew
            TribesModifiers![TRIBE] = crtab![TRIBE]
            TribesModifiers![Modifier] = "TRAPS"
            TribesModifiers![AMOUNT] = 10
            TribesModifiers.UPDATE
         Else
            TribesModifiers.Edit
            TribesModifiers![AMOUNT] = 10
            TribesModifiers.UPDATE
         End If
      End If
         
      If InStr(crtab![TOPIC], "STONES/person") Then
         TribesModifiers.MoveFirst
         TribesModifiers.Seek "=", crtab![TRIBE], "STONES QUARRIED"
         If TribesModifiers.NoMatch Then
            TribesModifiers.AddNew
            TribesModifiers![TRIBE] = crtab![TRIBE]
            TribesModifiers![Modifier] = "STONES QUARRIED"
            TribesModifiers![AMOUNT] = 5
            TribesModifiers.UPDATE
            TribesModifiers.MoveFirst
            TribesModifiers.Seek "=", crtab![TRIBE], "STONES QUARRIED"
         End If
      End If
      If Mid(crtab![TOPIC], 1, WORDLEN) = "6 STONES/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 6
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "7 STONES/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 7
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "8 STONES/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 8
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "9 STONES/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 9
         TribesModifiers.UPDATE
      End If
         
      If InStr(crtab![TOPIC], "logs/person") Then
         TribesModifiers.MoveFirst
         TribesModifiers.Seek "=", crtab![TRIBE], "LOGS"
         If TribesModifiers.NoMatch Then
            TribesModifiers.AddNew
            TribesModifiers![TRIBE] = crtab![TRIBE]
            TribesModifiers![Modifier] = "LOGS"
            TribesModifiers![AMOUNT] = 4
            TribesModifiers.UPDATE
            TribesModifiers.MoveFirst
            TribesModifiers.Seek "=", crtab![TRIBE], "LOGS"
         End If
      End If
      If Mid(crtab![TOPIC], 1, WORDLEN) = "5 logs/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 5
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "6 logs/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 6
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "7 logs/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 7
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "8 logs/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 8
         TribesModifiers.UPDATE
      ElseIf Mid(crtab![TOPIC], 1, WORDLEN) = "9 logs/person" Then
         TribesModifiers.Edit
         TribesModifiers![AMOUNT] = 9
         TribesModifiers.UPDATE
      End If
         
      If Mid(crtab![TOPIC], 1, WORDLEN) = "Medicine 1" Then
         TribesModifiers.MoveFirst
         TribesModifiers.Seek "=", crtab![TRIBE], "POPULATION INCREASE"
         If TribesModifiers.NoMatch Then
            TribesModifiers.AddNew
            TribesModifiers![TRIBE] = crtab![TRIBE]
            TribesModifiers![Modifier] = "POPULATION INCREASE"
            TribesModifiers![AMOUNT] = 1
            TribesModifiers.UPDATE
         Else
            TribesModifiers.Edit
            TribesModifiers![AMOUNT] = 1
            TribesModifiers.UPDATE
         End If
      End If
         
            
      'update level 11's etc
      If Right(crtab![TOPIC], 2) = "11" Then
         SPACE_POS = InStr(crtab![TOPIC], " ")
         If SPACE_POS > 0 Then
            Skill = Left(crtab![TOPIC], (SPACE_POS - 1))
         End If
        SKILLSTABLE.MoveFirst
        SKILLSTABLE.Seek "=", crtab![TRIBE], Skill
        If Not SKILLSTABLE.NoMatch Then
           SKILLSTABLE.Edit
           SKILLSTABLE![SKILL LEVEL] = 11
           SKILLSTABLE![SUCCESSFUL] = "Y"
           SKILLSTABLE![ATTEMPTED] = "Y"
           SKILLSTABLE.UPDATE
        End If
      ElseIf Right(crtab![TOPIC], 2) = "12" Then
         SPACE_POS = InStr(crtab![TOPIC], " ")
         If SPACE_POS > 0 Then
            Skill = Left(crtab![TOPIC], (SPACE_POS - 1))
         End If
        SKILLSTABLE.MoveFirst
        SKILLSTABLE.Seek "=", crtab![TRIBE], Skill
        If Not SKILLSTABLE.NoMatch Then
           SKILLSTABLE.Edit
           SKILLSTABLE![SKILL LEVEL] = 12
           SKILLSTABLE![SUCCESSFUL] = "Y"
           SKILLSTABLE![ATTEMPTED] = "Y"
           SKILLSTABLE.UPDATE
        End If
      ElseIf Right(crtab![TOPIC], 2) = "13" Then
         SPACE_POS = InStr(crtab![TOPIC], " ")
         If SPACE_POS > 0 Then
            Skill = Left(crtab![TOPIC], (SPACE_POS - 1))
         End If
        SKILLSTABLE.MoveFirst
        SKILLSTABLE.Seek "=", crtab![TRIBE], Skill
        If Not SKILLSTABLE.NoMatch Then
           SKILLSTABLE.Edit
           SKILLSTABLE![SKILL LEVEL] = 13
           SKILLSTABLE![SUCCESSFUL] = "Y"
           SKILLSTABLE![ATTEMPTED] = "Y"
           SKILLSTABLE.UPDATE
        End If
      ElseIf Right(crtab![TOPIC], 2) = "14" Then
         SPACE_POS = InStr(crtab![TOPIC], " ")
         If SPACE_POS > 0 Then
            Skill = Left(crtab![TOPIC], (SPACE_POS - 1))
         End If
        SKILLSTABLE.MoveFirst
        SKILLSTABLE.Seek "=", crtab![TRIBE], Skill
        If Not SKILLSTABLE.NoMatch Then
           SKILLSTABLE.Edit
           SKILLSTABLE![SKILL LEVEL] = 14
           SKILLSTABLE![SUCCESSFUL] = "Y"
           SKILLSTABLE![ATTEMPTED] = "Y"
           SKILLSTABLE.UPDATE
        End If
      ElseIf Right(crtab![TOPIC], 2) = "15" Then
         SPACE_POS = InStr(crtab![TOPIC], " ")
         If SPACE_POS > 0 Then
            Skill = Left(crtab![TOPIC], (SPACE_POS - 1))
         End If
        SKILLSTABLE.MoveFirst
        SKILLSTABLE.Seek "=", crtab![TRIBE], Skill
        If Not SKILLSTABLE.NoMatch Then
           SKILLSTABLE.Edit
           SKILLSTABLE![SKILL LEVEL] = 15
           SKILLSTABLE![SUCCESSFUL] = "Y"
           SKILLSTABLE![ATTEMPTED] = "Y"
           SKILLSTABLE.UPDATE
        End If
     End If
     crtab.MoveNext
     If crtab.EOF Then
         Exit Do
     End If
Loop

error_Fix_Modifiers_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint

crtab.Close
trtab.Close
SKILLSTABLE.Close
TribesModifiers.Close

   Exit Function

error_Fix_Modifiers:
If (Err = 3167) Or Err = 3022 Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_Fix_Modifiers_CLOSE
End If


End Function


Public Function Fix_General_Info()
On Error GoTo error_fix_general_info
TRIBE_STATUS = "Fix General Info"

Dim qdfCurrent As QueryDef
Dim TLENGTH As Long
Dim UNIT_NUMBER As String

Function_Name = "Fix_General_Info"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing General Info Table"
Forms![TRIBEVIBES].Repaint
    
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TribesModifiers = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TribesModifiers.index = "PRIMARYKEY"
TribesModifiers.MoveFirst

GoTo error_fix_general_info_CLOSE

Do Until TribesModifiers.EOF
'   If TribesModifiers![TRIBE] = "0330e9" Then
'      TribesModifiers.Edit
'      TribesModifiers![current hex] = "GJ 1720"
'      TribesModifiers.UPDATE
'   End If
   TLENGTH = Len(TribesModifiers![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(TribesModifiers![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(TribesModifiers![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(TribesModifiers![TRIBE], 5, 1) = "E" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "ele" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "C" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "cou" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "F" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "fle" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "G" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "gar" & UNIT_NUMBER
      TribesModifiers.UPDATE
   End If
   End If
   If Mid(TribesModifiers![Current Hex], 3, 1) = " " Then
      'ignore
   Else
      TribesModifiers.Edit
      TribesModifiers![Current Hex] = Mid(TribesModifiers![Current Hex], 1, 2) & " " & Mid(TribesModifiers![Current Hex], 3, 4)
      TribesModifiers.UPDATE
   End If
   TribesModifiers.MoveNext
Loop

error_fix_general_info_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   TribesModifiers.Close

   Exit Function

error_fix_general_info:
If (Err = 3167) Or Err = 3022 Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_fix_general_info_CLOSE
End If


End Function

Public Function Fix_Goods()
On Error GoTo error_Fix_Goods
TRIBE_STATUS = "Fix Goods"

Dim TLENGTH As Long
Dim UNIT_NUMBER As String

Function_Name = "Fix_Goods"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing Goods"
Forms![TRIBEVIBES].Repaint
    
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_Goods")
TRIBESGOODS.MoveFirst

GoTo error_Fix_Goods_CLOSE

TRIBE_STATUS = "Fix Goods - TibesGoods"
Do Until TRIBESGOODS.EOF
   TLENGTH = Len(TRIBESGOODS![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(TRIBESGOODS![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(TRIBESGOODS![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(TRIBESGOODS![TRIBE], 5, 1) = "E" Then
      TRIBESGOODS.Edit
      TRIBESGOODS![TRIBE] = Left(TRIBESGOODS![TRIBE], 4) & "ele" & UNIT_NUMBER
      TRIBESGOODS.UPDATE
   ElseIf Mid(TRIBESGOODS![TRIBE], 5, 1) = "C" Then
      TRIBESGOODS.Edit
      TRIBESGOODS![TRIBE] = Left(TRIBESGOODS![TRIBE], 4) & "cou" & UNIT_NUMBER
      TRIBESGOODS.UPDATE
   ElseIf Mid(TRIBESGOODS![TRIBE], 5, 1) = "F" Then
      TRIBESGOODS.Edit
      TRIBESGOODS![TRIBE] = Left(TRIBESGOODS![TRIBE], 4) & "fle" & UNIT_NUMBER
      TRIBESGOODS.UPDATE
   ElseIf Mid(TRIBESGOODS![TRIBE], 5, 1) = "G" Then
      TRIBESGOODS.Edit
      TRIBESGOODS![TRIBE] = Left(TRIBESGOODS![TRIBE], 4) & "gar" & UNIT_NUMBER
      TRIBESGOODS.UPDATE
  End If
  End If
  TRIBESGOODS.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - TibesSpecialists"

Set TribesSpecialists = TVDBGM.OpenRecordset("Tribes_Specialists")
TribesSpecialists.MoveFirst

Do Until TribesSpecialists.EOF
   TLENGTH = Len(TribesSpecialists![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(TribesSpecialists![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(TribesSpecialists![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(TribesSpecialists![TRIBE], 5, 1) = "E" Then
      TribesSpecialists.Edit
      TribesSpecialists![TRIBE] = Left(TribesSpecialists![TRIBE], 4) & "ele" & UNIT_NUMBER
      TribesSpecialists.UPDATE
   ElseIf Mid(TribesSpecialists![TRIBE], 5, 1) = "C" Then
      TribesSpecialists.Edit
      TribesSpecialists![TRIBE] = Left(TribesSpecialists![TRIBE], 4) & "cou" & UNIT_NUMBER
      TribesSpecialists.UPDATE
   ElseIf Mid(TribesSpecialists![TRIBE], 5, 1) = "F" Then
      TribesSpecialists.Edit
      TribesSpecialists![TRIBE] = Left(TribesSpecialists![TRIBE], 4) & "fle" & UNIT_NUMBER
      TribesSpecialists.UPDATE
   ElseIf Mid(TribesSpecialists![TRIBE], 5, 1) = "G" Then
      TribesSpecialists.Edit
      TribesSpecialists![TRIBE] = Left(TribesSpecialists![TRIBE], 4) & "gar" & UNIT_NUMBER
      TribesSpecialists.UPDATE
   End If
   End If
  TribesSpecialists.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - HEXMAPCONST"
Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXMAPCONST.MoveFirst

Do Until HEXMAPCONST.EOF
   TLENGTH = Len(HEXMAPCONST![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(HEXMAPCONST![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(HEXMAPCONST![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(HEXMAPCONST![TRIBE], 5, 1) = "E" Then
      HEXMAPCONST.Edit
      HEXMAPCONST![TRIBE] = Left(HEXMAPCONST![TRIBE], 4) & "ele" & UNIT_NUMBER
      HEXMAPCONST.UPDATE
   ElseIf Mid(HEXMAPCONST![TRIBE], 5, 1) = "C" Then
      HEXMAPCONST.Edit
      HEXMAPCONST![TRIBE] = Left(HEXMAPCONST![TRIBE], 4) & "cou" & UNIT_NUMBER
      HEXMAPCONST.UPDATE
   ElseIf Mid(HEXMAPCONST![TRIBE], 5, 1) = "F" Then
      HEXMAPCONST.Edit
      HEXMAPCONST![TRIBE] = Left(HEXMAPCONST![TRIBE], 4) & "fle" & UNIT_NUMBER
      HEXMAPCONST.UPDATE
   ElseIf Mid(HEXMAPCONST![TRIBE], 5, 1) = "G" Then
      HEXMAPCONST.Edit
      HEXMAPCONST![TRIBE] = Left(HEXMAPCONST![TRIBE], 4) & "gar" & UNIT_NUMBER
      HEXMAPCONST.UPDATE
   End If
   End If
  HEXMAPCONST.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Hexmap Farming"
Set FarmingTable = TVDBGM.OpenRecordset("HEXMAP_FARMING")
FarmingTable.MoveFirst

Do Until FarmingTable.EOF
   TLENGTH = Len(FarmingTable![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(FarmingTable![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(FarmingTable![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(FarmingTable![TRIBE], 5, 1) = "E" Then
      FarmingTable.Edit
      FarmingTable![TRIBE] = Left(FarmingTable![TRIBE], 4) & "ele" & UNIT_NUMBER
      FarmingTable.UPDATE
   ElseIf Mid(FarmingTable![TRIBE], 5, 1) = "C" Then
      FarmingTable.Edit
      FarmingTable![TRIBE] = Left(FarmingTable![TRIBE], 4) & "cou" & UNIT_NUMBER
      FarmingTable.UPDATE
   ElseIf Mid(FarmingTable![TRIBE], 5, 1) = "F" Then
      FarmingTable.Edit
      FarmingTable![TRIBE] = Left(FarmingTable![TRIBE], 4) & "fle" & UNIT_NUMBER
      FarmingTable.UPDATE
   ElseIf Mid(FarmingTable![TRIBE], 5, 1) = "G" Then
      FarmingTable.Edit
      FarmingTable![TRIBE] = Left(FarmingTable![TRIBE], 4) & "gar" & UNIT_NUMBER
      FarmingTable.UPDATE
   End If
   End If
  FarmingTable.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Hexmap Perm Farm"
Set PermFarmingTable = TVDBGM.OpenRecordset("HEXMAP_PERMANENT_FARMING")
PermFarmingTable.MoveFirst

Do Until PermFarmingTable.EOF
   TLENGTH = Len(PermFarmingTable![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(PermFarmingTable![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(PermFarmingTable![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(PermFarmingTable![TRIBE], 5, 1) = "E" Then
      PermFarmingTable.Edit
      PermFarmingTable![TRIBE] = Left(PermFarmingTable![TRIBE], 4) & "ele" & UNIT_NUMBER
      PermFarmingTable.UPDATE
   ElseIf Mid(PermFarmingTable![TRIBE], 5, 1) = "C" Then
      PermFarmingTable.Edit
      PermFarmingTable![TRIBE] = Left(PermFarmingTable![TRIBE], 4) & "cou" & UNIT_NUMBER
      PermFarmingTable.UPDATE
   ElseIf Mid(PermFarmingTable![TRIBE], 5, 1) = "F" Then
      PermFarmingTable.Edit
      PermFarmingTable![TRIBE] = Left(PermFarmingTable![TRIBE], 4) & "fle" & UNIT_NUMBER
      PermFarmingTable.UPDATE
   ElseIf Mid(PermFarmingTable![TRIBE], 5, 1) = "G" Then
      PermFarmingTable.Edit
      PermFarmingTable![TRIBE] = Left(PermFarmingTable![TRIBE], 4) & "gar" & UNIT_NUMBER
      PermFarmingTable.UPDATE
   End If
   End If
  PermFarmingTable.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Masstransfers"
Set MODTABLE = TVDBGM.OpenRecordset("MASSTRANSFERS")
MODTABLE.MoveFirst

Do Until MODTABLE.EOF
   TLENGTH = Len(MODTABLE![From])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(MODTABLE![From], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(MODTABLE![From], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(MODTABLE![From], 5, 1) = "E" Then
      MODTABLE.Edit
      MODTABLE![From] = Left(MODTABLE![From], 4) & "ele" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![From], 5, 1) = "C" Then
      MODTABLE.Edit
      MODTABLE![From] = Left(MODTABLE![From], 4) & "cou" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![From], 5, 1) = "F" Then
      MODTABLE.Edit
      MODTABLE![From] = Left(MODTABLE![From], 4) & "fle" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![From], 5, 1) = "G" Then
      MODTABLE.Edit
      MODTABLE![From] = Left(MODTABLE![From], 4) & "gar" & UNIT_NUMBER
      MODTABLE.UPDATE
   End If
   End If
  MODTABLE.MoveNext
Loop

MODTABLE.MoveFirst

Do Until MODTABLE.EOF
   TLENGTH = Len(MODTABLE![To])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(MODTABLE![To], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(MODTABLE![To], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(MODTABLE![To], 5, 1) = "E" Then
      MODTABLE.Edit
      MODTABLE![To] = Left(MODTABLE![To], 4) & "ele" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![To], 5, 1) = "C" Then
      MODTABLE.Edit
      MODTABLE![To] = Left(MODTABLE![To], 4) & "cou" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![To], 5, 1) = "F" Then
      MODTABLE.Edit
      MODTABLE![To] = Left(MODTABLE![To], 4) & "fle" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![To], 5, 1) = "G" Then
      MODTABLE.Edit
      MODTABLE![To] = Left(MODTABLE![To], 4) & "gar" & UNIT_NUMBER
      MODTABLE.UPDATE
   End If
   End If
  MODTABLE.MoveNext
Loop


TRIBE_STATUS = "Fix Goods - Modifiers"
Set MODTABLE = TVDBGM.OpenRecordset("MODIFIERS")
MODTABLE.MoveFirst

Do Until MODTABLE.EOF
   TLENGTH = Len(MODTABLE![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(MODTABLE![TRIBE], 5, 1) = "E" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "ele" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "C" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "cou" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "F" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "fle" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "G" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "gar" & UNIT_NUMBER
      MODTABLE.UPDATE
   End If
   End If
  MODTABLE.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Process Tribes Activity"
Set MODTABLE = TVDBGM.OpenRecordset("Process_Tribes_Activity")
MODTABLE.MoveFirst

Do Until MODTABLE.EOF
   TLENGTH = Len(MODTABLE![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(MODTABLE![TRIBE], 5, 1) = "E" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "ele" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "C" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "cou" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "F" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "fle" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "G" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "gar" & UNIT_NUMBER
      MODTABLE.UPDATE
   End If
   End If
  MODTABLE.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Process Tribes Item Allocation"
Set MODTABLE = TVDBGM.OpenRecordset("Process_Tribes_Item_Allocation")
MODTABLE.MoveFirst

Do Until MODTABLE.EOF
   TLENGTH = Len(MODTABLE![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(MODTABLE![TRIBE], 5, 1) = "E" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "ele" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "C" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "cou" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "F" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "fle" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "G" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "gar" & UNIT_NUMBER
      MODTABLE.UPDATE
   End If
   End If
  MODTABLE.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Process Tribe Movement"

Set MODTABLE = TVDBGM.OpenRecordset("Process_Tribe_Movement")
MODTABLE.MoveFirst

Do Until MODTABLE.EOF
   TLENGTH = Len(MODTABLE![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(MODTABLE![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(MODTABLE![TRIBE], 5, 1) = "E" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "ele" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "C" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "cou" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "F" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "fle" & UNIT_NUMBER
      MODTABLE.UPDATE
   ElseIf Mid(MODTABLE![TRIBE], 5, 1) = "G" Then
      MODTABLE.Edit
      MODTABLE![TRIBE] = Left(MODTABLE![TRIBE], 4) & "gar" & UNIT_NUMBER
      MODTABLE.UPDATE
   End If
   End If
  MODTABLE.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Permanent Messages Table"
Set Perm_Mess_Tab = TVDBGM.OpenRecordset("Permanent_Messages_Table")
Perm_Mess_Tab.MoveFirst

Do Until Perm_Mess_Tab.EOF
   TLENGTH = Len(Perm_Mess_Tab![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(Perm_Mess_Tab![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(Perm_Mess_Tab![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(Perm_Mess_Tab![TRIBE], 5, 1) = "E" Then
      Perm_Mess_Tab.Edit
      Perm_Mess_Tab![TRIBE] = Left(Perm_Mess_Tab![TRIBE], 4) & "ele" & UNIT_NUMBER
      Perm_Mess_Tab.UPDATE
   ElseIf Mid(Perm_Mess_Tab![TRIBE], 5, 1) = "C" Then
      Perm_Mess_Tab.Edit
      Perm_Mess_Tab![TRIBE] = Left(Perm_Mess_Tab![TRIBE], 4) & "cou" & UNIT_NUMBER
      Perm_Mess_Tab.UPDATE
   ElseIf Mid(Perm_Mess_Tab![TRIBE], 5, 1) = "F" Then
      Perm_Mess_Tab.Edit
      Perm_Mess_Tab![TRIBE] = Left(Perm_Mess_Tab![TRIBE], 4) & "fle" & UNIT_NUMBER
      Perm_Mess_Tab.UPDATE
   ElseIf Mid(Perm_Mess_Tab![TRIBE], 5, 1) = "G" Then
      Perm_Mess_Tab.Edit
      Perm_Mess_Tab![TRIBE] = Left(Perm_Mess_Tab![TRIBE], 4) & "gar" & UNIT_NUMBER
      Perm_Mess_Tab.UPDATE
   End If
   End If
  Perm_Mess_Tab.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Population Increase"
Set PopTable = TVDBGM.OpenRecordset("Population_Increase")
PopTable.MoveFirst

Do Until PopTable.EOF
   TLENGTH = Len(PopTable![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(PopTable![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(PopTable![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(PopTable![TRIBE], 5, 1) = "E" Then
      PopTable.Edit
      PopTable![TRIBE] = Left(PopTable![TRIBE], 4) & "ele" & UNIT_NUMBER
      PopTable.UPDATE
   ElseIf Mid(PopTable![TRIBE], 5, 1) = "C" Then
      PopTable.Edit
      PopTable![TRIBE] = Left(PopTable![TRIBE], 4) & "cou" & UNIT_NUMBER
      PopTable.UPDATE
   ElseIf Mid(PopTable![TRIBE], 5, 1) = "F" Then
      PopTable.Edit
      PopTable![TRIBE] = Left(PopTable![TRIBE], 4) & "fle" & UNIT_NUMBER
      PopTable.UPDATE
   ElseIf Mid(PopTable![TRIBE], 5, 1) = "G" Then
      PopTable.Edit
      PopTable![TRIBE] = Left(PopTable![TRIBE], 4) & "gar" & UNIT_NUMBER
      PopTable.UPDATE
   End If
   End If
  PopTable.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Provs Availability"
Set PROVS_AVAIL_TABLE = TVDBGM.OpenRecordset("Provs_Availability")
If Not PROVS_AVAIL_TABLE.EOF Then
   PROVS_AVAIL_TABLE.MoveFirst
Do Until PROVS_AVAIL_TABLE.EOF
   TLENGTH = Len(PROVS_AVAIL_TABLE![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(PROVS_AVAIL_TABLE![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(PROVS_AVAIL_TABLE![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(PROVS_AVAIL_TABLE![TRIBE], 5, 1) = "E" Then
      PROVS_AVAIL_TABLE.Edit
      PROVS_AVAIL_TABLE![TRIBE] = Left(PROVS_AVAIL_TABLE![TRIBE], 4) & "ele" & UNIT_NUMBER
      PROVS_AVAIL_TABLE.UPDATE
   ElseIf Mid(PROVS_AVAIL_TABLE![TRIBE], 5, 1) = "C" Then
      PROVS_AVAIL_TABLE.Edit
      PROVS_AVAIL_TABLE![TRIBE] = Left(PROVS_AVAIL_TABLE![TRIBE], 4) & "cou" & UNIT_NUMBER
      PROVS_AVAIL_TABLE.UPDATE
   ElseIf Mid(PROVS_AVAIL_TABLE![TRIBE], 5, 1) = "F" Then
      PROVS_AVAIL_TABLE.Edit
      PROVS_AVAIL_TABLE![TRIBE] = Left(PROVS_AVAIL_TABLE![TRIBE], 4) & "fle" & UNIT_NUMBER
      PROVS_AVAIL_TABLE.UPDATE
   ElseIf Mid(PROVS_AVAIL_TABLE![TRIBE], 5, 1) = "G" Then
      PROVS_AVAIL_TABLE.Edit
      PROVS_AVAIL_TABLE![TRIBE] = Left(PROVS_AVAIL_TABLE![TRIBE], 4) & "gar" & UNIT_NUMBER
      PROVS_AVAIL_TABLE.UPDATE
   End If
   End If
  PROVS_AVAIL_TABLE.MoveNext
Loop

End If

TRIBE_STATUS = "Fix Goods - Process Scout Movement"
Set SCOUT_MOVEMENT_TABLE = TVDBGM.OpenRecordset("SCOUT_MOVEMENT")
SCOUT_MOVEMENT_TABLE.MoveFirst

Do Until SCOUT_MOVEMENT_TABLE.EOF
   TLENGTH = Len(SCOUT_MOVEMENT_TABLE![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(SCOUT_MOVEMENT_TABLE![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(SCOUT_MOVEMENT_TABLE![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(SCOUT_MOVEMENT_TABLE![TRIBE], 5, 1) = "E" Then
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![TRIBE] = Left(SCOUT_MOVEMENT_TABLE![TRIBE], 4) & "ele" & UNIT_NUMBER
      SCOUT_MOVEMENT_TABLE.UPDATE
   ElseIf Mid(SCOUT_MOVEMENT_TABLE![TRIBE], 5, 1) = "C" Then
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![TRIBE] = Left(SCOUT_MOVEMENT_TABLE![TRIBE], 4) & "cou" & UNIT_NUMBER
      SCOUT_MOVEMENT_TABLE.UPDATE
   ElseIf Mid(SCOUT_MOVEMENT_TABLE![TRIBE], 5, 1) = "F" Then
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![TRIBE] = Left(SCOUT_MOVEMENT_TABLE![TRIBE], 4) & "fle" & UNIT_NUMBER
      SCOUT_MOVEMENT_TABLE.UPDATE
   ElseIf Mid(SCOUT_MOVEMENT_TABLE![TRIBE], 5, 1) = "G" Then
      SCOUT_MOVEMENT_TABLE.Edit
      SCOUT_MOVEMENT_TABLE![TRIBE] = Left(SCOUT_MOVEMENT_TABLE![TRIBE], 4) & "gar" & UNIT_NUMBER
      SCOUT_MOVEMENT_TABLE.UPDATE
   End If
   End If
  SCOUT_MOVEMENT_TABLE.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Turns Activities"
Set TribesModifiers = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
TribesModifiers.MoveFirst

Do Until TribesModifiers.EOF
   TLENGTH = Len(TribesModifiers![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(TribesModifiers![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(TribesModifiers![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(TribesModifiers![TRIBE], 5, 1) = "E" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "ele" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "C" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "cou" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "F" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "fle" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "G" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "gar" & UNIT_NUMBER
      TribesModifiers.UPDATE
   End If
   End If
  TribesModifiers.MoveNext
Loop


TRIBE_STATUS = "Fix Goods - Turn Info Reqd Next Turn"
Set Turn_Info_Req_NxTurn = TVDBGM.OpenRecordset("Turn_Info_Reqd_Next_Turn")
Turn_Info_Req_NxTurn.MoveFirst

Do Until Turn_Info_Req_NxTurn.EOF
   TLENGTH = Len(Turn_Info_Req_NxTurn![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(Turn_Info_Req_NxTurn![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(Turn_Info_Req_NxTurn![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(Turn_Info_Req_NxTurn![TRIBE], 5, 1) = "E" Then
      Turn_Info_Req_NxTurn.Edit
      Turn_Info_Req_NxTurn![TRIBE] = Left(Turn_Info_Req_NxTurn![TRIBE], 4) & "ele" & UNIT_NUMBER
      Turn_Info_Req_NxTurn.UPDATE
   ElseIf Mid(Turn_Info_Req_NxTurn![TRIBE], 5, 1) = "C" Then
      Turn_Info_Req_NxTurn.Edit
      Turn_Info_Req_NxTurn![TRIBE] = Left(Turn_Info_Req_NxTurn![TRIBE], 4) & "cou" & UNIT_NUMBER
      Turn_Info_Req_NxTurn.UPDATE
   ElseIf Mid(Turn_Info_Req_NxTurn![TRIBE], 5, 1) = "F" Then
      Turn_Info_Req_NxTurn.Edit
      Turn_Info_Req_NxTurn![TRIBE] = Left(Turn_Info_Req_NxTurn![TRIBE], 4) & "fle" & UNIT_NUMBER
      Turn_Info_Req_NxTurn.UPDATE
   ElseIf Mid(Turn_Info_Req_NxTurn![TRIBE], 5, 1) = "G" Then
      Turn_Info_Req_NxTurn.Edit
      Turn_Info_Req_NxTurn![TRIBE] = Left(Turn_Info_Req_NxTurn![TRIBE], 4) & "gar" & UNIT_NUMBER
      Turn_Info_Req_NxTurn.UPDATE
   End If
   End If
  Turn_Info_Req_NxTurn.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - Under Construction"
Set TribesModifiers = TVDBGM.OpenRecordset("UNDER_CONSTRUCTION")
TribesModifiers.MoveFirst

Do Until TribesModifiers.EOF
   TLENGTH = Len(TribesModifiers![TRIBE])
   If TLENGTH = 6 Then
      UNIT_NUMBER = Right(TribesModifiers![TRIBE], 1)
   ElseIf TLENGTH = 7 Then
      UNIT_NUMBER = Right(TribesModifiers![TRIBE], 2)
   End If
   
   If TLENGTH < 8 Then
   If Mid(TribesModifiers![TRIBE], 5, 1) = "E" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "ele" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "C" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "cou" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "F" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "fle" & UNIT_NUMBER
      TribesModifiers.UPDATE
   ElseIf Mid(TribesModifiers![TRIBE], 5, 1) = "G" Then
      TribesModifiers.Edit
      TribesModifiers![TRIBE] = Left(TribesModifiers![TRIBE], 4) & "gar" & UNIT_NUMBER
      TribesModifiers.UPDATE
   End If
   End If
  TribesModifiers.MoveNext
Loop

TRIBE_STATUS = "Fix Goods - About to end"

'Call UPDATE_TRIBES_GOODS_TABLES("030", "030", "ALE", "ADD", 30000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "030", "BRANDY", "ADD", 60000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "030", "MEAD", "ADD", 50000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "030", "WINE", "ADD", 60000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "030", "COAL", "ADD", 20000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "030", "WATER", "SUBTRACT", 160000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "230", "ALE", "ADD", 60000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "230", "MEAD", "ADD", 50000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "230", "WINE", "ADD", 30000)
'Call UPDATE_TRIBES_GOODS_TABLES("030", "230", "WATER", "SUBTRACT", 30000)

error_Fix_Goods_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint

   Exit Function

error_Fix_Goods:
If (Err = 3167) Or (Err = 3022) Or (Err = 3163) Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_Fix_Goods_CLOSE
End If

End Function

Public Function FIX_TABLE()
On Error GoTo error_FIX_TABLE
TRIBE_STATUS = "Fix Table"

Dim skilltab As Recordset        ' PROCESS_SKILLS
Dim researchtab As Recordset     ' PROCESS_RESEARCH
Dim strSQL As String

Function_Name = "FIX_TABLE"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing table"
Forms![TRIBEVIBES].Repaint
    
Set TVWKSPACE = DBEngine.Workspaces(0)
Dim td As TableDef
Dim fld As field
Dim tblName As String
Dim fldSize As Integer
Dim myP As Property

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
'TVDBGM.Execute "ALTER TABLE RESEARCH_ATTEMPTS ADD COLUMN Cost number;"
'TVDBGM.Execute "ALTER TABLE SKILL_ATTEMPTS ADD COLUMN Cost number;"
'TVDBGM.Execute "ALTER TABLE TRIBE_RESEARCH ADD COLUMN Cost number;"
'TVDBGM.Execute "ALTER TABLE TRIBES_PROCESSING ADD COLUMN Warriors_Assigned number;"
'TVDBGM.Execute "ALTER TABLE Tribes_General_Info ADD COLUMN Commodity_2 text(50);"
'TVDBGM.Execute "ALTER TABLE Process_Tribes_Activity add COLUMN BUILDING Number;"
'TVDBGM.Execute "ALTER TABLE Process_Tribes_Activity_Copy add COLUMN BUILDING Number;"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement drop COLUMN PROCESSED);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_31 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_32 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_33 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_34 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_35 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_36 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_37 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_38 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_39 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN MOVEMENT_40 Text(6);"
TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement add COLUMN PROCESSED Text(1);"
'TVDBGM.Execute "ALTER TABLE Process_Tribes_Activity DROP COLUMN ORDER;"
'TVDBGM.Execute "ALTER TABLE Process_Skills ADD COLUMN Comment TEXT(50);"
'TVDBGM.Execute "ALTER TABLE Special_Transfer_Routes ADD COLUMN Unit2 TEXT(50);"
'TVDBGM.Execute "CREATE INDEX Route_Name ON Special_Transfer_Routes (Route_Name);"
'TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement alter COLUMN MOVEMENT_2 TEXT(6);"
'TVDBGM.Execute "ALTER TABLE HEX_MAP_CITY alter COLUMN CITY_2 TEXT(100);"
'TVDBGM.Execute "ALTER TABLE MASSTRANSFERS ADD COLUMN REPORT_CLAN TEXT(255);"
'TVDBGM.Execute "ALTER TABLE MASSTRANSFERS ADD COLUMN PROCESS_MSG TEXT(255);"

'TVDBGM.Execute "ALTER TABLE TRIBES_general_info ADD COLUMN GT_WALKING_CAPACITY Number;"

'TVDBGM.Execute "CREATE INDEX SECONDARYKEY ON Process_Tribe_Movement (FOLLOW_TRIBE);"

'to delete a field from a table
'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Process_Tribes_Activity;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Process_Tribes_Activity (TRIBE);"
'TVDBGM.Execute "ALTER TABLE Process_Tribes_Activity DROP COLUMN ORDER;"

'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Process_Tribes_Activity_Copy;"
'TVDBGM.Execute "ALTER TABLE Process_Tribes_Activity_Copy DROP COLUMN ORDER;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Process_Tribes_Activity_Copy (TRIBE);"

'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Scout_Movement;"
'TVDBGM.Execute "ALTER TABLE Scout_Movement DROP COLUMN SCOUT;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Scout_Movement (TRIBE);"

'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Process_Tribes_Item_Allocation_Copy;"
'TVDBGM.Execute "ALTER TABLE Process_Tribes_Item_Allocation_Copy DROP COLUMN CLAN;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Process_Tribes_Item_Allocation_Copy (TRIBE, ACTIVITY, ITEM, ITEM_USED);"

'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Process_Scout_Movement;"
'TVDBGM.Execute "ALTER TABLE Process_Scout_Movement DROP COLUMN CLAN;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Process_Scout_Movement (TRIBE, SCOUT);"

'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Process_Tribe_Movement;"
'TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement DROP COLUMN CLAN;"
'TVDBGM.Execute "ALTER TABLE Process_Tribe_Movement DROP COLUMN FOLLOW_CLAN;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Process_Tribe_Movement (TRIBE);"

'TVDBGM.Execute "DROP INDEX PRIMARYKEY ON Process_Skills;"
'TVDBGM.Execute "ALTER TABLE Process_Skills DROP COLUMN CLAN;"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON Process_Skills (TRIBE,ORDER);"

'DoCmd.CopyObject FILEGM, "Process_Scout_Movement_Copy", acTable, "Process_Scout_Movement"
'DoCmd.TransferDatabase acExport, "Microsoft Access", FILEGM, acTable, "Process_Scout_Movement_Copy", "Process_Scout_Movement", True
'TVDBGM.Execute "CREATE TABLE PROCESS_RESEARCH (TRIBE TEXT(10), ORDER TEXT(5), TOPIC TEXT(50), PROCESSED TEXT(1));"
'TVDBGM.Execute "CREATE INDEX PRIMARYKEY ON PROCESS_RESEARCH (TRIBE, ORDER);"


' the following inserts the new field into a specific spot in the table
'Set td = TVDBGM.TableDefs("Process_Tribes_Activity")
'For Each fld In td.Fields
'   With fld
'        If .OrdinalPosition >= 9 Then
'            .OrdinalPosition = .OrdinalPosition + 1
'        End If
'    End With
'Next fld

'With td
'    .Fields![TRIBE].OrdinalPosition = 1
'    .Fields![ORDER].OrdinalPosition = 2
'    .Fields![ACTIVITY].OrdinalPosition = 3
'    .Fields![ITEM].OrdinalPosition = 4
'    .Fields![DISTINCTION].OrdinalPosition = 5
'    .Fields![PEOPLE].OrdinalPosition = 6
'    .Fields![Slaves].OrdinalPosition = 7
'    .Fields![SPECIALISTS].OrdinalPosition = 8
'    .Fields![OWNING_TRIBE].OrdinalPosition = 9
'    .Fields![Number_of_Seeking_Groups].OrdinalPosition = 11
'    .Fields![Whale_Size].OrdinalPosition = 12
'    .Fields![MINING_DIRECTION].OrdinalPosition = 13
'    .Fields![PROCESSED].OrdinalPosition = 14
'    .Fields![BUILDING].OrdinalPosition = 10
'    .Fields.Refresh
'End With

'With td
'  .Fields.Append .CreateField("BUILDING", dbInteger, 1)
'  .Fields![BUILDING].OrdinalPosition = 9
'  .Fields![BUILDING].Required = False
'  .Fields.Refresh
'End With
'With td
'  .Fields.Append .CreateField("FOLLOW_TRIBE", dbText, 10)
'  .Fields![Follow_Tribe].OrdinalPosition = 3
'  .Fields![Follow_Tribe].Required = False
'  .Fields![Follow_Tribe].AllowZeroLength = True
'  .Fields.Refresh
'End With

'Set TribesModifiers = TVDBGM.OpenRecordset("GM_Costs_Table")
'TribesModifiers.AddNew
'TribesModifiers![Group] = "Courier"
'TribesModifiers![COST] = 1.4
'TribesModifiers.UPDATE
 
' DoCmd.TransferDatabase acExport, "Microsoft Access", FILEGM, acTable, "MassTransfers_new", "MassTransfers", True
' DoCmd.TransferDatabase acExport, "Microsoft Access", FILEGM, acTable, "Special_Transfer_Routes_new", "Special_Transfer_Routes", True

' GO THROUGH MODIFIERS AND WHERE THE SUB GROUP HAS A TRIBE MOVEMENT MODIFIER
' THEN ENSURE THAT THE PRIMARY TRIBE HAS THE SAME MODIFIER





error_FIX_TABLE_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint

TVDBGM.Close
   
   Exit Function

error_FIX_TABLE:
If Err = 3010 Or Err = 3167 Or Err = 3022 Or Err = 3191 Or Err = 3372 Or Err = 3375 Or Err = 3380 Or Err = 3381 Then ' if record deleted then continue.
   Resume Next
ElseIf Err = 3021 Then
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_FIX_TABLE_CLOSE
End If


End Function


Public Function FIX_HEXMAP()
On Error GoTo error_FIX_HEXMAP
TRIBE_STATUS = "Fix Hexmap"

Function_Name = "FIX_HEXMAP"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing Hexmap"
Forms![TRIBEVIBES].Repaint
    
Dim BORDER_N As String
Dim BORDER_NE As String
Dim BORDER_SE As String
Dim BORDER_S As String
Dim BORDER_SW As String
Dim BORDER_NW As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TribesModifiers = TVDBGM.OpenRecordset("HEX_MAP")
TribesModifiers.index = "PRIMARYKEY"
TribesModifiers.MoveFirst

Do
   BORDER_N = "N"
   BORDER_NE = "N"
   BORDER_SE = "N"
   BORDER_S = "N"
   BORDER_SW = "N"
   BORDER_NW = "N"
   If Mid(TribesModifiers![BEACHES], 1, 1) = "Y" Then
      BORDER_N = "BE"
   End If
   If Mid(TribesModifiers![BEACHES], 2, 1) = "Y" Then
      BORDER_NE = "BE"
   End If
   If Mid(TribesModifiers![BEACHES], 3, 1) = "Y" Then
      BORDER_SE = "BE"
   End If
   If Mid(TribesModifiers![BEACHES], 4, 1) = "Y" Then
      BORDER_S = "BE"
   End If
   If Mid(TribesModifiers![BEACHES], 5, 1) = "Y" Then
      BORDER_SW = "BE"
   End If
   If Mid(TribesModifiers![BEACHES], 6, 1) = "Y" Then
      BORDER_NW = "BE"
   End If
   If Mid(TribesModifiers![CANALS], 1, 1) = "Y" Then
      BORDER_N = "CA"
   End If
   If Mid(TribesModifiers![CANALS], 2, 1) = "Y" Then
      BORDER_NE = "CA"
   End If
   If Mid(TribesModifiers![CANALS], 3, 1) = "Y" Then
      BORDER_SE = "CA"
   End If
   If Mid(TribesModifiers![CANALS], 4, 1) = "Y" Then
      BORDER_S = "CA"
   End If
   If Mid(TribesModifiers![CANALS], 5, 1) = "Y" Then
      BORDER_SW = "CA"
   End If
   If Mid(TribesModifiers![CANALS], 6, 1) = "Y" Then
      BORDER_NW = "CA"
   End If
   If Mid(TribesModifiers![CLIFFS], 1, 1) = "Y" Then
      BORDER_N = "CL"
   End If
   If Mid(TribesModifiers![CLIFFS], 2, 1) = "Y" Then
      BORDER_NE = "CL"
   End If
   If Mid(TribesModifiers![CLIFFS], 3, 1) = "Y" Then
      BORDER_SE = "CL"
   End If
   If Mid(TribesModifiers![CLIFFS], 4, 1) = "Y" Then
      BORDER_S = "CL"
   End If
   If Mid(TribesModifiers![CLIFFS], 5, 1) = "Y" Then
      BORDER_SW = "CL"
   End If
   If Mid(TribesModifiers![CLIFFS], 6, 1) = "Y" Then
      BORDER_NW = "CL"
   End If
   If Mid(TribesModifiers![PASSES], 1, 1) = "Y" Then
      BORDER_N = "PA"
   End If
   If Mid(TribesModifiers![PASSES], 2, 1) = "Y" Then
      BORDER_NE = "PA"
   End If
   If Mid(TribesModifiers![PASSES], 3, 1) = "Y" Then
      BORDER_SE = "PA"
   End If
   If Mid(TribesModifiers![PASSES], 4, 1) = "Y" Then
      BORDER_S = "PA"
   End If
   If Mid(TribesModifiers![PASSES], 5, 1) = "Y" Then
      BORDER_SW = "PA"
   End If
   If Mid(TribesModifiers![PASSES], 6, 1) = "Y" Then
      BORDER_NW = "PA"
   End If
   If Mid(TribesModifiers![RIVERS], 1, 1) = "Y" Then
      BORDER_N = "RI"
   End If
   If Mid(TribesModifiers![RIVERS], 2, 1) = "Y" Then
      BORDER_NE = "RI"
   End If
   If Mid(TribesModifiers![RIVERS], 3, 1) = "Y" Then
      BORDER_SE = "RI"
   End If
   If Mid(TribesModifiers![RIVERS], 4, 1) = "Y" Then
      BORDER_S = "RI"
   End If
   If Mid(TribesModifiers![RIVERS], 5, 1) = "Y" Then
      BORDER_SW = "RI"
   End If
   If Mid(TribesModifiers![RIVERS], 6, 1) = "Y" Then
      BORDER_NW = "RI"
   End If
   If Mid(TribesModifiers![FORDS], 1, 1) = "Y" Then
      BORDER_N = "FO"
   End If
   If Mid(TribesModifiers![FORDS], 2, 1) = "Y" Then
      BORDER_NE = "FO"
   End If
   If Mid(TribesModifiers![FORDS], 3, 1) = "Y" Then
      BORDER_SE = "FO"
   End If
   If Mid(TribesModifiers![FORDS], 4, 1) = "Y" Then
      BORDER_S = "FO"
   End If
   If Mid(TribesModifiers![FORDS], 5, 1) = "Y" Then
      BORDER_SW = "FO"
   End If
   If Mid(TribesModifiers![FORDS], 6, 1) = "Y" Then
      BORDER_NW = "FO"
   End If
   If BORDER_N = "N" Then
      BORDER_N = "NN"
   End If
   If BORDER_NE = "N" Then
      BORDER_NE = "NN"
   End If
   If BORDER_SE = "N" Then
      BORDER_SE = "NN"
   End If
   If BORDER_S = "N" Then
      BORDER_S = "NN"
   End If
   If BORDER_SW = "N" Then
      BORDER_SW = "NN"
   End If
   If BORDER_NW = "N" Then
      BORDER_NW = "NN"
   End If
   TribesModifiers.Edit
   TribesModifiers![Borders] = BORDER_N & BORDER_NE & BORDER_SE & BORDER_S & BORDER_SW & BORDER_NW
   TribesModifiers.UPDATE
   
   TribesModifiers.MoveNext
   If TribesModifiers.EOF Then
      Exit Do
   End If
Loop
 
error_FIX_HEXMAP_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   TribesModifiers.Close

   Exit Function

error_FIX_HEXMAP:
If (Err = 3167) Or Err = 3022 Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_FIX_HEXMAP_CLOSE
End If



End Function
Public Function FIX_UNDER_CONSTRUCTION()
On Error GoTo error_FIX_UNDER_CONSTRUCTION
TRIBE_STATUS = "Fix Under COnstruction"

Function_Name = "FIX_UNDER_CONSTRUCTION"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Fixing under construction"
Forms![TRIBEVIBES].Repaint
    
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TribesModifiers = TVDBGM.OpenRecordset("UNDER_CONSTRUCTION")
TribesModifiers.index = "PRIMARYKEY"
TribesModifiers.MoveFirst
TribesModifiers.Seek "=", "030", "030", "BRICKWORKS"

If Not TribesModifiers.NoMatch Then
   TribesModifiers.Delete
End If

TribesModifiers.Seek "=", "030", "030", "CATHEDRAL"

If Not TribesModifiers.NoMatch Then
   TribesModifiers.Edit
   TribesModifiers![LOGS] = 686
   TribesModifiers![STONES] = 2540
   TribesModifiers![BRASS] = 140
   TribesModifiers.UPDATE
End If

error_FIX_UNDER_CONSTRUCTION_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint

TribesModifiers.Close
   
   Exit Function

error_FIX_UNDER_CONSTRUCTION:
If (Err = 3167) Or Err = 3022 Then  ' if record deleted then continue.
   Resume Next
Else
   Call A999_ERROR_HANDLING
   Resume error_FIX_UNDER_CONSTRUCTION_CLOSE
End If

End Function


Public Function Populate_GM_Costs_Table()
On Error GoTo error_Populate_GM_Costs_Table
TRIBE_STATUS = "Populate GM Costs Table"

Function_Name = "Populate_GM_Costs_Table"
Function_Section = "Main"

Forms![TRIBEVIBES]![Status] = "Populate GM Costs"
Forms![TRIBEVIBES].Repaint
    
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Set TribesModifiers = TVDBGM.OpenRecordset("GM_Costs_Table")
TribesModifiers.AddNew
TribesModifiers![Group] = "COST CLAN"
TribesModifiers![Cost] = 3
TribesModifiers.UPDATE
TribesModifiers.AddNew
TribesModifiers![Group] = "TRIBE"
TribesModifiers![Cost] = 1
TribesModifiers.UPDATE
TribesModifiers.AddNew
TribesModifiers![Group] = "VILLAGE"
TribesModifiers![Cost] = 1
TribesModifiers.UPDATE
TribesModifiers.AddNew
TribesModifiers![Group] = "ELEMENT"
TribesModifiers![Cost] = 0.5
TribesModifiers.UPDATE
TribesModifiers.AddNew
TribesModifiers![Group] = "FLEET"
TribesModifiers![Cost] = 0.5
TribesModifiers.UPDATE
TribesModifiers.AddNew
TribesModifiers![Group] = "GARRISON"
TribesModifiers![Cost] = 0.5
TribesModifiers.UPDATE
TribesModifiers.AddNew
TribesModifiers![Group] = "BANDIT"
TribesModifiers![Cost] = 0
TribesModifiers.UPDATE

error_Populate_GM_Costs_Table_CLOSE:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   TribesModifiers.Close
   
   Exit Function


error_Populate_GM_Costs_Table:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume error_Populate_GM_Costs_Table_CLOSE
End If


End Function
Function RESET_PROCESSED()
Dim qdfCurrent As QueryDef

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

GMTABLE.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

'Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE Scout_Movement SET Scout_Movement.[PROCESSED] = 'N';")
'qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE Process_skill SET Process_Skill.[PROCESSED] = 'N';")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "UPDATE Process_Research SET Process_Research.[PROCESSED] = 'N';")
qdfCurrent.Execute

End Function

Public Function A999_FIXES_ERROR_HANDLING()
  errorstring = Err.Description
 
  Msg = "The Process " & Function_Name & " " & Function_Section & " has received the following error message "
  Msg = Msg & Chr(13) & Chr(10) & "Error # " & Err & " " & Error$ & " " & errorstring & Chr(13) & Chr(10)
  Msg = Msg & Chr(13) & Chr(10) & " GIVE THIS INFO TO Jeff."
  MsgBox (Msg)


End Function

Public Function Close_the_Table()
    Close_Table ("Valid_Ships")
End Function
