Attribute VB_Name = "SCREEN_HANDLING"
Option Compare Database
Option Explicit
Global MYFORM As Form
Global AVAILABLESKILLS As Recordset

Public Function OPEN_FORM_TRIBES_GOODS(ITEM_TYPE)

' SEE IF FORM ALREADY OPEN - IF SO CLOSE IT
If ITEM_TYPE = "MODIFIERS" Then
   DoCmd.Close acForm, "TRIBES - MODIFIERS"
   DoCmd.OpenForm "TRIBES - MODIFIERS"
   DoCmd.Maximize
   
Else
   DoCmd.Close acForm, "TRIBES - GOODS"
   DoCmd.OpenForm "TRIBES - GOODS"
   DoCmd.Maximize
   
End If

If ITEM_TYPE = "ANIMAL" Then
   Forms![TRIBES - GOODS]![ITEM_TYPE] = "ANIMAL"
   Forms![TRIBES - GOODS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - GOODS]![TITLE_BAR] = "TRIBES - ANIMALS"
ElseIf ITEM_TYPE = "FINISHED" Then
   Forms![TRIBES - GOODS]![ITEM_TYPE] = "FINISHED"
   Forms![TRIBES - GOODS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - GOODS]![TITLE_BAR] = "TRIBES - FINISHED GOODS"
ElseIf ITEM_TYPE = "MINERAL" Then
   Forms![TRIBES - GOODS]![ITEM_TYPE] = "MINERAL"
   Forms![TRIBES - GOODS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - GOODS]![TITLE_BAR] = "TRIBES - MINERALS"
ElseIf ITEM_TYPE = "RAW" Then
   Forms![TRIBES - GOODS]![ITEM_TYPE] = "RAW"
   Forms![TRIBES - GOODS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - GOODS]![TITLE_BAR] = "TRIBES - RAW MATERIALS"
ElseIf ITEM_TYPE = "SHIP" Then
   Forms![TRIBES - GOODS]![ITEM_TYPE] = "SHIP"
   Forms![TRIBES - GOODS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - GOODS]![TITLE_BAR] = "TRIBES - SHIPS"
ElseIf ITEM_TYPE = "WAR" Then
   Forms![TRIBES - GOODS]![ITEM_TYPE] = "WAR"
   Forms![TRIBES - GOODS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - GOODS]![TITLE_BAR] = "TRIBES - WAR EQUIPMENT"
ElseIf ITEM_TYPE = "MODIFIERS" Then
   Forms![TRIBES - MODIFIERS]![ITEM_TYPE] = "MODIFIERS"
   Forms![TRIBES - MODIFIERS]![ITEM_TYPE].Visible = False
   Forms![TRIBES - MODIFIERS]![TITLE_BAR] = "TRIBES - MODIFIERS"
End If

End Function


Public Function EXIT_FORMS(SCREEN_NAME)
On Error GoTo Err_EXIT_FORMS

' if gm = ' jeff then exit form else
'    DoCmd.Quit

DoCmd.Close acForm, SCREEN_NAME

Exit_EXIT_FORMS:
    Exit Function

Err_EXIT_FORMS:
    MsgBox Err.Description
    Resume Exit_EXIT_FORMS
    
End Function


Public Function Check_Current_Skill_Level(Skill_Attempt As String)
Dim FORM_FIELD As String

If IsNull(Forms![SKILLS_1]![TRIBE NAME]) Then
   Exit Function
End If

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILETV = GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILETV, False, False)

Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "PRIMARYKEY"

If Skill_Attempt = "Primary" Then
   SKILLSTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], Forms![SKILLS_1]![PRIMARY SKILL ATTEMPT]
   If SKILLSTABLE.NoMatch Then
      Forms![SKILLS_1]![PRIMARY SKILL LEVEL] = 0
   Else
      Forms![SKILLS_1]![PRIMARY SKILL LEVEL] = SKILLSTABLE![SKILL LEVEL]
   End If
ElseIf Skill_Attempt = "Secondary" Then
   SKILLSTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], Forms![SKILLS_1]![SECONDARY SKILL ATTEMPT]
   If SKILLSTABLE.NoMatch Then
      Forms![SKILLS_1]![SECONDARY SKILL LEVEL] = 0
   Else
      Forms![SKILLS_1]![SECONDARY SKILL LEVEL] = SKILLSTABLE![SKILL LEVEL]
   End If
ElseIf Skill_Attempt = "Tertiary" Then
   SKILLSTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], Forms![SKILLS_1]![TERTIARY SKILL ATTEMPT]
   If SKILLSTABLE.NoMatch Then
      Forms![SKILLS_1]![TERTIARY SKILL LEVEL] = 0
   Else
      Forms![SKILLS_1]![TERTIARY SKILL LEVEL] = SKILLSTABLE![SKILL LEVEL]
   End If
End If

SKILLSTABLE.Close

End Function

Function Determine_Research_Allowances()
Dim FORM_FIELD As String
Dim MAP_REFERENCE As String
Dim AVAILABLE_PEOPLE As Long
Dim MRT_FOUND As String
Dim MRT1_FOUND As String
Dim MRT2_FOUND As String
Dim LABS_FOUND As String
Dim SEM_FOUND As String
Dim BOOK_FOUND As Long
Dim LEVEL_10S As Long
Dim uni_found As String
Dim TOPICS As Double
Dim CHECK_TOPICS As Long
Dim COUNT_TOPICS As Double
Dim i As Long
Dim QUERY_STRING As String
Dim POSITION As Long
Dim WORDLEN As Long
Dim SEARCHVALUE As String

If IsNull(Forms![SKILLS_1]![TRIBE NAME]) Then
   Exit Function
End If

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM AVAILABLE_RESEARCH;")
qdfCurrent.Execute

Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM AVAILABLE_SKILLS;")
qdfCurrent.Execute

Set RESEARCH_TABLE = TVDB.OpenRecordset("RESEARCH")
RESEARCH_TABLE.index = "TOPIC"

Set TRIBE_RESEARCH = TVDBGM.OpenRecordset("TRIBE_RESEARCH")
TRIBE_RESEARCH.index = "SECONDARYKEY"
TRIBE_RESEARCH.Seek "=", Forms![SKILLS_1]![TRIBE NAME]
   
Set MOVEFORM = Forms!SKILLS_1
count = 1

If TRIBE_RESEARCH.NoMatch Then
   'Clear topics, just in case
   Do Until count > 40
      stext1 = "RESEARCH TOPIC " & CStr(count)
      stext2 = "RESEARCH SKILL" & CStr(count)
      MOVEFORM(stext1) = "EMPTY"
      MOVEFORM(stext2) = "NIL"
      count = count + 1
   Loop
   GoTo CALC_RESEARCH_TOPICS
Else
   'Clear topics just in case
   Do Until count > 40
      stext1 = "RESEARCH TOPIC " & CStr(count)
      stext2 = "RESEARCH SKILL" & CStr(count)
      MOVEFORM(stext1) = "EMPTY"
      MOVEFORM(stext2) = "NIL"
      count = count + 1
   Loop
End If

count = 1
Do Until Not TRIBE_RESEARCH![TRIBE] = Forms![SKILLS_1]![TRIBE NAME]
   TRIBE_RESEARCH.Edit
   POSITION = InStr(TRIBE_RESEARCH![TOPIC], "(")
   WORDLEN = Len(TRIBE_RESEARCH![TOPIC])
   If Right(Mid(TRIBE_RESEARCH![TOPIC], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(TRIBE_RESEARCH![TOPIC], 1, WORDLEN)
   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   
   Select Case count
   Case 1
      Forms![SKILLS_1]![RESEARCH TOPIC 1] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL1] = RESEARCH_TABLE![Skill]
   Case 2
      Forms![SKILLS_1]![RESEARCH TOPIC 2] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL2] = RESEARCH_TABLE![Skill]
   Case 3
      Forms![SKILLS_1]![RESEARCH TOPIC 3] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL3] = RESEARCH_TABLE![Skill]
   Case 4
      Forms![SKILLS_1]![RESEARCH TOPIC 4] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL4] = RESEARCH_TABLE![Skill]
   Case 5
      Forms![SKILLS_1]![RESEARCH TOPIC 5] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL5] = RESEARCH_TABLE![Skill]
   Case 6
      Forms![SKILLS_1]![RESEARCH TOPIC 6] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL6] = RESEARCH_TABLE![Skill]
   Case 7
      Forms![SKILLS_1]![RESEARCH TOPIC 7] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL7] = RESEARCH_TABLE![Skill]
   Case 8
      Forms![SKILLS_1]![RESEARCH TOPIC 8] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL8] = RESEARCH_TABLE![Skill]
   Case 9
      Forms![SKILLS_1]![RESEARCH TOPIC 9] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL9] = RESEARCH_TABLE![Skill]
   Case 10
      Forms![SKILLS_1]![RESEARCH TOPIC 10] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL10] = RESEARCH_TABLE![Skill]
   Case 11
      Forms![SKILLS_1]![RESEARCH TOPIC 11] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL11] = RESEARCH_TABLE![Skill]
   Case 12
      Forms![SKILLS_1]![RESEARCH TOPIC 12] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL12] = RESEARCH_TABLE![Skill]
   Case 13
      Forms![SKILLS_1]![RESEARCH TOPIC 13] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL13] = RESEARCH_TABLE![Skill]
   Case 14
      Forms![SKILLS_1]![RESEARCH TOPIC 14] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL14] = RESEARCH_TABLE![Skill]
   Case 15
      Forms![SKILLS_1]![RESEARCH TOPIC 15] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL15] = RESEARCH_TABLE![Skill]
   Case 16
      Forms![SKILLS_1]![RESEARCH TOPIC 16] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL16] = RESEARCH_TABLE![Skill]
   Case 17
      Forms![SKILLS_1]![RESEARCH TOPIC 17] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL17] = RESEARCH_TABLE![Skill]
   Case 18
      Forms![SKILLS_1]![RESEARCH TOPIC 18] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL18] = RESEARCH_TABLE![Skill]
   Case 19
      Forms![SKILLS_1]![RESEARCH TOPIC 19] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL19] = RESEARCH_TABLE![Skill]
   Case 20
      Forms![SKILLS_1]![RESEARCH TOPIC 20] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL20] = RESEARCH_TABLE![Skill]
   Case 21
      Forms![SKILLS_1]![RESEARCH TOPIC 21] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL21] = RESEARCH_TABLE![Skill]
   Case 22
      Forms![SKILLS_1]![RESEARCH TOPIC 22] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL22] = RESEARCH_TABLE![Skill]
   Case 23
      Forms![SKILLS_1]![RESEARCH TOPIC 23] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL23] = RESEARCH_TABLE![Skill]
   Case 24
      Forms![SKILLS_1]![RESEARCH TOPIC 24] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL24] = RESEARCH_TABLE![Skill]
   Case 25
      Forms![SKILLS_1]![RESEARCH TOPIC 25] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL25] = RESEARCH_TABLE![Skill]
   Case 26
      Forms![SKILLS_1]![RESEARCH TOPIC 26] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL26] = RESEARCH_TABLE![Skill]
   Case 27
      Forms![SKILLS_1]![RESEARCH TOPIC 27] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL27] = RESEARCH_TABLE![Skill]
   Case 28
      Forms![SKILLS_1]![RESEARCH TOPIC 28] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL28] = RESEARCH_TABLE![Skill]
   Case 29
      Forms![SKILLS_1]![RESEARCH TOPIC 29] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL29] = RESEARCH_TABLE![Skill]
   Case 30
      Forms![SKILLS_1]![RESEARCH TOPIC 30] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL30] = RESEARCH_TABLE![Skill]
   Case 31
      Forms![SKILLS_1]![RESEARCH TOPIC 31] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL31] = RESEARCH_TABLE![Skill]
   Case 32
      Forms![SKILLS_1]![RESEARCH TOPIC 32] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL32] = RESEARCH_TABLE![Skill]
   Case 33
      Forms![SKILLS_1]![RESEARCH TOPIC 33] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL33] = RESEARCH_TABLE![Skill]
   Case 34
      Forms![SKILLS_1]![RESEARCH TOPIC 34] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL34] = RESEARCH_TABLE![Skill]
   Case 35
      Forms![SKILLS_1]![RESEARCH TOPIC 35] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL35] = RESEARCH_TABLE![Skill]
   Case 36
      Forms![SKILLS_1]![RESEARCH TOPIC 36] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL36] = RESEARCH_TABLE![Skill]
   Case 37
      Forms![SKILLS_1]![RESEARCH TOPIC 37] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL37] = RESEARCH_TABLE![Skill]
   Case 38
      Forms![SKILLS_1]![RESEARCH TOPIC 38] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL38] = RESEARCH_TABLE![Skill]
   Case 39
      Forms![SKILLS_1]![RESEARCH TOPIC 39] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL39] = RESEARCH_TABLE![Skill]
   Case 40
      Forms![SKILLS_1]![RESEARCH TOPIC 40] = TRIBE_RESEARCH![TOPIC]
      Forms![SKILLS_1]![RESEARCH SKILL40] = RESEARCH_TABLE![Skill]
   End Select
   TRIBE_RESEARCH.MoveNext
   count = count + 1
   If TRIBE_RESEARCH.EOF Then
      Exit Do
   End If
Loop

TRIBE_RESEARCH.Close

CALC_RESEARCH_TOPICS:

Set TRIBESTABLE = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESTABLE.index = "PRIMARYKEY"
TRIBESTABLE.MoveFirst
TRIBESTABLE.Seek "=", Forms![SKILLS_1]![CLAN NAME], Forms![SKILLS_1]![TRIBE NAME]

TRIBESTABLE.Edit
MAP_REFERENCE = TRIBESTABLE![Current Hex]
AVAILABLE_PEOPLE = TRIBESTABLE![WARRIORS] + TRIBESTABLE![ACTIVES] + TRIBESTABLE![INACTIVES]
TRIBESTABLE.Close
   
Set COMPRESTABLE = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTABLE.index = "PRIMARYKEY"
COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], "LAB ASSISTANTS"
   
If COMPRESTABLE.NoMatch Then
   LABS_FOUND = "NO"
Else
   LABS_FOUND = "YES"
End If

COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], "Seminary"
If COMPRESTABLE.NoMatch Then
   SEM_FOUND = "NO"
Else
   SEM_FOUND = "YES"
End If

COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], "Multiple Research Topics 2"
If COMPRESTABLE.NoMatch Then
   MRT2_FOUND = "NO"
Else
   MRT2_FOUND = "YES"
End If

COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], "Multiple Topic development 2"
If COMPRESTABLE.NoMatch Then
   MRT2_FOUND = "NO"
Else
   MRT2_FOUND = "YES"
End If

COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], "Multiple Topic development 1"
If COMPRESTABLE.NoMatch Then
   MRT_FOUND = "NO"
Else
   MRT_FOUND = "YES"
End If

COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], "Multiple Topics"
If COMPRESTABLE.NoMatch Then
   MRT_FOUND = "NO"
Else
   MRT_FOUND = "YES"
End If

COMPRESTABLE.Close

Set ValidSkills = TVDB.OpenRecordset("VALID_SKILLS")
ValidSkills.MoveFirst

Set AVAILABLESKILLS = TVDB.OpenRecordset("AVAILABLE_SKILLS")

Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "PRIMARYKEY"

Do
     SKILLSTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], ValidSkills![Skill]
     If SKILLSTABLE.NoMatch Then
        AVAILABLESKILLS.AddNew
        AVAILABLESKILLS![Skill] = ValidSkills![Skill]
        AVAILABLESKILLS.UPDATE
     ElseIf SKILLSTABLE![SKILL LEVEL] < 10 Then
        AVAILABLESKILLS.AddNew
        AVAILABLESKILLS![Skill] = ValidSkills![Skill]
        AVAILABLESKILLS.UPDATE
     End If
     ValidSkills.MoveNext
     If ValidSkills.EOF Then
         Exit Do
     End If
Loop

AVAILABLESKILLS.Close
ValidSkills.Close

SKILLSTABLE.MoveFirst
SKILLSTABLE.index = "TRIBE"
SKILLSTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME]

LEVEL_10S = 0

RESEARCH_TABLE.index = "SKILL"
RESEARCH_TABLE.MoveFirst

If Not SKILLSTABLE.NoMatch Then
Do While SKILLSTABLE![TRIBE] = Forms![SKILLS_1]![TRIBE NAME]
      If SKILLSTABLE![SKILL LEVEL] >= 10 Then
         LEVEL_10S = LEVEL_10S + 1
         QUERY_STRING = "INSERT INTO AVAILABLE_RESEARCH ( SKILL, TOPIC, [DL REQUIRED],"
         QUERY_STRING = QUERY_STRING & " INFO ) SELECT RESEARCH.SKILL,"
         QUERY_STRING = QUERY_STRING & " RESEARCH.TOPIC, RESEARCH.[DL REQUIRED], RESEARCH.INFO"
         QUERY_STRING = QUERY_STRING & " FROM RESEARCH WHERE (((RESEARCH.SKILL)="
         QUERY_STRING = QUERY_STRING & "'" & SKILLSTABLE![Skill] & "'));"
         Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
         qdfCurrent.Execute
      End If
      SKILLSTABLE.MoveNext
      If SKILLSTABLE.EOF Then
         Exit Do
      ElseIf Not SKILLSTABLE![TRIBE] = Forms![SKILLS_1]![TRIBE NAME] Then
         Exit Do
      End If
Loop
End If

' add into available_research any books that are on hand in the tribe as well plus relevant details.
Set TRIBESBOOKS = TVDBGM.OpenRecordset("TRIBES_BOOKS")
TRIBESBOOKS.index = "PRIMARYKEY"
TRIBESBOOKS.MoveFirst
Do Until TRIBESBOOKS![TRIBE] = Forms![SKILLS_1]![TRIBE NAME]
   TRIBESBOOKS.MoveNext
   If TRIBESBOOKS.EOF Then
      Exit Do
   End If
Loop

BOOK_FOUND = 0

If Not TRIBESBOOKS.EOF Then
Do Until Not TRIBESBOOKS![TRIBE] = Forms![SKILLS_1]![TRIBE NAME]
   'CHECK FOR A LEVEL 10
   Set RESEARCHTABLE = TVDB.OpenRecordset("RESEARCH")
   RESEARCHTABLE.index = "TOPIC"
   RESEARCHTABLE.Seek "=", TRIBESBOOKS![BOOK]
   SKILLSTABLE.index = "PRIMARYKEY"
   SKILLSTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME], RESEARCHTABLE![Skill]
   If Not SKILLSTABLE.NoMatch Then
      If SKILLSTABLE![SKILL LEVEL] = 10 Then
         QUERY_STRING = "INSERT INTO AVAILABLE_RESEARCH ( SKILL, TOPIC, [DL REQUIRED],"
         QUERY_STRING = QUERY_STRING & " INFO ) SELECT RESEARCH.SKILL,"
         QUERY_STRING = QUERY_STRING & " RESEARCH.TOPIC, RESEARCH.[DL REQUIRED], RESEARCH.INFO"
         QUERY_STRING = QUERY_STRING & " FROM RESEARCH WHERE (((RESEARCH.TOPIC)="
         QUERY_STRING = QUERY_STRING & "'" & TRIBESBOOKS![BOOK] & "'));"
         Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
         qdfCurrent.Execute
         
         BOOK_FOUND = BOOK_FOUND + 1
      End If
   End If
   TRIBESBOOKS.MoveNext
   If TRIBESBOOKS.EOF Then
      Exit Do
   End If
Loop
End If

SKILLSTABLE.Close
TRIBESBOOKS.Close
If BOOK_FOUND > 0 Then
   LEVEL_10S = LEVEL_10S + BOOK_FOUND
End If

' delete from available_research table all research achieved to date and all in progress.

Set AVAIL_RES_TABLE = TVDB.OpenRecordset("AVAILABLE_RESEARCH")
AVAIL_RES_TABLE.index = "TOPIC"
If AVAIL_RES_TABLE.EOF Then
   ' NEXT
Else
   AVAIL_RES_TABLE.MoveFirst
End If

Set COMPRESTABLE = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTABLE.index = "TRIBE"
COMPRESTABLE.Seek "=", Forms![SKILLS_1]![TRIBE NAME]

If Not COMPRESTABLE.NoMatch Then
Do
   
   AVAIL_RES_TABLE.Seek "=", COMPRESTABLE![TOPIC]
   If Not AVAIL_RES_TABLE.NoMatch Then
      AVAIL_RES_TABLE.Delete
   End If
   
   COMPRESTABLE.MoveNext
   If COMPRESTABLE.EOF Then
      Exit Do
   End If
   If Not (COMPRESTABLE![TRIBE] = Forms![SKILLS_1]![TRIBE NAME]) Then
      Exit Do
   End If
    
Loop
End If

COMPRESTABLE.Close

Set TRIBE_RESEARCH = TVDBGM.OpenRecordset("TRIBE_RESEARCH")
TRIBE_RESEARCH.index = "SECONDARYKEY"
TRIBE_RESEARCH.Seek "=", Forms![SKILLS_1]![TRIBE NAME]

If Not TRIBE_RESEARCH.NoMatch Then
Do Until Not TRIBE_RESEARCH![TRIBE] = Forms![SKILLS_1]![TRIBE NAME]
   AVAIL_RES_TABLE.Seek "=", TRIBE_RESEARCH![TOPIC]
   If Not AVAIL_RES_TABLE.NoMatch Then
      AVAIL_RES_TABLE.Delete
   End If
   TRIBE_RESEARCH.MoveNext
   If TRIBE_RESEARCH.EOF Then
      Exit Do
   End If
Loop
End If

TRIBE_RESEARCH.Close

AVAIL_RES_TABLE.Close

Set CONSTTABLE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
CONSTTABLE.index = "FORTHKEY"
    CONSTTABLE.Seek "=", MAP_REFERENCE, Forms![SKILLS_1]![CLAN NAME], "UNIVERSITY"

If CONSTTABLE.NoMatch Then
    uni_found = "NO"
Else
    uni_found = "YES"
End If

CONSTTABLE.Close

If LEVEL_10S = 0 Then
   TOPICS = 0
ElseIf LEVEL_10S = 1 Then
   TOPICS = 1
ElseIf uni_found = "YES" Then
    If SEM_FOUND = "YES" Then
       TOPICS = (AVAILABLE_PEOPLE / 200)
    ElseIf LABS_FOUND = "YES" Then
       TOPICS = (AVAILABLE_PEOPLE / 300)
    Else
       TOPICS = (AVAILABLE_PEOPLE / 500)
    End If
Else
    TOPICS = 1
End If
       
CHECK_TOPICS = CLng(TOPICS / 1)
COUNT_TOPICS = TOPICS - CHECK_TOPICS
If COUNT_TOPICS > 0 Then
   TOPICS = CHECK_TOPICS + 1
Else
   TOPICS = CHECK_TOPICS
End If

' CURRENTLY CAN ONLY HAVE 1 RESEARCH TOPIC PER LEVEL 10
' if have 'Multiple Research Topics 2' then can have 2 topics per level 10
If MRT2_FOUND = "YES" Then
   If TOPICS > (LEVEL_10S * 2) Then
      TOPICS = LEVEL_10S * 2
   Else
      TOPICS = LEVEL_10S * 2
   End If
ElseIf MRT_FOUND = "YES" Then
   If TOPICS >= LEVEL_10S Then
      TOPICS = LEVEL_10S + 1
   End If
ElseIf TOPICS > LEVEL_10S Then
   TOPICS = LEVEL_10S
End If
       
'MSG1 = "LEVEL 10'S = " & LEVEL_10S
'MSG2 = "UNI_FOUND = " & UNI_FOUND
'MSG3 = "TOPICS = " & TOPICS
'MsgBox (MSG1 & MSG2 & MSG3)

Forms![SKILLS_1]![RESEARCH TOPIC 40].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 39].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 38].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 37].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 36].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 35].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 34].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 33].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 32].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 31].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 30].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 29].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 28].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 27].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 26].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 25].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 24].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 23].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 22].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 21].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 20].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 19].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 18].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 17].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 16].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 15].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 14].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 13].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 12].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 11].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 10].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 9].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 8].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 7].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 6].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 5].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 4].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 3].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 2].Enabled = True
Forms![SKILLS_1]![RESEARCH TOPIC 1].Enabled = True

i = 40
Do While i > TOPICS
   
   Select Case i
   Case 40
      Forms![SKILLS_1]![RESEARCH TOPIC 40].Enabled = False
   Case 39
      Forms![SKILLS_1]![RESEARCH TOPIC 39].Enabled = False
   Case 38
      Forms![SKILLS_1]![RESEARCH TOPIC 38].Enabled = False
   Case 37
      Forms![SKILLS_1]![RESEARCH TOPIC 37].Enabled = False
   Case 36
      Forms![SKILLS_1]![RESEARCH TOPIC 36].Enabled = False
   Case 35
      Forms![SKILLS_1]![RESEARCH TOPIC 35].Enabled = False
   Case 34
      Forms![SKILLS_1]![RESEARCH TOPIC 34].Enabled = False
   Case 33
      Forms![SKILLS_1]![RESEARCH TOPIC 33].Enabled = False
   Case 32
      Forms![SKILLS_1]![RESEARCH TOPIC 32].Enabled = False
   Case 31
      Forms![SKILLS_1]![RESEARCH TOPIC 31].Enabled = False
   Case 30
      Forms![SKILLS_1]![RESEARCH TOPIC 30].Enabled = False
   Case 29
      Forms![SKILLS_1]![RESEARCH TOPIC 29].Enabled = False
   Case 28
      Forms![SKILLS_1]![RESEARCH TOPIC 28].Enabled = False
   Case 27
      Forms![SKILLS_1]![RESEARCH TOPIC 27].Enabled = False
   Case 26
      Forms![SKILLS_1]![RESEARCH TOPIC 26].Enabled = False
   Case 25
      Forms![SKILLS_1]![RESEARCH TOPIC 25].Enabled = False
   Case 24
      Forms![SKILLS_1]![RESEARCH TOPIC 24].Enabled = False
   Case 23
      Forms![SKILLS_1]![RESEARCH TOPIC 23].Enabled = False
   Case 22
      Forms![SKILLS_1]![RESEARCH TOPIC 22].Enabled = False
   Case 21
      Forms![SKILLS_1]![RESEARCH TOPIC 21].Enabled = False
   Case 20
      Forms![SKILLS_1]![RESEARCH TOPIC 20].Enabled = False
   Case 19
      Forms![SKILLS_1]![RESEARCH TOPIC 19].Enabled = False
   Case 18
      Forms![SKILLS_1]![RESEARCH TOPIC 18].Enabled = False
   Case 17
      Forms![SKILLS_1]![RESEARCH TOPIC 17].Enabled = False
   Case 16
      Forms![SKILLS_1]![RESEARCH TOPIC 16].Enabled = False
   Case 15
      Forms![SKILLS_1]![RESEARCH TOPIC 15].Enabled = False
   Case 14
      Forms![SKILLS_1]![RESEARCH TOPIC 14].Enabled = False
   Case 13
      Forms![SKILLS_1]![RESEARCH TOPIC 13].Enabled = False
   Case 12
      Forms![SKILLS_1]![RESEARCH TOPIC 12].Enabled = False
   Case 11
      Forms![SKILLS_1]![RESEARCH TOPIC 11].Enabled = False
   Case 10
      Forms![SKILLS_1]![RESEARCH TOPIC 10].Enabled = False
   Case 9
      Forms![SKILLS_1]![RESEARCH TOPIC 9].Enabled = False
   Case 8
      Forms![SKILLS_1]![RESEARCH TOPIC 8].Enabled = False
   Case 7
      Forms![SKILLS_1]![RESEARCH TOPIC 7].Enabled = False
   Case 6
      Forms![SKILLS_1]![RESEARCH TOPIC 6].Enabled = False
   Case 5
      Forms![SKILLS_1]![RESEARCH TOPIC 5].Enabled = False
   Case 4
      Forms![SKILLS_1]![RESEARCH TOPIC 4].Enabled = False
   Case 3
      Forms![SKILLS_1]![RESEARCH TOPIC 3].Enabled = False
   Case 2
      Forms![SKILLS_1]![RESEARCH TOPIC 2].Enabled = False
   Case 1
      Forms![SKILLS_1]![RESEARCH TOPIC 1].Enabled = False
   End Select
      
   i = i - 1
Loop

End Function


Public Function Promote_Specialists_From_Training()
Dim Total_in_Training As Long
Dim Total_to_Promote As Long
Dim CLAN As String
Dim PERCENTAGE As Long

If IsNull(Forms![Tribes_Specialists]![TRIBE]) Then
   Exit Function
End If

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set MYFORM = Forms![Tribes_Specialists]

Set TribesSpecialists = TVDBGM.OpenRecordset("Tribes_Specialists")
TribesSpecialists.index = "PRIMARYKEY"
TribesSpecialists.MoveFirst
TribesSpecialists.Seek "=", MYFORM![CLAN], MYFORM![TRIBE], "TRAINING"

Total_in_Training = TribesSpecialists![SPECIALISTS]

' FIND OUT CURRENT SKILL_LEVEL
Set SKILLSTABLE = TVDBGM.OpenRecordset("Skills")
SKILLSTABLE.index = "PRIMARYKEY"
SKILLSTABLE.MoveFirst
SKILLSTABLE.Seek "=", MYFORM![TRIBE], Forms![Tribes_Specialists]![Skill]

PERCENTAGE = SKILLSTABLE![SKILL LEVEL] - 10

Total_to_Promote = CLng((Total_in_Training / 10) * PERCENTAGE)

If Total_to_Promote > 0 Then
   TribesSpecialists.Edit
   TribesSpecialists![SPECIALISTS] = TribesSpecialists![SPECIALISTS] - Total_to_Promote
   TribesSpecialists.UPDATE

   TribesSpecialists.MoveFirst
   TribesSpecialists.Seek "=", MYFORM![CLAN], MYFORM![TRIBE], MYFORM![Skill]

   If TribesSpecialists.NoMatch Then
      TribesSpecialists.AddNew
      TribesSpecialists![CLAN] = MYFORM![CLAN]
      TribesSpecialists![TRIBE] = MYFORM![TRIBE]
      TribesSpecialists![ITEM] = MYFORM![Skill]
      TribesSpecialists![SPECIALISTS] = Total_to_Promote
      TribesSpecialists.UPDATE
   Else
      TribesSpecialists.Edit
      TribesSpecialists![SPECIALISTS] = TribesSpecialists![SPECIALISTS] + Total_to_Promote
      TribesSpecialists.UPDATE
   End If
   
End If

CLAN = MYFORM![CLAN]

DoCmd.Close A_FORM, "Tribes_Specialists"
DoCmd.OpenForm "Tribes_Specialists"
DoCmd.FindRecord CLAN, acEntire

End Function

Public Function OPEN_FORMS(SCREEN_NAME)

   DoCmd.OpenForm SCREEN_NAME
   DoCmd.Maximize

End Function

Public Function Determine_Horses_For_Scouting(Field_Number)
Dim Total_Horses As Long

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set MYFORM = Forms![Tribes_Specialists]

Set TribesGoodsUsage = TVDB.OpenRecordset("Tribes_Goods_Usage")
TribesGoodsUsage.index = "PRIMARYKEY"
TribesGoodsUsage.MoveFirst
TribesGoodsUsage.Seek "=", MYFORM![CLAN], MYFORM![TRIBE], "HORSE"

Set Tribes_Goods = TVDBGM.OpenRecordset("Tribes_Goods")
Tribes_Goods.index = "PRIMARYKEY"
Tribes_Goods.MoveFirst
Tribes_Goods.Seek "=", MYFORM![CLAN], MYFORM![TRIBE], "ANIMAL", "HORSE"

Total_Horses = Tribes_Goods![ITEM_NUMBER]

Forms![SCOUT MOVEMENT]![HORSES1] = Forms![SCOUT MOVEMENT]![SCOUTS1]

TribesGoodsUsage.Edit
TribesGoodsUsage![Number_Used] = TribesGoodsUsage![Number_Used] + Forms![SCOUT MOVEMENT]![SCOUTS1]
TribesGoodsUsage.UPDATE

End Function

Function POPULATE_SKILLS_1_RESEARCH_SKILLS(field)
On Error GoTo POP_SKILLS_ERROR

Dim POSITION As Long
Dim WORDLEN As Long
Dim SEARCHVALUE As String

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")

Set RESEARCH_TABLE = TVDB.OpenRecordset("RESEARCH")
RESEARCH_TABLE.index = "TOPIC"

Select Case field
Case 1
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 1], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 1])
Case 2
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 2], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 2])
Case 3
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 3], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 3])
Case 4
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 4], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 4])
Case 5
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 5], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 5])
Case 6
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 6], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 6])
Case 7
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 7], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 7])
Case 8
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 8], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 8])
Case 9
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 9], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 9])
Case 10
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 10], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 10])
Case 11
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 11], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 11])
Case 12
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 12], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 12])
Case 13
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 13], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 13])
Case 14
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 14], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 14])
Case 15
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 15], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 15])
Case 16
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 16], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 16])
Case 17
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 17], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 17])
Case 18
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 18], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 18])
Case 19
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 19], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 19])
Case 20
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 20], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 20])
Case 21
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 21], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 21])
Case 22
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 22], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 22])
Case 23
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 23], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 23])
Case 24
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 24], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 24])
Case 25
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 25], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 25])
Case 26
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 26], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 26])
Case 27
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 27], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 27])
Case 28
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 28], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 28])
Case 29
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 29], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 29])
Case 30
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 30], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 30])
Case 31
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 31], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 31])
Case 32
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 32], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 32])
Case 33
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 33], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 33])
Case 34
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 34], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 34])
Case 35
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 35], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 35])
Case 36
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 36], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 36])
Case 37
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 37], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 37])
Case 38
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 38], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 38])
Case 39
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 39], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 39])
Case 40
   POSITION = InStr(Forms![SKILLS_1]![RESEARCH TOPIC 40], "(")
   WORDLEN = Len(Forms![SKILLS_1]![RESEARCH TOPIC 40])
Case Else
   '"SHIT"
End Select

If POSITION > 0 Then
   WORDLEN = POSITION - 1
End If

Select Case field
Case 1
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 1], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 1], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL1] = RESEARCH_TABLE![Skill]
Case 2
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 2], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 2], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL2] = RESEARCH_TABLE![Skill]
Case 3
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 3], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 3], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL3] = RESEARCH_TABLE![Skill]
Case 4
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 4], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 4], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL4] = RESEARCH_TABLE![Skill]
Case 5
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 5], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 5], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL5] = RESEARCH_TABLE![Skill]
Case 6
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 6], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 6], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL6] = RESEARCH_TABLE![Skill]
Case 7
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 7], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 7], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL7] = RESEARCH_TABLE![Skill]
Case 8
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 8], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 8], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL8] = RESEARCH_TABLE![Skill]
Case 9
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 9], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 9], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL9] = RESEARCH_TABLE![Skill]
Case 10
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 10], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 10], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL10] = RESEARCH_TABLE![Skill]
Case 11
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 11], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 11], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL11] = RESEARCH_TABLE![Skill]
Case 12
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 12], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 12], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL12] = RESEARCH_TABLE![Skill]
Case 13
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 13], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 13], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL13] = RESEARCH_TABLE![Skill]
Case 14
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 14], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 14], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL14] = RESEARCH_TABLE![Skill]
Case 15
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 15], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 15], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL15] = RESEARCH_TABLE![Skill]
Case 16
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 16], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 16], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL16] = RESEARCH_TABLE![Skill]
Case 17
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 17], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 17], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL17] = RESEARCH_TABLE![Skill]
Case 18
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 18], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 18], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL18] = RESEARCH_TABLE![Skill]
Case 19
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 19], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 19], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL19] = RESEARCH_TABLE![Skill]
Case 20
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 20], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 20], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL20] = RESEARCH_TABLE![Skill]
Case 21
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 21], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 21], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL21] = RESEARCH_TABLE![Skill]
Case 22
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 22], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 2], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL22] = RESEARCH_TABLE![Skill]
Case 23
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 23], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 23], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL23] = RESEARCH_TABLE![Skill]
Case 24
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 24], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 24], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL24] = RESEARCH_TABLE![Skill]
Case 25
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 25], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 25], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL25] = RESEARCH_TABLE![Skill]
Case 26
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 26], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 26], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL26] = RESEARCH_TABLE![Skill]
Case 27
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 27], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 27], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL27] = RESEARCH_TABLE![Skill]
Case 28
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 28], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 28], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL28] = RESEARCH_TABLE![Skill]
Case 29
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 29], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 29], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL29] = RESEARCH_TABLE![Skill]
Case 30
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 30], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 30], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL30] = RESEARCH_TABLE![Skill]
Case 31
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 31], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 31], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL31] = RESEARCH_TABLE![Skill]
Case 32
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 32], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 32], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL32] = RESEARCH_TABLE![Skill]
Case 33
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 33], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 33], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL33] = RESEARCH_TABLE![Skill]
Case 34
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 34], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 34], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL34] = RESEARCH_TABLE![Skill]
Case 35
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 35], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 35], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL35] = RESEARCH_TABLE![Skill]
Case 36
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 36], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 36], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL36] = RESEARCH_TABLE![Skill]
Case 37
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 37], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 37], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL37] = RESEARCH_TABLE![Skill]
Case 38
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 38], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 38], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL38] = RESEARCH_TABLE![Skill]
Case 39
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 39], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 39], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL39] = RESEARCH_TABLE![Skill]
Case 40
   If Right(Mid(Forms![SKILLS_1]![RESEARCH TOPIC 40], 1, WORDLEN), 1) = " " Then
      WORDLEN = WORDLEN - 1
   End If
   SEARCHVALUE = Mid(Forms![SKILLS_1]![RESEARCH TOPIC 40], 1, WORDLEN)

   RESEARCH_TABLE.Seek "=", SEARCHVALUE
   Forms![SKILLS_1]![RESEARCH SKILL40] = RESEARCH_TABLE![Skill]
Case Else
  '"SHIT"
End Select

RESEARCH_TABLE.Close

ERR_close:
   Exit Function

POP_SKILLS_ERROR:
If (Err = 2113) Then
   
   Resume Next
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  Resume ERR_close

End If

End Function

Public Function Clear_Scout_Screen()

   Forms![SCOUT MOVEMENT]![Scout1Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout1Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout2Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout3Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout4Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout5Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout6Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout7Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move01] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move02] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move03] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move04] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move05] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move06] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move07] = "EMPTY"
   Forms![SCOUT MOVEMENT]![Scout8Move08] = "EMPTY"
   Forms![SCOUT MOVEMENT]![SCOUTS1] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS2] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS3] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS4] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS5] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS6] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS7] = 0
   Forms![SCOUT MOVEMENT]![SCOUTS8] = 0
   Forms![SCOUT MOVEMENT]![HORSES1] = 0
   Forms![SCOUT MOVEMENT]![HORSES2] = 0
   Forms![SCOUT MOVEMENT]![HORSES3] = 0
   Forms![SCOUT MOVEMENT]![HORSES4] = 0
   Forms![SCOUT MOVEMENT]![HORSES5] = 0
   Forms![SCOUT MOVEMENT]![HORSES6] = 0
   Forms![SCOUT MOVEMENT]![HORSES7] = 0
   Forms![SCOUT MOVEMENT]![HORSES8] = 0
   Forms![SCOUT MOVEMENT]![Elephants1] = 0
   Forms![SCOUT MOVEMENT]![Elephants2] = 0
   Forms![SCOUT MOVEMENT]![Elephants3] = 0
   Forms![SCOUT MOVEMENT]![Elephants4] = 0
   Forms![SCOUT MOVEMENT]![Elephants5] = 0
   Forms![SCOUT MOVEMENT]![Elephants6] = 0
   Forms![SCOUT MOVEMENT]![Elephants7] = 0
   Forms![SCOUT MOVEMENT]![Elephants8] = 0
   Forms![SCOUT MOVEMENT]![Camels1] = 0
   Forms![SCOUT MOVEMENT]![Camels2] = 0
   Forms![SCOUT MOVEMENT]![Camels3] = 0
   Forms![SCOUT MOVEMENT]![Camels4] = 0
   Forms![SCOUT MOVEMENT]![Camels5] = 0
   Forms![SCOUT MOVEMENT]![Camels6] = 0
   Forms![SCOUT MOVEMENT]![Camels7] = 0
   Forms![SCOUT MOVEMENT]![Camels8] = 0
   Forms![SCOUT MOVEMENT]![MISSION1] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION2] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION3] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION4] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION5] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION6] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION7] = "NONE"
   Forms![SCOUT MOVEMENT]![MISSION8] = "NONE"
   Forms![SCOUT MOVEMENT]![Scout1Move01].SetFocus

End Function

Public Function Populate_Land_Combat_Information(TRIBE As Integer)

Dim FORM_FIELD As String
Dim Com_clan As String
Dim Com_tribe As String

If TRIBE = 1 Then
   If IsNull(Forms![Land_Combat_Information]![1st Tribe]) Then
      Exit Function
   Else
      Com_clan = Forms![Land_Combat_Information]![1st Clan]
      Com_tribe = Forms![Land_Combat_Information]![1st Tribe]
   End If
Else
   If IsNull(Forms![Land_Combat_Information]![2nd Tribe]) Then
      Exit Function
   Else
      Com_clan = Forms![Land_Combat_Information]![2nd Clan]
      Com_tribe = Forms![Land_Combat_Information]![2nd Tribe]
   End If
End If

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILETV = GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILETV, False, False)

Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "PRIMARYKEY"

SKILLSTABLE.Seek "=", Com_tribe, "ARCHERY"
If SKILLSTABLE.NoMatch Then
   ARCHERY_LEVEL = 0
Else
   ARCHERY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "COMBAT"
If SKILLSTABLE.NoMatch Then
   COMBAT_LEVEL = 0
Else
   COMBAT_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "HEALING"
If SKILLSTABLE.NoMatch Then
   HEALING_LEVEL = 0
Else
   HEALING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "HEAVY WEAPONS"
If SKILLSTABLE.NoMatch Then
   HVYWPNS_LEVEL = 0
Else
   HVYWPNS_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "HORSEMANSHIP"
If SKILLSTABLE.NoMatch Then
   HORSEMANSHIP_LEVEL = 0
Else
   HORSEMANSHIP_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "LEADERSHIP"
If SKILLSTABLE.NoMatch Then
   LEADERSHIP_LEVEL = 0
Else
   LEADERSHIP_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "SCOUTING"
If SKILLSTABLE.NoMatch Then
   SCOUTING_LEVEL = 0
Else
   SCOUTING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "SECURITY"
If SKILLSTABLE.NoMatch Then
   SECURITY_LEVEL = 0
Else
   SECURITY_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "SPYING"
If SKILLSTABLE.NoMatch Then
   SPYING_LEVEL = 0
Else
   SPYING_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "TACTICS"
If SKILLSTABLE.NoMatch Then
   TACTICS_LEVEL = 0
Else
   TACTICS_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If
SKILLSTABLE.Seek "=", Com_tribe, "TORTURE"
If SKILLSTABLE.NoMatch Then
   TORTURE_LEVEL = 0
Else
   TORTURE_LEVEL = SKILLSTABLE![SKILL LEVEL]
End If

SKILLSTABLE.Close

If TRIBE = 1 Then
   Forms![Land_Combat_Information]![1st Archery] = ARCHERY_LEVEL
   Forms![Land_Combat_Information]![1st Combat] = COMBAT_LEVEL
   Forms![Land_Combat_Information]![1st Healing] = HEALING_LEVEL
   Forms![Land_Combat_Information]![1st Heavy Weapons] = HVYWPNS_LEVEL
   Forms![Land_Combat_Information]![1st Horsemanship] = HORSEMANSHIP_LEVEL
   Forms![Land_Combat_Information]![1st Leadership] = LEADERSHIP_LEVEL
   Forms![Land_Combat_Information]![1st Scouting] = SCOUTING_LEVEL
   Forms![Land_Combat_Information]![1st Security] = SECURITY_LEVEL
   Forms![Land_Combat_Information]![1st Spying] = SPYING_LEVEL
   Forms![Land_Combat_Information]![1st Tactics] = TACTICS_LEVEL
   Forms![Land_Combat_Information]![1st Torture] = TORTURE_LEVEL
Else
   Forms![Land_Combat_Information]![2nd Archery] = ARCHERY_LEVEL
   Forms![Land_Combat_Information]![2nd Combat] = COMBAT_LEVEL
   Forms![Land_Combat_Information]![2nd Healing] = HEALING_LEVEL
   Forms![Land_Combat_Information]![2nd Heavy Weapons] = HVYWPNS_LEVEL
   Forms![Land_Combat_Information]![2nd Horsemanship] = HORSEMANSHIP_LEVEL
   Forms![Land_Combat_Information]![2nd Leadership] = LEADERSHIP_LEVEL
   Forms![Land_Combat_Information]![2nd Scouting] = SCOUTING_LEVEL
   Forms![Land_Combat_Information]![2nd Security] = SECURITY_LEVEL
   Forms![Land_Combat_Information]![2nd Spying] = SPYING_LEVEL
   Forms![Land_Combat_Information]![2nd Tactics] = TACTICS_LEVEL
   Forms![Land_Combat_Information]![2nd Torture] = TORTURE_LEVEL
End If


End Function
Function Scouting_Losses_TribeNumber_Exit()

CLAN = Forms![SCOUTING_LOSSES]![CLANNUMBER]
TRIBE = Forms![SCOUTING_LOSSES]![TRIBENUMBER]

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.MoveFirst
SKILLSTABLE.index = "PRIMARYKEY"
SKILLSTABLE.Seek "=", TRIBE, "SCOUTING"

If Not SKILLSTABLE.NoMatch Then
   Forms![SCOUTING_LOSSES]![SCOUTING LEVEL] = SKILLSTABLE![SKILL LEVEL]
Else
   Forms![SCOUTING_LOSSES]![SCOUTING LEVEL] = 0
End If

SKILLSTABLE.Close

End Function


Public Function Refresh_Screen(SCREEN_NAME)

DoCmd.Close acForm, SCREEN_NAME
DoCmd.OpenForm SCREEN_NAME
DoCmd.Maximize

End Function


Public Function Go_To_Field(FIELD_NAME)
   DoCmd.GoToControl (FIELD_NAME)
End Function

Public Function Turn_Activities_Tribenumber_Exit()
On Error GoTo ERR_Turn_Activities_Tribenumber_Exit
Dim hex_pop As Long

If IsNull(Forms![TURNS ACTIVITIES]![TRIBENUMBER]) Then
   Exit Function
End If

TRIBE = Forms![TURNS ACTIVITIES]![TRIBENUMBER]
CLAN = "0" & Mid(TRIBE, 2, 3)

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set TRIBEINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBEINFO.MoveFirst
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.Seek "=", CLAN, TRIBE

If Not IsNull(TRIBEINFO![GOODS TRIBE]) Then
   GOODS_TRIBE = TRIBEINFO![GOODS TRIBE]
Else
   GOODS_TRIBE = TRIBE
End If

Forms![TURNS ACTIVITIES]![GOODS_TRIBE] = GOODS_TRIBE
TSpecialists = 0

CURRENT_HEX = TRIBEINFO![Current Hex]
TActivesAvailable = TRIBEINFO![WARRIORS] + TRIBEINFO![ACTIVES] + TRIBEINFO![SLAVE]
TActivesAvailable = TActivesAvailable + TRIBEINFO![HIRELINGS]
Set TribesSpecialists = TVDBGM.OpenRecordset("TRIBES_SPECIALISTS")
TribesSpecialists.index = "PRIMARYKEY"
  If TribesSpecialists.BOF Then
      ' do nothing
  Else
      TribesSpecialists.MoveFirst
  End If
TribesSpecialists.Seek "=", CLAN, TRIBE, "TRAINING"

If TribesSpecialists.NoMatch Then
    'DO NOTHING
Else
     TActivesAvailable = TActivesAvailable - TribesSpecialists![SPECIALISTS]
End If
TribesSpecialists.Close

Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
If Not HEXMAPMINERALS.EOF Then
   HEXMAPMINERALS.MoveFirst
End If
HEXMAPMINERALS.index = "PRIMARYKEY"
HEXMAPMINERALS.Seek "=", CURRENT_HEX

Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
If Not HEXMAPCONST.EOF Then
   HEXMAPCONST.MoveFirst
End If
HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.Seek "=", CURRENT_HEX, CLAN, TRIBE, "TRADING POST"

If Not HEXMAPCONST.EOF Then
   HEXMAPCONST.MoveFirst
End If
HEXMAPCONST.index = "PRIMARYKEY"
HEXMAPCONST.Seek "=", CURRENT_HEX, CLAN, TRIBE, "TRADING POST"

Forms![TURNS ACTIVITIES]![Current Hex] = CURRENT_HEX
Forms![TURNS ACTIVITIES]![CURRENT TERRAIN] = TRIBEINFO![CURRENT TERRAIN]
If HEXMAPMINERALS.NoMatch Then
     Forms![TURNS ACTIVITIES]![Current_Mineral] = "None"
     Forms![TURNS ACTIVITIES]![2nd_Mineral] = "None"
Else
     Forms![TURNS ACTIVITIES]![Current_Mineral] = HEXMAPMINERALS![ORE_TYPE]
     Forms![TURNS ACTIVITIES]![2nd_Mineral] = HEXMAPMINERALS![SECOND_ORE]
End If
Forms![TURNS ACTIVITIES]![Available_Actives] = TActivesAvailable
Forms![TURNS ACTIVITIES]![Available_Slaves] = TRIBEINFO![SLAVE]

If Not HEXMAPCONST.NoMatch Then
   If HEXMAPCONST![1] > 0 Then
      Forms![TURNS ACTIVITIES]![TP] = "Y"
   Else
      Forms![TURNS ACTIVITIES]![TP] = "N"
   End If
End If

'GET HEXMAP POPULATION

hex_pop = HEX_POPULATION(CLAN, TRIBE, CURRENT_HEX)

Forms![TURNS ACTIVITIES]![HEX_MAP_POP] = hex_pop


HEXMAPMINERALS.Close
HEXMAPCONST.Close

ERR_Turn_Activities_Tribenumber_Exit_CLOSE:
   Exit Function

ERR_Turn_Activities_Tribenumber_Exit:
If (Err = 3420) Then
   Resume Next
   
Else
  Resume ERR_Turn_Activities_Tribenumber_Exit_CLOSE
End If


End Function
Public Function Update_Available_Actives(ACTIVE_NUMBER)

TActivesAvailable = Forms![TURNS ACTIVITIES]![Available_Actives]

If Forms![TURNS ACTIVITIES]![ACTIVITY01] = "KILLING" Then
   Exit Function
ElseIf Forms![TURNS ACTIVITIES]![ACTIVITY01] = "CONVERT" Then
   Exit Function
ElseIf Forms![TURNS ACTIVITIES]![ACTIVITY01] = "EATING" Then
     If Forms![TURNS ACTIVITIES]![item01] = "PROVS" Then
        Exit Function
     End If
ElseIf Forms![TURNS ACTIVITIES]![ACTIVITY01] = "SPECIALISTS" Then
     If Forms![TURNS ACTIVITIES]![item01] = "PROMOTION" Then
         Exit Function
     End If
End If

If ACTIVE_NUMBER = "ACTIVES" Then
   If Not IsNull(Forms![TURNS ACTIVITIES]![ACTIVES]) Then
      TActivesAvailable = TActivesAvailable - Forms![TURNS ACTIVITIES]![ACTIVES]
   End If
End If

Forms![TURNS ACTIVITIES]![Available_Actives] = TActivesAvailable

End Function



'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 28/01/96  *  Add function for updating Activities Screen                    **'
'** 07/11/96  *  Fix end of table bugs                                          **'
'*===============================================================================*'
 

Function POPULATE_ACTIVITIES_SCREEN()
ReDim Description(9) As String
ReDim NUMBER(9) As Long
Dim COUNTER As Long
Dim ACTIVITY As String
Dim ITEM As String
Dim TYPE_OF_ACTIVITY As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![Activities]

Set ActivitiesTable = TVDB.OpenRecordset("ACTIVITIES")
ActivitiesTable.index = "primarykey"
ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", MYFORM![ACTIVITY TYPE], MYFORM![ITEM TYPE], MYFORM![TYPE]

Set activitytable = TVDB.OpenRecordset("ACTIVITY")
activitytable.index = "SECONDARYKEY"
activitytable.MoveFirst
activitytable.Seek "=", MYFORM![ACTIVITY TYPE], MYFORM![ITEM TYPE], MYFORM![TYPE]

ACTIVITY = MYFORM![ACTIVITY TYPE]
ITEM = MYFORM![ITEM TYPE]
TYPE_OF_ACTIVITY = MYFORM![TYPE]

For COUNTER = 1 To 9
   Description(COUNTER) = ""
   NUMBER(COUNTER) = 0
Next

COUNTER = 1

If Not activitytable.NoMatch Then
   Do While ((activitytable![ACTIVITY] = ACTIVITY) And (activitytable![ITEM] = ITEM) And (activitytable![TYPE] = TYPE_OF_ACTIVITY))
      Description(COUNTER) = activitytable![GOOD]
      NUMBER(COUNTER) = activitytable![NUMBER]
      COUNTER = COUNTER + 1
 
      activitytable.MoveNext

      If activitytable.EOF Then
         Exit Do
      End If

   Loop
End If

If Not ActivitiesTable.NoMatch Then
   Forms![Activities]![SHORTNAME] = ActivitiesTable![SHORTNAME]
   Forms![Activities]![SKILL LEVEL] = ActivitiesTable![SKILL LEVEL]
   Forms![Activities]![SECOND SKILL] = ActivitiesTable![SECOND SKILL]
   Forms![Activities]![SECOND SKILL LEVEL] = ActivitiesTable![SECOND SKILL LEVEL]
   Forms![Activities]![THIRD SKILL] = ActivitiesTable![THIRD SKILL]
   Forms![Activities]![THIRD SKILL LEVEL] = ActivitiesTable![THIRD SKILL LEVEL]
   Forms![Activities]![FORTH SKILL] = ActivitiesTable![FORTH SKILL]
   Forms![Activities]![FORTH SKILL LEVEL] = ActivitiesTable![FORTH SKILL LEVEL]
   Forms![Activities]![NUMBER OF ITEMS] = ActivitiesTable![NUMBER OF ITEMS]
   Forms![Activities]![PEOPLE] = ActivitiesTable![PEOPLE]
   Forms![Activities]![research] = ActivitiesTable![research]
   Forms![Activities]![GOODS_USED] = ActivitiesTable![GOODS_USED]
   Forms![Activities]![GOOD_PRODUCED] = ActivitiesTable![GOOD_PRODUCED]
Else
   Forms![Activities]![SHORTNAME] = " "
   Forms![Activities]![SKILL LEVEL] = 0
   Forms![Activities]![SECOND SKILL] = " "
   Forms![Activities]![SECOND SKILL LEVEL] = 0
   Forms![Activities]![THIRD SKILL] = " "
   Forms![Activities]![THIRD SKILL LEVEL] = 0
   Forms![Activities]![FORTH SKILL] = " "
   Forms![Activities]![FORTH SKILL LEVEL] = 0
   Forms![Activities]![NUMBER OF ITEMS] = 0
   Forms![Activities]![PEOPLE] = 0
   Forms![Activities]![research] = "N"
   Forms![Activities]![GOODS_USED] = "N"
   Forms![Activities]![GOOD_PRODUCED] = "N"

End If

  Forms![Activities]![DESCRIPTION 1] = Description(1)
  Forms![Activities]![VALUE 1] = NUMBER(1)
  Forms![Activities]![DESCRIPTION 2] = Description(2)
  Forms![Activities]![VALUE 2] = NUMBER(2)
  Forms![Activities]![DESCRIPTION 3] = Description(3)
  Forms![Activities]![VALUE 3] = NUMBER(3)
  Forms![Activities]![DESCRIPTION 4] = Description(4)
  Forms![Activities]![VALUE 4] = NUMBER(4)
  Forms![Activities]![DESCRIPTION 5] = Description(5)
  Forms![Activities]![VALUE 5] = NUMBER(5)
  Forms![Activities]![DESCRIPTION 6] = Description(6)
  Forms![Activities]![VALUE 6] = NUMBER(6)
  Forms![Activities]![DESCRIPTION 7] = Description(7)
  Forms![Activities]![VALUE 7] = NUMBER(7)
  Forms![Activities]![DESCRIPTION 8] = Description(8)
  Forms![Activities]![VALUE 8] = NUMBER(8)
  Forms![Activities]![DESCRIPTION 9] = Description(9)
  Forms![Activities]![VALUE 9] = NUMBER(9)


End Function


Function POPULATE_HEX_MAP()
Dim FEATURE_1 As String
Dim FEATURE_2 As String
Dim FEATURE_3 As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![HEX_MAP]

FEATURE_1 = "NONE"
FEATURE_2 = "NONE"
FEATURE_3 = "NONE"

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
Else
   If Mid(hexmaptable![Borders], 1, 2) = "BE" Then
      MYFORM![North_Border] = "Beach"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "BR" Then
      MYFORM![North_Border] = "Bridge"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "CA" Then
      MYFORM![North_Border] = "Canal"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "CL" Then
      MYFORM![North_Border] = "Cliff"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "FO" Then
      MYFORM![North_Border] = "Ford"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "PA" Then
      MYFORM![North_Border] = "Pass"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "RE" Then
      MYFORM![North_Border] = "Reef"
   ElseIf Mid(hexmaptable![Borders], 1, 2) = "RI" Then
      MYFORM![North_Border] = "River"
   Else
      MYFORM![North_Border] = "None"
   End If
   If Mid(hexmaptable![Borders], 3, 2) = "BE" Then
      MYFORM![North_East_Border] = "Beach"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "BR" Then
      MYFORM![North_East_Border] = "Bridge"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "CA" Then
      MYFORM![North_East_Border] = "Canal"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "CL" Then
      MYFORM![North_East_Border] = "Cliff"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "FO" Then
      MYFORM![North_East_Border] = "Ford"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "PA" Then
      MYFORM![North_East_Border] = "Pass"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "RE" Then
      MYFORM![North_East_Border] = "Reef"
   ElseIf Mid(hexmaptable![Borders], 3, 2) = "RI" Then
      MYFORM![North_East_Border] = "River"
   Else
      MYFORM![North_East_Border] = "None"
   End If
   If Mid(hexmaptable![Borders], 5, 2) = "BE" Then
      MYFORM![South_East_Border] = "Beach"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "BR" Then
      MYFORM![South_East_Border] = "Bridge"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "CA" Then
      MYFORM![South_East_Border] = "Canal"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "CL" Then
      MYFORM![South_East_Border] = "Cliff"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "FO" Then
      MYFORM![South_East_Border] = "Ford"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "PA" Then
      MYFORM![South_East_Border] = "Pass"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "RE" Then
      MYFORM![South_East_Border] = "Reef"
   ElseIf Mid(hexmaptable![Borders], 5, 2) = "RI" Then
      MYFORM![South_East_Border] = "River"
   Else
      MYFORM![South_East_Border] = "None"
   End If
   If Mid(hexmaptable![Borders], 7, 2) = "BE" Then
      MYFORM![South_Border] = "Beach"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "BR" Then
      MYFORM![South_Border] = "Bridge"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "CA" Then
      MYFORM![South_Border] = "Canal"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "CL" Then
      MYFORM![South_Border] = "Cliff"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "FO" Then
      MYFORM![South_Border] = "Ford"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "PA" Then
      MYFORM![South_Border] = "Pass"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "RE" Then
      MYFORM![South_Border] = "Reef"
   ElseIf Mid(hexmaptable![Borders], 7, 2) = "RI" Then
      MYFORM![South_Border] = "River"
   Else
      MYFORM![South_Border] = "None"
   End If
   If Mid(hexmaptable![Borders], 9, 2) = "BE" Then
      MYFORM![South_West_Border] = "Beach"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "BR" Then
      MYFORM![South_West_Border] = "Bridge"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "CA" Then
      MYFORM![South_West_Border] = "Canal"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "CL" Then
      MYFORM![South_West_Border] = "Cliff"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "FO" Then
      MYFORM![South_West_Border] = "Ford"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "PA" Then
      MYFORM![South_West_Border] = "Pass"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "RE" Then
      MYFORM![South_West_Border] = "Reef"
   ElseIf Mid(hexmaptable![Borders], 9, 2) = "RI" Then
      MYFORM![South_West_Border] = "River"
   Else
      MYFORM![South_West_Border] = "None"
   End If
   If Mid(hexmaptable![Borders], 11, 2) = "BE" Then
      MYFORM![North_West_Border] = "Beach"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "BR" Then
      MYFORM![North_West_Border] = "Bridge"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "CA" Then
      MYFORM![North_West_Border] = "Canal"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "CL" Then
      MYFORM![North_West_Border] = "Cliff"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "FO" Then
      MYFORM![North_West_Border] = "Ford"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "PA" Then
      MYFORM![North_West_Border] = "Pass"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "RE" Then
      MYFORM![North_West_Border] = "Reef"
   ElseIf Mid(hexmaptable![Borders], 11, 2) = "RI" Then
      MYFORM![North_West_Border] = "River"
   Else
      MYFORM![North_West_Border] = "None"
   End If
   
   MYFORM![ROAD(N)] = Mid(hexmaptable![ROADS], 1, 1)
   MYFORM![ROAD(NE)] = Mid(hexmaptable![ROADS], 2, 1)
   MYFORM![ROAD(SE)] = Mid(hexmaptable![ROADS], 3, 1)
   MYFORM![ROAD(S)] = Mid(hexmaptable![ROADS], 4, 1)
   MYFORM![ROAD(SW)] = Mid(hexmaptable![ROADS], 5, 1)
   MYFORM![ROAD(NW)] = Mid(hexmaptable![ROADS], 6, 1)

If hexmaptable![QUARRYING] = "Y" Then
    FEATURE_1 = "QUARRY"
End If
If hexmaptable![SPRINGS] = "Y" Then
    If FEATURE_1 = "NONE" Then
        FEATURE_1 = "SPRING"
    ElseIf FEATURE_2 = "NONE" Then
        FEATURE_2 = "SPRING"
    ElseIf FEATURE_3 = "NONE" Then
        FEATURE_3 = "SPRING"
    End If
End If
If hexmaptable![SALMON RUN] = "Y" Then
    If FEATURE_1 = "NONE" Then
        FEATURE_1 = "SALMON"
    ElseIf FEATURE_2 = "NONE" Then
        FEATURE_2 = "SALMON"
    ElseIf FEATURE_3 = "NONE" Then
        FEATURE_3 = "SALMON"
    End If
End If
If hexmaptable![FISH AREA] = "Y" Then
    If FEATURE_1 = "NONE" Then
        FEATURE_1 = "FISH"
    ElseIf FEATURE_2 = "NONE" Then
        FEATURE_2 = "FISH"
    ElseIf FEATURE_3 = "NONE" Then
        FEATURE_3 = "FISH"
    End If
End If
If hexmaptable![WHALE AREA] = "Y" Then
    If FEATURE_1 = "NONE" Then
        FEATURE_1 = "WHALE"
    ElseIf FEATURE_2 = "NONE" Then
        FEATURE_2 = "WHALE"
    ElseIf FEATURE_3 = "NONE" Then
        FEATURE_3 = "WHALE"
    End If
End If
End If
MYFORM![Feature_One] = FEATURE_1
MYFORM![Feature_Two] = FEATURE_2
MYFORM![Feature_Three] = FEATURE_3

Set HEXMAPCITY = TVDBGM.OpenRecordset("HEX_MAP_CITY")
HEXMAPCITY.index = "PRIMARYKEY"
HEXMAPCITY.MoveFirst
HEXMAPCITY.Seek "=", MYFORM![MAP]

If HEXMAPCITY.NoMatch Then
   MYFORM![CITY] = Null
Else
   MYFORM![CITY] = HEXMAPCITY![CITY]
   End If

Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
HEXMAPMINERALS.index = "PRIMARYKEY"
HEXMAPMINERALS.MoveFirst
HEXMAPMINERALS.Seek "=", MYFORM![MAP]

If HEXMAPMINERALS.NoMatch Then
   MYFORM![ORE TYPE] = Null
   MYFORM![SECOND ORE] = Null
   MYFORM![THIRD ORE] = Null
Else
   MYFORM![ORE TYPE] = HEXMAPMINERALS![ORE_TYPE]
   MYFORM![SECOND ORE] = HEXMAPMINERALS![SECOND_ORE]
   MYFORM![THIRD ORE] = HEXMAPMINERALS![THIRD_ORE]
End If

hexmaptable.Close
HEXMAPCITY.Close
HEXMAPMINERALS.Close

End Function


Public Function POPULATE_TRIBES_GOODS_SCREEN()
On Error GoTo ERR_POP_GOODS_SCREEN

ReDim Description(60) As String
ReDim NUMBER(60) As Double
Dim ITEM_TYPE As String
Dim ITEM_TYPE_NF As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
'IDENTIFY FORM WITH FOCUS

Set MYFORM = SCREEN.ActiveForm
  
ITEM_TYPE = MYFORM![ITEM_TYPE]
MYFORM![ITEM_TYPE].Visible = False
ITEM_TYPE_NF = "NO"

If ITEM_TYPE = "MODIFIERS" Then
   Set TribesModifiers = TVDBGM.OpenRecordset("MODIFIERS")
   TribesModifiers.index = "TRIBE"
   TribesModifiers.MoveFirst
   TribesModifiers.Seek "=", MYFORM![TRIBENUMBER]
 
   For COUNTER = 1 To 60
      Description(COUNTER) = ""
      NUMBER(COUNTER) = 0
   Next

   COUNTER = 1
  
   If Not TribesModifiers.NoMatch Then
      Do While TribesModifiers![TRIBE] = MYFORM![TRIBENUMBER]
         Description(COUNTER) = TribesModifiers![Modifier]
         NUMBER(COUNTER) = TribesModifiers![AMOUNT]
         COUNTER = COUNTER + 1

         TribesModifiers.MoveNext

         If TribesModifiers.EOF Then
            Exit Do
         End If
         If Not TribesModifiers![TRIBE] = MYFORM![TRIBENUMBER] Then
            Exit Do
         End If
      Loop
   End If

Else
   Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_GOODS")
   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.MoveFirst
   
 
   For COUNTER = 1 To 60
      Description(COUNTER) = ""
      NUMBER(COUNTER) = 0
   Next

   COUNTER = 1
   
   Do While Not TRIBESGOODS![CLAN] = MYFORM![CLANNUMBER]
      TRIBESGOODS.MoveNext
      If TRIBESGOODS.EOF Then
         Exit Function
      End If
   Loop
   
   Do While Not TRIBESGOODS![TRIBE] = MYFORM![TRIBENUMBER]
      TRIBESGOODS.MoveNext
      If TRIBESGOODS.EOF Then
         ITEM_TYPE_NF = "YES"
         Exit Do
      End If
   Loop
   
   If ITEM_TYPE_NF = "NO" Then
      Do While Not TRIBESGOODS![ITEM_TYPE] = MYFORM![ITEM_TYPE]
         TRIBESGOODS.MoveNext
         If TRIBESGOODS.EOF Then
            ITEM_TYPE_NF = "YES"
            Exit Do
         End If
      Loop
   End If
   
   If ITEM_TYPE_NF = "NO" Then
      Do While TRIBESGOODS![TRIBE] = MYFORM![TRIBENUMBER] And TRIBESGOODS![ITEM_TYPE] = ITEM_TYPE
         Description(COUNTER) = TRIBESGOODS![ITEM]
         NUMBER(COUNTER) = TRIBESGOODS![ITEM_NUMBER]
         COUNTER = COUNTER + 1
         
         TRIBESGOODS.MoveNext

         If TRIBESGOODS.EOF Then
            Exit Do
         End If
         If Not TRIBESGOODS![TRIBE] = MYFORM![TRIBENUMBER] Then
            Exit Do
         End If
         If Not TRIBESGOODS![ITEM_TYPE] = ITEM_TYPE Then
            Exit Do
         End If
      Loop
   End If
End If
  
MYFORM![DESCRIPTION 1] = Description(1)
MYFORM![VALUE 1] = NUMBER(1)
MYFORM![DESCRIPTION 2] = Description(2)
MYFORM![VALUE 2] = NUMBER(2)
MYFORM![DESCRIPTION 3] = Description(3)
MYFORM![VALUE 3] = NUMBER(3)
MYFORM![DESCRIPTION 4] = Description(4)
MYFORM![VALUE 4] = NUMBER(4)
MYFORM![DESCRIPTION 5] = Description(5)
MYFORM![VALUE 5] = NUMBER(5)
MYFORM![DESCRIPTION 6] = Description(6)
MYFORM![VALUE 6] = NUMBER(6)
MYFORM![DESCRIPTION 7] = Description(7)
MYFORM![VALUE 7] = NUMBER(7)
MYFORM![DESCRIPTION 8] = Description(8)
MYFORM![VALUE 8] = NUMBER(8)
MYFORM![DESCRIPTION 9] = Description(9)
MYFORM![VALUE 9] = NUMBER(9)
MYFORM![DESCRIPTION 10] = Description(10)
MYFORM![VALUE 10] = NUMBER(10)
MYFORM![DESCRIPTION 11] = Description(11)
MYFORM![VALUE 11] = NUMBER(11)
MYFORM![DESCRIPTION 12] = Description(12)
MYFORM![VALUE 12] = NUMBER(12)
MYFORM![DESCRIPTION 13] = Description(13)
MYFORM![VALUE 13] = NUMBER(13)
MYFORM![DESCRIPTION 14] = Description(14)
MYFORM![VALUE 14] = NUMBER(14)
MYFORM![DESCRIPTION 15] = Description(15)
MYFORM![VALUE 15] = NUMBER(15)
MYFORM![DESCRIPTION 16] = Description(16)
MYFORM![VALUE 16] = NUMBER(16)
MYFORM![DESCRIPTION 17] = Description(17)
MYFORM![VALUE 17] = NUMBER(17)
MYFORM![DESCRIPTION 18] = Description(18)
MYFORM![VALUE 18] = NUMBER(18)
MYFORM![DESCRIPTION 19] = Description(19)
MYFORM![VALUE 19] = NUMBER(19)
MYFORM![DESCRIPTION 20] = Description(20)
MYFORM![VALUE 20] = NUMBER(20)
MYFORM![DESCRIPTION 21] = Description(21)
MYFORM![VALUE 21] = NUMBER(21)
MYFORM![DESCRIPTION 22] = Description(22)
MYFORM![VALUE 22] = NUMBER(22)
MYFORM![DESCRIPTION 23] = Description(23)
MYFORM![VALUE 23] = NUMBER(23)
MYFORM![DESCRIPTION 24] = Description(24)
MYFORM![VALUE 24] = NUMBER(24)
MYFORM![DESCRIPTION 25] = Description(25)
MYFORM![VALUE 25] = NUMBER(25)
MYFORM![DESCRIPTION 26] = Description(26)
MYFORM![VALUE 26] = NUMBER(26)
MYFORM![DESCRIPTION 27] = Description(27)
MYFORM![VALUE 27] = NUMBER(27)
MYFORM![DESCRIPTION 28] = Description(28)
MYFORM![VALUE 28] = NUMBER(28)
MYFORM![DESCRIPTION 29] = Description(29)
MYFORM![VALUE 29] = NUMBER(29)
MYFORM![DESCRIPTION 30] = Description(30)
MYFORM![VALUE 30] = NUMBER(30)
MYFORM![DESCRIPTION 31] = Description(31)
MYFORM![VALUE 31] = NUMBER(31)
MYFORM![DESCRIPTION 32] = Description(32)
MYFORM![VALUE 32] = NUMBER(32)
MYFORM![DESCRIPTION 33] = Description(33)
MYFORM![VALUE 33] = NUMBER(33)
MYFORM![DESCRIPTION 34] = Description(34)
MYFORM![VALUE 34] = NUMBER(34)
MYFORM![DESCRIPTION 35] = Description(35)
MYFORM![VALUE 35] = NUMBER(35)
MYFORM![DESCRIPTION 36] = Description(36)
MYFORM![VALUE 36] = NUMBER(36)
MYFORM![DESCRIPTION 37] = Description(37)
MYFORM![VALUE 37] = NUMBER(37)
MYFORM![DESCRIPTION 38] = Description(38)
MYFORM![VALUE 38] = NUMBER(38)
MYFORM![DESCRIPTION 39] = Description(39)
MYFORM![VALUE 39] = NUMBER(39)
MYFORM![DESCRIPTION 40] = Description(40)
MYFORM![VALUE 40] = NUMBER(40)
MYFORM![DESCRIPTION 41] = Description(41)
MYFORM![VALUE 41] = NUMBER(41)
MYFORM![DESCRIPTION 42] = Description(42)
MYFORM![VALUE 42] = NUMBER(42)
MYFORM![DESCRIPTION 43] = Description(43)
MYFORM![VALUE 43] = NUMBER(43)
MYFORM![DESCRIPTION 44] = Description(44)
MYFORM![VALUE 44] = NUMBER(44)
MYFORM![DESCRIPTION 45] = Description(45)
MYFORM![VALUE 45] = NUMBER(45)
MYFORM![DESCRIPTION 46] = Description(46)
MYFORM![VALUE 46] = NUMBER(46)
MYFORM![DESCRIPTION 47] = Description(47)
MYFORM![VALUE 47] = NUMBER(47)
MYFORM![DESCRIPTION 48] = Description(48)
MYFORM![VALUE 48] = NUMBER(48)
MYFORM![DESCRIPTION 49] = Description(49)
MYFORM![VALUE 49] = NUMBER(49)
MYFORM![DESCRIPTION 50] = Description(50)
MYFORM![VALUE 50] = NUMBER(50)
MYFORM![DESCRIPTION 51] = Description(51)
MYFORM![VALUE 51] = NUMBER(51)
MYFORM![DESCRIPTION 52] = Description(52)
MYFORM![VALUE 52] = NUMBER(52)
MYFORM![DESCRIPTION 53] = Description(53)
MYFORM![VALUE 53] = NUMBER(53)
MYFORM![DESCRIPTION 54] = Description(54)
MYFORM![VALUE 54] = NUMBER(54)
MYFORM![DESCRIPTION 55] = Description(55)
MYFORM![VALUE 55] = NUMBER(55)
MYFORM![DESCRIPTION 56] = Description(56)
MYFORM![VALUE 56] = NUMBER(56)
MYFORM![DESCRIPTION 57] = Description(57)
MYFORM![VALUE 57] = NUMBER(57)
MYFORM![DESCRIPTION 58] = Description(58)
MYFORM![VALUE 58] = NUMBER(58)
MYFORM![DESCRIPTION 59] = Description(59)
MYFORM![VALUE 59] = NUMBER(59)
MYFORM![DESCRIPTION 60] = Description(60)
MYFORM![VALUE 60] = NUMBER(60)

ERR_POP_GOODS_SCREEN_CLOSE:
   Exit Function


ERR_POP_GOODS_SCREEN:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Resume ERR_POP_GOODS_SCREEN_CLOSE
End If

End Function

Function Populate_Scout_Movement_Screen()
On Error GoTo ERR_TRIBE_NAME_EXIT

If IsNull(Forms![SCOUT MOVEMENT]![TRIBE NAME]) Then
   Exit Function
End If

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

CLAN = "0" & Mid(Forms![SCOUT MOVEMENT]![TRIBE NAME], 2, 3)

Tribe_Checking_Hex = ""
Call Tribe_Checking("Get_Hex", CLAN, Forms![SCOUT MOVEMENT]![TRIBE NAME], "")
Forms![SCOUT MOVEMENT]![Current Hex] = Tribe_Checking_Hex

Set TRIBESTABLE = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESTABLE.MoveFirst
TRIBESTABLE.index = "PRIMARYKEY"
TRIBESTABLE.Seek "=", CLAN, Forms![SCOUT MOVEMENT]![TRIBE NAME]

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst

Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst
HEXTABLE.Seek "=", TRIBESTABLE![Current Hex]
CURRENT_WEATHER_ZONE = HEXTABLE![WEATHER_ZONE]
TRIBESTABLE.Close


If CURRENT_WEATHER_ZONE = "GREEN" Then
   Forms![SCOUT MOVEMENT]![WEATHER] = globalinfo![Zone1]
ElseIf CURRENT_WEATHER_ZONE = "RED" Then
   Forms![SCOUT MOVEMENT]![WEATHER] = globalinfo![Zone2]
ElseIf CURRENT_WEATHER_ZONE = "ORANGE" Then
   Forms![SCOUT MOVEMENT]![WEATHER] = globalinfo![Zone3]
ElseIf CURRENT_WEATHER_ZONE = "YELLOW" Then
   Forms![SCOUT MOVEMENT]![WEATHER] = globalinfo![Zone4]
ElseIf CURRENT_WEATHER_ZONE = "BLUE" Then
   Forms![SCOUT MOVEMENT]![WEATHER] = globalinfo![Zone5]
ElseIf CURRENT_WEATHER_ZONE = "BROWN" Then
   Forms![SCOUT MOVEMENT]![WEATHER] = globalinfo![Zone6]
End If

globalinfo.Close
HEXTABLE.Close

Set SCOUT_MOVEMENT_TABLE = TVDBGM.OpenRecordset("SCOUT_MOVEMENT")
SCOUT_MOVEMENT_TABLE.index = "PRIMARYKEY"
'SCOUT_MOVEMENT_TABLE.MoveFirst
SCOUT_MOVEMENT_TABLE.Seek "=", Forms![SCOUT MOVEMENT]![TRIBE NAME]
If SCOUT_MOVEMENT_TABLE.NoMatch Then
   GoTo ERR_TRIBE_NAME_EXIT_CLOSE
Else
   Forms![SCOUT MOVEMENT]![Scout1Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout1Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout1Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout1Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout1Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout1Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout1Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout1Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS1] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES1] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants1] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels1] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION1] = SCOUT_MOVEMENT_TABLE![MISSION]
End If

SCOUT_MOVEMENT_TABLE.MoveNext
Forms![SCOUT MOVEMENT]![Scout2Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
Forms![SCOUT MOVEMENT]![Scout2Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
Forms![SCOUT MOVEMENT]![Scout2Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
Forms![SCOUT MOVEMENT]![Scout2Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
Forms![SCOUT MOVEMENT]![Scout2Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
Forms![SCOUT MOVEMENT]![Scout2Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
Forms![SCOUT MOVEMENT]![Scout2Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
Forms![SCOUT MOVEMENT]![Scout2Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
Forms![SCOUT MOVEMENT]![SCOUTS2] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
Forms![SCOUT MOVEMENT]![HORSES2] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
Forms![SCOUT MOVEMENT]![Elephants2] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
Forms![SCOUT MOVEMENT]![Camels2] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
Forms![SCOUT MOVEMENT]![MISSION2] = SCOUT_MOVEMENT_TABLE![MISSION]

SCOUT_MOVEMENT_TABLE.MoveNext
   Forms![SCOUT MOVEMENT]![Scout3Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout3Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout3Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout3Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout3Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout3Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout3Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout3Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS3] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES3] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants3] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels3] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION3] = SCOUT_MOVEMENT_TABLE![MISSION]

SCOUT_MOVEMENT_TABLE.MoveNext
   Forms![SCOUT MOVEMENT]![Scout4Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout4Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout4Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout4Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout4Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout4Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout4Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout4Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS4] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES4] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants4] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels4] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION4] = SCOUT_MOVEMENT_TABLE![MISSION]

SCOUT_MOVEMENT_TABLE.MoveNext
   Forms![SCOUT MOVEMENT]![Scout5Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout5Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout5Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout5Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout5Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout5Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout5Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout5Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS5] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES5] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants5] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels5] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION5] = SCOUT_MOVEMENT_TABLE![MISSION]

SCOUT_MOVEMENT_TABLE.MoveNext
   Forms![SCOUT MOVEMENT]![Scout6Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout6Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout6Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout6Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout6Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout6Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout6Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout6Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS6] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES6] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants6] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels6] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION6] = SCOUT_MOVEMENT_TABLE![MISSION]

SCOUT_MOVEMENT_TABLE.MoveNext
   Forms![SCOUT MOVEMENT]![Scout7Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout7Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout7Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout7Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout7Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout7Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout7Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout7Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS7] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES7] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants7] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels7] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION7] = SCOUT_MOVEMENT_TABLE![MISSION]

SCOUT_MOVEMENT_TABLE.MoveNext
   Forms![SCOUT MOVEMENT]![Scout8Move01] = SCOUT_MOVEMENT_TABLE![Movement1]
   Forms![SCOUT MOVEMENT]![Scout8Move02] = SCOUT_MOVEMENT_TABLE![Movement2]
   Forms![SCOUT MOVEMENT]![Scout8Move03] = SCOUT_MOVEMENT_TABLE![Movement3]
   Forms![SCOUT MOVEMENT]![Scout8Move04] = SCOUT_MOVEMENT_TABLE![Movement4]
   Forms![SCOUT MOVEMENT]![Scout8Move05] = SCOUT_MOVEMENT_TABLE![Movement5]
   Forms![SCOUT MOVEMENT]![Scout8Move06] = SCOUT_MOVEMENT_TABLE![Movement6]
   Forms![SCOUT MOVEMENT]![Scout8Move07] = SCOUT_MOVEMENT_TABLE![Movement7]
   Forms![SCOUT MOVEMENT]![Scout8Move08] = SCOUT_MOVEMENT_TABLE![Movement8]
   Forms![SCOUT MOVEMENT]![SCOUTS8] = SCOUT_MOVEMENT_TABLE![No_of_Scouts]
   Forms![SCOUT MOVEMENT]![HORSES8] = SCOUT_MOVEMENT_TABLE![No_of_Horses]
   Forms![SCOUT MOVEMENT]![Elephants8] = SCOUT_MOVEMENT_TABLE![No_of_Elephants]
   Forms![SCOUT MOVEMENT]![Camels8] = SCOUT_MOVEMENT_TABLE![No_of_Camels]
   Forms![SCOUT MOVEMENT]![MISSION8] = SCOUT_MOVEMENT_TABLE![MISSION]

ERR_TRIBE_NAME_EXIT_CLOSE:
   Exit Function


ERR_TRIBE_NAME_EXIT:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Resume ERR_TRIBE_NAME_EXIT_CLOSE
End If


End Function



Public Function Populate_Trading_Post_Screen()
Dim TYPE_OF_TRADING_POST As String
Dim TRIBE As String
Dim GOOD As String
Dim HEX_MAP_ID As String
Dim BUY_PRICE As Double
Dim BUY_LIMIT As Double
Dim BUY_RESET_WAIT As String
Dim NORMAL_BUY_LIMIT As Double
Dim TURNS_SINCE_LAST_BUY As Double
Dim BUY_THIS_TURN As String
Dim BUY_TOTAL As Double
Dim SELL_PRICE As Double
Dim SELL_LIMIT As Double
Dim SELL_RESET_WAIT As String
Dim NORMAL_SELL_LIMIT As Double
Dim TURNS_SINCE_LAST_SELL As Double
Dim SELL_THIS_TURN As String
Dim SELL_TOTAL As Double

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![TRADING POST]
HEX_MAP_ID = Left(MYFORM![Current Hex], 2)

Set activitytable = TVDBGM.OpenRecordset("Trading_Post_Goods")
' NEED TO MODIFY THIS TO HANDLE NEW FORMAT TABLE
If Left(MYFORM![CITY], 1) = "0" Or Left(MYFORM![CITY], 1) = "1" _
Or Left(MYFORM![CITY], 1) = "2" Or Left(MYFORM![CITY], 1) = "3" _
Or Left(MYFORM![CITY], 1) = "4" Or Left(MYFORM![CITY], 1) = "5" _
Or Left(MYFORM![CITY], 1) = "6" Or Left(MYFORM![CITY], 1) = "7" _
Or Left(MYFORM![CITY], 1) = "8" Or Left(MYFORM![CITY], 1) = "9" Then
   activitytable.index = "TRIBESGOOD"
   activitytable.MoveFirst
   If MYFORM![TRADE_TYPE] = "SELL" Then
      If Not IsNull(MYFORM![ITEM]) Then
         activitytable.Seek "=", MYFORM![CITY], MYFORM![ITEM]
         Forms![TRADING POST]![PRICE] = activitytable![BUY PRICE]
      End If
   Else
      If Not IsNull(MYFORM![ITEM]) Then
         activitytable.Seek "=", MYFORM![CITY], MYFORM![ITEM]
         Forms![TRADING POST]![PRICE] = activitytable![SELL PRICE]
      End If
   End If
ElseIf MYFORM![CITY] = "GM SALE" Or MYFORM![CITY] = "FAIR" Then
   activitytable.index = "HEX_MAP_ID"
   activitytable.MoveFirst
   ' CHECK TO SEE IF PRICE LIST EXISTS FOR THE HEX_MAP_ID
   ' IF NOT THEN CREATE IT.
   activitytable.Seek "=", HEX_MAP_ID, MYFORM![ITEM]
   If activitytable.NoMatch Then
      ' get the records in a view, modify the id and then update.
      
      activitytable.MoveFirst
      Do While Not activitytable.EOF
         If activitytable![TRIBE] = MYFORM![CITY] And activitytable![HEX_MAP_ID] = HEX_MAP_ID Then
            ' can i copy all of the records and insert them changing 1 field
            TYPE_OF_TRADING_POST = activitytable![TYPE_OF_TRADING_POST]
            TRIBE = activitytable![TRIBE]
            GOOD = activitytable![GOOD]
            HEX_MAP_ID = Left(MYFORM![Current Hex], 2)
            BUY_PRICE = activitytable![BUY PRICE]
            BUY_LIMIT = activitytable![BUY LIMIT]
            BUY_RESET_WAIT = activitytable![BUY_RESET_WAIT]
            NORMAL_BUY_LIMIT = activitytable![NORMAL_BUY_LIMIT]
            TURNS_SINCE_LAST_BUY = activitytable![TURNS_SINCE_LAST_BUY]
            BUY_THIS_TURN = activitytable![BUY_THIS_TURN]
            BUY_TOTAL = activitytable![BUY_TOTAL]
            SELL_PRICE = activitytable![SELL PRICE]
            SELL_LIMIT = activitytable![SELL LIMIT]
            SELL_RESET_WAIT = activitytable![SELL_RESET_WAIT]
            NORMAL_SELL_LIMIT = activitytable![NORMAL_SELL_LIMIT]
            TURNS_SINCE_LAST_SELL = activitytable![TURNS_SINCE_LAST_SELL]
            SELL_THIS_TURN = activitytable![SELL_THIS_TURN]
            SELL_TOTAL = activitytable![SELL_TOTAL]
            activitytable.AddNew
            activitytable![TYPE_OF_TRADING_POST] = TYPE_OF_TRADING_POST
            activitytable![TRIBE] = TRIBE
            activitytable![GOOD] = GOOD
            activitytable![HEX_MAP_ID] = HEX_MAP_ID
            activitytable![BUY PRICE] = BUY_PRICE
            activitytable![BUY LIMIT] = BUY_LIMIT
            activitytable![BUY_RESET_WAIT] = BUY_RESET_WAIT
            activitytable![NORMAL_BUY_LIMIT] = NORMAL_BUY_LIMIT
            activitytable![TURNS_SINCE_LAST_BUY] = TURNS_SINCE_LAST_BUY
            activitytable![BUY_THIS_TURN] = BUY_THIS_TURN
            activitytable![BUY_TOTAL] = BUY_TOTAL
            activitytable![SELL PRICE] = SELL_PRICE
            activitytable![SELL LIMIT] = SELL_LIMIT
            activitytable![SELL_RESET_WAIT] = SELL_RESET_WAIT
            activitytable![NORMAL_SELL_LIMIT] = NORMAL_SELL_LIMIT
            activitytable![TURNS_SINCE_LAST_SELL] = TURNS_SINCE_LAST_SELL
            activitytable![SELL_THIS_TURN] = SELL_THIS_TURN
            activitytable![SELL_TOTAL] = SELL_TOTAL
            activitytable.UPDATE
         End If
         activitytable.MoveNext
         If activitytable.EOF Then
            Exit Do
         End If
      Loop
   End If
   If MYFORM![TRADE_TYPE] = "SELL" Then
      If Not IsNull(MYFORM![ITEM]) Then
         activitytable.Seek "=", HEX_MAP_ID, MYFORM![ITEM]
         Forms![TRADING POST]![PRICE] = activitytable![BUY PRICE]
      End If
   Else
      If Not IsNull(MYFORM![ITEM]) Then
         activitytable.Seek "=", HEX_MAP_ID, MYFORM![ITEM]
         Forms![TRADING POST]![PRICE] = activitytable![SELL PRICE]
      End If
   End If

Else
   activitytable.index = "TRIBESGOOD"
   
   activitytable.MoveFirst
   If MYFORM![TRADE_TYPE] = "SELL" Then
      If Not IsNull(MYFORM![ITEM]) Then
         activitytable.Seek "=", MYFORM![CITY], MYFORM![ITEM]
         Forms![TRADING POST]![PRICE] = activitytable![BUY PRICE]
      End If
   Else
      If Not IsNull(MYFORM![ITEM]) Then
         activitytable.Seek "=", MYFORM![CITY], MYFORM![ITEM]
         Forms![TRADING POST]![PRICE] = activitytable![SELL PRICE]
      End If
   End If
End If

activitytable.Close

End Function


Public Function Show_Turns_Activities_Controls()
Dim AMOUNT_OF_ITEM As Long
Dim Slaves As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![TURNS ACTIVITIES]
TRIBE = MYFORM![TRIBENUMBER]
CLAN = "0" & Mid(TRIBE, 2, 3)

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", CLAN, TRIBE
  
If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
   GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
Else
   GOODS_TRIBE = TRIBE
End If
  
If IsNull(TRIBESINFO![SLAVE]) Then
   Slaves = "N"
ElseIf TRIBESINFO![SLAVE] = 0 Then
   Slaves = "N"
Else
   Slaves = "Y"
End If

If MYFORM![ACTIVITY01] = "ALCHEMY" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "APIARISM" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "BEEKEEPER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   Else
      AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "APIARIST")
      If AMOUNT_OF_ITEM > 0 Then
         MYFORM![SPECIALISTS].Visible = True
         MYFORM![Specialists_text].Visible = True
      End If
   End If
ElseIf MYFORM![ACTIVITY01] = "ARMOUR" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "ARMOURER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "ATHEISM" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "BAKING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "BAKER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "BAMBOOWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "BLUBBERWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "BONEWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "BONING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "BRICK MAKING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "BRICKLAYER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "CHEESE MAKING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "Convert" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "Cooking" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "CURING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "DANCING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "Default" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "DEFENCE" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "DISTILLING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "DISTILLER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "DRESSING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "DRINKING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "EATING" Then
   ' DO NOTHING
ElseIf (MYFORM![ACTIVITY01] = "ENGINEERING") Or (MYFORM![ACTIVITY01] = "SHIPBUILDING") Then
   MYFORM![Building].Visible = True
   MYFORM![Building_Text].Visible = True
   MYFORM![JOINT_PROJECT].Visible = True
   MYFORM![JOINT_PROJECT_Text].Visible = True
   MYFORM![Eng_Clan].Visible = True
   MYFORM![Eng_clan_text].Visible = True
   MYFORM![Eng_Tribe].Visible = True
   MYFORM![Eng_Tribe_Text].Visible = True
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "ENGINEER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "EXCAVATION" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "FISHING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "FISHER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "FLENSING" Then
   MYFORM![Whale_Size].Visible = True
   MYFORM![Whale_Size_Text].Visible = True
ElseIf MYFORM![ACTIVITY01] = "FLENSING&PEELING" Then
   MYFORM![Whale_Size].Visible = True
   MYFORM![Whale_Size_Text].Visible = True
ElseIf MYFORM![ACTIVITY01] = "FLENSING&PEELING&BONING" Then
   MYFORM![Whale_Size].Visible = True
   MYFORM![Whale_Size_Text].Visible = True
ElseIf MYFORM![ACTIVITY01] = "FLETCHING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "FORAGING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "FORESTRY" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "FORESTER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "FURRIER" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "FURRIER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "GATHERING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "GLASSWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "GUT&BONE" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "GUT&SKIN" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "GUTTING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "HARVEST" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "FARMER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "HERDING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "HERDER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "HUNTING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "HUNTER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "JEWELLERY" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "KILLING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "LEATHERWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "LITERACY" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "MAINTAINING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "METALWORK" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "METALWORKER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "MILLING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "MILLER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "MINING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "MINER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   If MYFORM![item01] = "LOW YIELD EXTRACTION" Then
      MYFORM![Mine_Direction].Visible = True
      MYFORM![Mine_Direction_Text].Visible = True
   End If
   If Slaves = "Y" Then
      MYFORM![Slaves].Visible = True
      MYFORM![SLAVES_TEXT].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "MUSIC" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "PACIFICATION" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "PEELING" Then
   MYFORM![Whale_Size].Visible = True
   MYFORM![Whale_Size_Text].Visible = True
ElseIf MYFORM![ACTIVITY01] = "PLANTING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "FARMER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "PLOWING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "FARMER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "POLITICS" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "POTTERY" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "QUARRYING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "QUARRIER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "REFINING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "REFINER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "RELIGION" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "RESEARCH" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SALTING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SAND GATHERING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SB" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SCOUTING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SECURITY" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SEEKING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SEWING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SG" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SGB" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SHEARING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SIEGE EQUIPMENT" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SKIN&BONE" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SKIN&GUT" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SKIN&GUT&BONE" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SKINNING" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SLAVERY" Then
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], "ALL", Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "SMOKING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SPECIALISTS" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "STONEWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "SUPPRESSION" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "TANNING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "TRIBE DEFENCE" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "WAXWORK" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "WEAPONS" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "WEAPONSMITH")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
   Call Process_Implements_Table_and_Update_Form(CLAN, GOODS_TRIBE, MYFORM![ACTIVITY01], MYFORM![item01], Forms![TURNS ACTIVITIES])
ElseIf MYFORM![ACTIVITY01] = "WEAVING" Then
   AMOUNT_OF_ITEM = GET_TRIBES_SPECIALISTS_QUANTITY(CLAN, TRIBE, "WEAVER")
   If AMOUNT_OF_ITEM > 0 Then
      MYFORM![SPECIALISTS].Visible = True
      MYFORM![Specialists_text].Visible = True
   End If
ElseIf MYFORM![ACTIVITY01] = "WHALING" Then
   ' DO NOTHING
ElseIf MYFORM![ACTIVITY01] = "WOODWORK" Then
   ' DO NOTHING
Else
   MsgBox "The Show_Turns_Activities_Controls procedure does not cater for this activity"
End If

End Function

Public Function Process_Implements_Table_and_Update_Form(CLAN, TRIBE, ACTIVITY, ITEM, Form)

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Form

Set ImplementsTable = TVDB.OpenRecordset("IMPLEMENTS")
ImplementsTable.index = "MODIFIER"
ImplementsTable.MoveFirst

Set ImplementUsage = TVDB.OpenRecordset("IMPLEMENT_USAGE")
ImplementUsage.index = "PRIMARYKEY"
ImplementUsage.MoveFirst

count = 1

Do
  If ImplementsTable![ACTIVITY] = ACTIVITY And ImplementsTable![ITEM] = ITEM Then
     Exit Do
  Else
     ImplementsTable.MoveNext
     If ImplementsTable.EOF Then
        GoTo CLOSE_PROCESS_IMPLEMENTS_TABLE_AND_UPDATE
     End If
  End If
  
Loop

Do While ImplementsTable![ACTIVITY] = ACTIVITY And ImplementsTable![ITEM] = ITEM
   If count = 1 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_01].Visible = True
            MYFORM![Item1_Text].Visible = True
            MYFORM![Item_01] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_01] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_01].Visible = True
            MYFORM![Amt1_text].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 2 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_02].Visible = True
            MYFORM![Item2_text].Visible = True
            MYFORM![Item_02] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_02] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_02].Visible = True
            MYFORM![Amt2_text].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 3 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_03].Visible = True
            MYFORM![Item_03] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_03] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_03].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 4 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_04].Visible = True
            MYFORM![Item_04] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_04] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_04].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 5 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_05].Visible = True
            MYFORM![Item_05] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_05] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_05].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 6 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_06].Visible = True
            MYFORM![Item_06] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_06] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_06].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 7 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_07].Visible = True
            MYFORM![Item_07] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_07] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_07].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 8 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_08].Visible = True
            MYFORM![Item_08] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_08] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_08].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 9 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![Item_09].Visible = True
            MYFORM![Item_09] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_09] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_09].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 10 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![ITEM_10].Visible = True
            MYFORM![ITEM_10] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_10] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_10].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 11 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![ITEM_11].Visible = True
            MYFORM![ITEM_11] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_11] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_11].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 12 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![ITEM_12].Visible = True
            MYFORM![ITEM_12] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_12] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_12].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 13 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![ITEM_13].Visible = True
            MYFORM![ITEM_13] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_13] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_13].Visible = True
            count = count + 1
         End If
      End If
   ElseIf count = 14 Then
      ImplementUsage.Seek "=", CLAN, TRIBE, ImplementsTable![IMPLEMENT]
      If Not ImplementUsage.NoMatch Then
         If (ImplementUsage![total_available] - ImplementUsage![Number_Used]) > 0 Then
            MYFORM![ITEM_14].Visible = True
            MYFORM![ITEM_14] = ImplementsTable![IMPLEMENT]
            MYFORM![Use_Amt_14] = ImplementUsage![total_available] - ImplementUsage![Number_Used]
            MYFORM![Use_Amt_14].Visible = True
            count = count + 1
         End If
      End If
   End If
   ImplementsTable.MoveNext
   If ImplementsTable.EOF Then
      Exit Do
   End If
Loop
ImplementsTable.Close
ImplementUsage.Close

CLOSE_PROCESS_IMPLEMENTS_TABLE_AND_UPDATE:

End Function

Public Function Tribe_Movement_Orders(reference)
Dim strMoves As String

Set MYFORM = Forms![TRIBE MOVEMENT]

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

Set VALID_DIRECTIONS = TVDB.OpenRecordset("VALID_DIRECTIONS")
VALID_DIRECTIONS.MoveFirst
VALID_DIRECTIONS.index = "SECONDARYKEY"

If Not reference = "Description" Then
VALID_DIRECTIONS.Seek "=", MYFORM(reference).Value, "Y"
If Not VALID_DIRECTIONS.NoMatch Then
   If MYFORM(reference) = "FOLLOW" Then
      MYFORM![Follow_Tribe].Visible = True
      MYFORM![Label134].Visible = True
   End If
   
   If Right(reference, 2) < 35 Then
      MYFORM![Movement35].Visible = False
   End If
   If Right(reference, 2) < 34 Then
      MYFORM![Movement34].Visible = False
   End If
   If Right(reference, 2) < 33 Then
      MYFORM![Movement33].Visible = False
   End If
   If Right(reference, 2) < 32 Then
      MYFORM![Movement32].Visible = False
   End If
   If Right(reference, 2) < 31 Then
      MYFORM![Movement31].Visible = False
   End If
   If Right(reference, 2) < 30 Then
      MYFORM![Movement30].Visible = False
   End If
   If Right(reference, 2) < 29 Then
      MYFORM![Movement29].Visible = False
   End If
   If Right(reference, 2) < 28 Then
      MYFORM![Movement28].Visible = False
   End If
   If Right(reference, 2) < 27 Then
      MYFORM![Movement27].Visible = False
   End If
   If Right(reference, 2) < 26 Then
      MYFORM![Movement26].Visible = False
   End If
   If Right(reference, 2) < 25 Then
      MYFORM![Movement25].Visible = False
   End If
   If Right(reference, 2) < 24 Then
      MYFORM![Movement24].Visible = False
   End If
   If Right(reference, 2) < 23 Then
      MYFORM![Movement23].Visible = False
   End If
   If Right(reference, 2) < 22 Then
      MYFORM![Movement22].Visible = False
   End If
   If Right(reference, 2) < 21 Then
      MYFORM![Movement21].Visible = False
   End If
   If Right(reference, 2) < 20 Then
      MYFORM![Movement20].Visible = False
   End If
   If Right(reference, 2) < 19 Then
      MYFORM![Movement19].Visible = False
   End If
   If Right(reference, 2) < 18 Then
      MYFORM![Movement18].Visible = False
   End If
   If Right(reference, 2) < 17 Then
      MYFORM![Movement17].Visible = False
   End If
   If Right(reference, 2) < 16 Then
      MYFORM![Movement16].Visible = False
   End If
   If Right(reference, 2) < 15 Then
      MYFORM![Movement15].Visible = False
   End If
   If Right(reference, 2) < 14 Then
      MYFORM![Movement14].Visible = False
   End If
   If Right(reference, 2) < 13 Then
      MYFORM![Movement13].Visible = False
   End If
   If Right(reference, 2) < 12 Then
      MYFORM![Movement12].Visible = False
   End If
   If Right(reference, 2) < 11 Then
      MYFORM![Movement11].Visible = False
   End If
   If Right(reference, 2) < 10 Then
      MYFORM![Movement10].Visible = False
   End If
   If Right(reference, 2) < 9 Then
      MYFORM![Movement09].Visible = False
   End If
   If Right(reference, 2) < 8 Then
      MYFORM![Movement08].Visible = False
   End If
   If Right(reference, 2) < 7 Then
      MYFORM![Movement07].Visible = False
   End If
   If Right(reference, 2) < 6 Then
      MYFORM![Movement06].Visible = False
   End If
   If Right(reference, 2) < 5 Then
      MYFORM![Movement05].Visible = False
   End If
   If Right(reference, 2) < 4 Then
      MYFORM![Movement04].Visible = False
   End If
   If Right(reference, 2) < 3 Then
      MYFORM![Movement03].Visible = False
   End If
   If Right(reference, 2) < 2 Then
      MYFORM![Movement02].Visible = False
   End If
End If
ElseIf reference = "Description" Then
   VALID_DIRECTIONS.MoveFirst
   VALID_DIRECTIONS.Seek "=", MYFORM![Direction], "Y"
   MYFORM![Description] = VALID_DIRECTIONS![Description]
   VALID_DIRECTIONS.Close
End If

End Function



Public Function Scout_Movement_Orders(reference)
Dim strMoves As String

Set MYFORM = Forms![SCOUT MOVEMENT]

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

Set VALID_DIRECTIONS = TVDB.OpenRecordset("VALID_DIRECTIONS")
VALID_DIRECTIONS.MoveFirst
VALID_DIRECTIONS.index = "SECONDARYKEY"

If reference = "Scout1Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move02].Visible = False
      MYFORM![Scout1Move03].Visible = False
      MYFORM![Scout1Move04].Visible = False
      MYFORM![Scout1Move05].Visible = False
      MYFORM![Scout1Move06].Visible = False
      MYFORM![Scout1Move07].Visible = False
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout1Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move03].Visible = False
      MYFORM![Scout1Move04].Visible = False
      MYFORM![Scout1Move05].Visible = False
      MYFORM![Scout1Move06].Visible = False
      MYFORM![Scout1Move07].Visible = False
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout1Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move04].Visible = False
      MYFORM![Scout1Move05].Visible = False
      MYFORM![Scout1Move06].Visible = False
      MYFORM![Scout1Move07].Visible = False
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout1Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move05].Visible = False
      MYFORM![Scout1Move06].Visible = False
      MYFORM![Scout1Move07].Visible = False
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout1Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move06].Visible = False
      MYFORM![Scout1Move07].Visible = False
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout1Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move07].Visible = False
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout1Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout1Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout1Move08].Visible = False
   End If
ElseIf reference = "Scout2Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move02].Visible = False
      MYFORM![Scout2Move03].Visible = False
      MYFORM![Scout2Move04].Visible = False
      MYFORM![Scout2Move05].Visible = False
      MYFORM![Scout2Move06].Visible = False
      MYFORM![Scout2Move07].Visible = False
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout2Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move03].Visible = False
      MYFORM![Scout2Move04].Visible = False
      MYFORM![Scout2Move05].Visible = False
      MYFORM![Scout2Move06].Visible = False
      MYFORM![Scout2Move07].Visible = False
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout2Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move04].Visible = False
      MYFORM![Scout2Move05].Visible = False
      MYFORM![Scout2Move06].Visible = False
      MYFORM![Scout2Move07].Visible = False
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout2Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move05].Visible = False
      MYFORM![Scout2Move06].Visible = False
      MYFORM![Scout2Move07].Visible = False
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout2Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move06].Visible = False
      MYFORM![Scout2Move07].Visible = False
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout2Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move07].Visible = False
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout2Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout2Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout2Move08].Visible = False
   End If
ElseIf reference = "Scout3Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move02].Visible = False
      MYFORM![Scout3Move03].Visible = False
      MYFORM![Scout3Move04].Visible = False
      MYFORM![Scout3Move05].Visible = False
      MYFORM![Scout3Move06].Visible = False
      MYFORM![Scout3Move07].Visible = False
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout3Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move03].Visible = False
      MYFORM![Scout3Move04].Visible = False
      MYFORM![Scout3Move05].Visible = False
      MYFORM![Scout3Move06].Visible = False
      MYFORM![Scout3Move07].Visible = False
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout3Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move04].Visible = False
      MYFORM![Scout3Move05].Visible = False
      MYFORM![Scout3Move06].Visible = False
      MYFORM![Scout3Move07].Visible = False
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout3Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move05].Visible = False
      MYFORM![Scout3Move06].Visible = False
      MYFORM![Scout3Move07].Visible = False
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout3Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move06].Visible = False
      MYFORM![Scout3Move07].Visible = False
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout3Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move07].Visible = False
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout3Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout3Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout3Move08].Visible = False
   End If
ElseIf reference = "Scout4Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move02].Visible = False
      MYFORM![Scout4Move03].Visible = False
      MYFORM![Scout4Move04].Visible = False
      MYFORM![Scout4Move05].Visible = False
      MYFORM![Scout4Move06].Visible = False
      MYFORM![Scout4Move07].Visible = False
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout4Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move03].Visible = False
      MYFORM![Scout4Move04].Visible = False
      MYFORM![Scout4Move05].Visible = False
      MYFORM![Scout4Move06].Visible = False
      MYFORM![Scout4Move07].Visible = False
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout4Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move04].Visible = False
      MYFORM![Scout4Move05].Visible = False
      MYFORM![Scout4Move06].Visible = False
      MYFORM![Scout4Move07].Visible = False
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout4Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move05].Visible = False
      MYFORM![Scout4Move06].Visible = False
      MYFORM![Scout4Move07].Visible = False
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout4Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move06].Visible = False
      MYFORM![Scout4Move07].Visible = False
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout4Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move07].Visible = False
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout4Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout4Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout4Move08].Visible = False
   End If
ElseIf reference = "Scout5Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move02].Visible = False
      MYFORM![Scout5Move03].Visible = False
      MYFORM![Scout5Move04].Visible = False
      MYFORM![Scout5Move05].Visible = False
      MYFORM![Scout5Move06].Visible = False
      MYFORM![Scout5Move07].Visible = False
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout5Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move03].Visible = False
      MYFORM![Scout5Move04].Visible = False
      MYFORM![Scout5Move05].Visible = False
      MYFORM![Scout5Move06].Visible = False
      MYFORM![Scout5Move07].Visible = False
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout5Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move04].Visible = False
      MYFORM![Scout5Move05].Visible = False
      MYFORM![Scout5Move06].Visible = False
      MYFORM![Scout5Move07].Visible = False
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout5Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move05].Visible = False
      MYFORM![Scout5Move06].Visible = False
      MYFORM![Scout5Move07].Visible = False
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout5Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move06].Visible = False
      MYFORM![Scout5Move07].Visible = False
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout5Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move07].Visible = False
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout5Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout5Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout5Move08].Visible = False
   End If
ElseIf reference = "Scout6Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move02].Visible = False
      MYFORM![Scout6Move03].Visible = False
      MYFORM![Scout6Move04].Visible = False
      MYFORM![Scout6Move05].Visible = False
      MYFORM![Scout6Move06].Visible = False
      MYFORM![Scout6Move07].Visible = False
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout6Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move03].Visible = False
      MYFORM![Scout6Move04].Visible = False
      MYFORM![Scout6Move05].Visible = False
      MYFORM![Scout6Move06].Visible = False
      MYFORM![Scout6Move07].Visible = False
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout6Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move04].Visible = False
      MYFORM![Scout6Move05].Visible = False
      MYFORM![Scout6Move06].Visible = False
      MYFORM![Scout6Move07].Visible = False
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout6Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move05].Visible = False
      MYFORM![Scout6Move06].Visible = False
      MYFORM![Scout6Move07].Visible = False
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout6Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move06].Visible = False
      MYFORM![Scout6Move07].Visible = False
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout6Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move07].Visible = False
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout6Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout6Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout6Move08].Visible = False
   End If
ElseIf reference = "Scout7Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move02].Visible = False
      MYFORM![Scout7Move03].Visible = False
      MYFORM![Scout7Move04].Visible = False
      MYFORM![Scout7Move05].Visible = False
      MYFORM![Scout7Move06].Visible = False
      MYFORM![Scout7Move07].Visible = False
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout7Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move03].Visible = False
      MYFORM![Scout7Move04].Visible = False
      MYFORM![Scout7Move05].Visible = False
      MYFORM![Scout7Move06].Visible = False
      MYFORM![Scout7Move07].Visible = False
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout7Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move04].Visible = False
      MYFORM![Scout7Move05].Visible = False
      MYFORM![Scout7Move06].Visible = False
      MYFORM![Scout7Move07].Visible = False
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout7Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move05].Visible = False
      MYFORM![Scout7Move06].Visible = False
      MYFORM![Scout7Move07].Visible = False
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout7Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move06].Visible = False
      MYFORM![Scout7Move07].Visible = False
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout7Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move07].Visible = False
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout7Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout7Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout7Move08].Visible = False
   End If
ElseIf reference = "Scout8Move01" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move01], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move02].Visible = False
      MYFORM![Scout8Move03].Visible = False
      MYFORM![Scout8Move04].Visible = False
      MYFORM![Scout8Move05].Visible = False
      MYFORM![Scout8Move06].Visible = False
      MYFORM![Scout8Move07].Visible = False
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Scout8Move02" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move02], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move03].Visible = False
      MYFORM![Scout8Move04].Visible = False
      MYFORM![Scout8Move05].Visible = False
      MYFORM![Scout8Move06].Visible = False
      MYFORM![Scout8Move07].Visible = False
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Scout8Move03" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move03], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move04].Visible = False
      MYFORM![Scout8Move05].Visible = False
      MYFORM![Scout8Move06].Visible = False
      MYFORM![Scout8Move07].Visible = False
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Scout8Move04" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move04], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move05].Visible = False
      MYFORM![Scout8Move06].Visible = False
      MYFORM![Scout8Move07].Visible = False
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Scout8Move05" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move05], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move06].Visible = False
      MYFORM![Scout8Move07].Visible = False
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Scout8Move06" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move06], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move07].Visible = False
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Scout8Move07" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Scout8Move07], "Y"
   If Not VALID_DIRECTIONS.NoMatch Then
      MYFORM![Scout8Move08].Visible = False
   End If
ElseIf reference = "Description" Then
   VALID_DIRECTIONS.Seek "=", MYFORM![Direction]
   MYFORM![Description] = VALID_DIRECTIONS![Description]
   VALID_DIRECTIONS.Close
End If

End Function



Public Function Update_Transfers_Screen_Hex_Info(Direction)

If Direction = "FROM" Then
   If IsNull(Forms![TRANSFER_GOODS]![FROM CLAN]) Then
      Exit Function
   End If
   
   If IsNull(Forms![TRANSFER_GOODS]![FROM TRIBE]) Then
      Exit Function
   End If
   
   CLANNUMBER = Forms![TRANSFER_GOODS]![FROM CLAN]
   TRIBENUMBER = Forms![TRANSFER_GOODS]![FROM TRIBE]
Else
   If IsNull(Forms![TRANSFER_GOODS]![TO CLAN]) Then
      Exit Function
   End If
   
   If IsNull(Forms![TRANSFER_GOODS]![TO TRIBE]) Then
      Exit Function
   End If
   
   CLANNUMBER = Forms![TRANSFER_GOODS]![TO CLAN]
   TRIBENUMBER = Forms![TRANSFER_GOODS]![TO TRIBE]
End If

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

If Direction = "FROM" Then
   Tribe_Checking_Hex = ""
   Call Tribe_Checking("Get_Hex", CLANNUMBER, TRIBENUMBER, "")
   Forms![TRANSFER_GOODS]![FROM HEX] = Tribe_Checking_Hex
Else
   Tribe_Checking_Hex = ""
   Call Tribe_Checking("Get_Hex", CLANNUMBER, TRIBENUMBER, "")
   Forms![TRANSFER_GOODS]![TO HEX] = Tribe_Checking_Hex
   Call Go_To_Field("ITEM")
End If

End Function
Public Function Trading_Post_Tribenumber_Exit()

If IsNull(Forms![TRADING POST]![CLANNUMBER]) Then
   Exit Function
End If

If IsNull(Forms![TRADING POST]![TRIBENUMBER]) Then
   Exit Function
End If

CLAN = Forms![TRADING POST]![CLANNUMBER]
TRIBE = Forms![TRADING POST]![TRIBENUMBER]

' TRIBE MOVEMENT
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set TRIBEINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBEINFO.MoveFirst
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.Seek "=", CLAN, TRIBE

If Not IsNull(TRIBEINFO![GOODS TRIBE]) Then
   GOODS_TRIBE = TRIBEINFO![GOODS TRIBE]
Else
   GOODS_TRIBE = TRIBE
End If

CURRENT_HEX = TRIBEINFO![Current Hex]

Forms![TRADING POST]![Current Hex] = CURRENT_HEX

End Function



Public Function Go_To_Record(CLAN, TRIBE)
    
    DoCmd.FindRecord TRIBE, acEntire

End Function

Public Function Populate_the_Pacification_Screen()
On Error GoTo ERR_Populate_the_Pacification_Screen_Exit
Dim fullfield As String
Dim hex_pop As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![PACIFICATION]
CLAN = MYFORM![CLAN NAME]
TRIBE = MYFORM![TRIBE NAME]

Set PACIFICATION_TABLE = TVDBGM.OpenRecordset("PACIFICATION_TABLE")
PACIFICATION_TABLE.index = "PRIMARYKEY"
If PACIFICATION_TABLE.BOF Then
     ' do nothing
Else
     PACIFICATION_TABLE.MoveFirst
End If

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", CLAN, TRIBE
  
Tribes_Current_Hex = TRIBESINFO![Current Hex]
TRIBES_TERRAIN = TRIBESINFO![CURRENT TERRAIN]

Call Populate_the_pacification_Back_color("PRAIRIE", "PRAIRIE")
Call Populate_the_pacification_Back_color("HILLS", "HILLS")
Call Populate_the_pacification_Back_color("FOREST", "FOREST")
Call Populate_the_pacification_Back_color("SWAMP", "SWAMP")
Call Populate_the_pacification_Back_color("OCEAN", "OCEAN")
Call Populate_the_pacification_Back_color("ALPS", "ALPS")
Call Populate_the_pacification_Back_color("OTHERS", "OTHERS")

PACIFICATION_TABLE.Seek "=", CLAN, TRIBE
If PACIFICATION_TABLE.NoMatch Then
    PACIFICATION_TABLE.AddNew
    PACIFICATION_TABLE![CLAN] = CLAN
    PACIFICATION_TABLE![TRIBE] = TRIBE
    PACIFICATION_TABLE.UPDATE
    PACIFICATION_TABLE.Seek "=", CLAN, TRIBE
End If

If TRIBESINFO![GOVT LEVEL] >= 0 Then
     Call Populate_the_pacification_Back_color("PRIMARY_HEX_hex", TRIBES_TERRAIN)
     MYFORM![PRIMARY_HEX_hex].Visible = True
     MYFORM![primary_hex].Visible = True
     MYFORM![primary_hex] = PACIFICATION_TABLE![primary_hex]
     MYFORM![primary_hex].SetFocus
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 1 Then
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL1_1_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL1_2_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL1_3_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL1_4_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL1_5_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL1_6_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL1_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL1_" & count)
         If count > 6 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 2 Then
     '   HEX TO N/N
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_1_hex", CURRENT_TERRAIN)
    '   HEX TO N/NE
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_2_hex", CURRENT_TERRAIN)
    '   HEX TO NE/NE
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_3_hex", CURRENT_TERRAIN)
    '   HEX TO NE/SE
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_4_hex", CURRENT_TERRAIN)
    '   HEX TO SE/SE
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_5_hex", CURRENT_TERRAIN)
    '   HEX TO S/SE
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_6_hex", CURRENT_TERRAIN)
    '   HEX TO S/S
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_7_hex", CURRENT_TERRAIN)
    '   HEX TO S/SW
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "SW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_8_hex", CURRENT_TERRAIN)
    '   HEX TO SW/SW
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_9_hex", CURRENT_TERRAIN)
    '   HEX TO SW/NW
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_10_hex", CURRENT_TERRAIN)
    '   HEX TO NW/NW
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_11_hex", CURRENT_TERRAIN)
    '   HEX TO N/NW
    Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
    Call Populate_the_pacification_Back_color("GL2_12_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL2_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL2_" & count)
         If count > 12 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 3 Then
     '   HEX TO N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_1_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_2_hex", CURRENT_TERRAIN)
     '   HEX TO N/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_3_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_4_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_5_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_6_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_7_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "S", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_8_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SE", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_9_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_10_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SW", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_11_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "S", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_12_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_13_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_14_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "SW", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_15_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_16_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "N", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_17_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NW", "NONE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL3_18_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL3_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL3_" & count)
         If count > 18 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 4 Then
     '   HEX TO N/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_1_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_2_hex", CURRENT_TERRAIN)
     '   HEX TO N/NE/NE/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NE", "NE", "N", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_3_hex", CURRENT_TERRAIN)
     '   HEX TO N/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NE", "NE", "NE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_4_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_5_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "SE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_6_hex", CURRENT_TERRAIN)
     '   HEX TO NE/SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "SE", "SE", "NE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_7_hex", CURRENT_TERRAIN)
     '   HEX TO NE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "SE", "SE", "SE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_8_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_9_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "S", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_10_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SE", "SE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_11_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SE", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_12_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_13_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_14_hex", CURRENT_TERRAIN)
     '   HEX TO S/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "SW", "SW", "S", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_15_hex", CURRENT_TERRAIN)
     '   HEX TO S/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "SW", "SW", "SW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_16_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_17_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "NW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_18_hex", CURRENT_TERRAIN)
     '   HEX TO SW/NW/NW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "NW", "NW", "SW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_19_hex", CURRENT_TERRAIN)
      '   HEX TO SW/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "NW", "NW", "NW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_20_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_21_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "N", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_22_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NW", "NW", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_23_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NW/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NW", "N", "NONE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL4_24_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL4_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL4_" & count)
         If count > 24 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 5 Then
     '   HEX TO N/N/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_1_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_2_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NE", "NE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_3_hex", CURRENT_TERRAIN)
     '   HEX TO N/NE/NE/NE/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NE", "NE", "NE", "N", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_4_hex", CURRENT_TERRAIN)
     '   HEX TO N/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NE", "NE", "NE", "NE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_5_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_6_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "SE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_7_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "SE", "SE", "NE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_8_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "SE", "SE", "SE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_9_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "NE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_10_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_11_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "S", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_12_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SE", "SE", "SE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_13_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SE", "SE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_14_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "SE", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_15_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_16_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "SW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_17_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SW", "SW", "S", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_18_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SW", "SW", "SW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_19_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "S", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_20_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_21_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "NW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_22_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "SW", "SW", "SW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_23_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "SW", "SW", "NW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_24_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "SW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_25_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "SW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_26_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "N", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_27_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NW", "NW", "NW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_28_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NW", "NW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_29_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NW", "NONE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL5_30_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL5_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL5_" & count)
         If count > 30 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 6 Then
     '   HEX TO N/N/N/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_1_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_2_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NE/NE/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NE", "NE", "N", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_3_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NE", "NE", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_4_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NE", "NE", "NE", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_5_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "N", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_6_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_7_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "SE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_8_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "SE", "SE", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_9_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "SE", "SE", "SE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_10_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "NE", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_11_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "NE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_12_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_13_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "S", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_14_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/S/S/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "S", "S", "SE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_15_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "S", "S", "S", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_16_hex", CURRENT_TERRAIN)
     '   HEX TO SE/S/S/S/S/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "S", "S", "S", "S", "SE", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_17_hex", CURRENT_TERRAIN)
     '   HEX TO SE/S/S/S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "S", "S", "S", "S", "S", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_18_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_19_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_20_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SW", "SW", "S", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_21_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SW", "SW", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_22_hex", CURRENT_TERRAIN)
     '   HEX TO S/SW/SW/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "SW", "SW", "SW", "SW", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_23_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "S", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_24_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_25_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "NW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_26_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/SW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "SW", "SW", "SW", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_27_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/SW/SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "SW", "SW", "SW", "NW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_28_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "SW", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_29_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "SW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_30_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_31_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "N", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_32_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "N", "N", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_33_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NW", "NW", "NW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_34_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NW", "NW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_35_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NW", "NONE", "NONE")
     Call Populate_the_pacification_Back_color("GL6_36_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL6_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL6_" & count)
         If count > 36 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 7 Then
     '   HEX TO N/N/N/N/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "N", "NONE")
     Call Populate_the_pacification_Back_color("GL7_1_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/N/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_2_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_3_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NE", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_4_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NE", "NE", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_5_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/NE/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NE", "NE", "NE", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_6_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/NE/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "N", "NONE")
     Call Populate_the_pacification_Back_color("GL7_7_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_8_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/NE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_9_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_10_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "SE", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_11_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "NE", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_12_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "NE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_13_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/SE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "NE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_14_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_15_hex", CURRENT_TERRAIN)
     '   HEX TO SE/SE/SE/SE/SE/SE/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "S", "NONE")
     Call Populate_the_pacification_Back_color("GL7_16_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/SE/SE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SE", "SE", "SE", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_17_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SE/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SE", "SE", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_18_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/SE/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "SE", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_19_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/SE/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "SE", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_20_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/S/SE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "SE", "NONE")
     Call Populate_the_pacification_Back_color("GL7_21_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "S", "NONE")
     Call Populate_the_pacification_Back_color("GL7_22_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/S/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_23_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/S/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "SW", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_24_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/S/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "SW", "SW", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_25_hex", CURRENT_TERRAIN)
     '   HEX TO S/S/S/SW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SW", "SW", "SW", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_26_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/S/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "S", "S", "NONE")
     Call Populate_the_pacification_Back_color("GL7_27_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/SW/S
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "S", "NONE")
     Call Populate_the_pacification_Back_color("GL7_28_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_29_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/SW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_30_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/SW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "NW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_31_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/SW/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "NW", "NW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_32_hex", CURRENT_TERRAIN)
     '   HEX TO SW/SW/SW/NW/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "NW", "NW", "NW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_33_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/SW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "SW", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_34_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/NW/SW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "SW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_35_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_36_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/NW/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "N", "NONE")
     Call Populate_the_pacification_Back_color("GL7_37_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/NW/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "N", "N", "NONE")
     Call Populate_the_pacification_Back_color("GL7_38_hex", CURRENT_TERRAIN)
     '   HEX TO NW/NW/NW/NW/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "N", "N", "N", "NONE")
     Call Populate_the_pacification_Back_color("GL7_39_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/NW/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NW", "NW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_40_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/NW/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NW", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_41_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/N/NW
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "NW", "NONE")
     Call Populate_the_pacification_Back_color("GL7_42_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL7_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL7_" & count)
         If count > 42 Then
            Exit Do
         End If
     Loop
End If

count = 1
If TRIBESINFO![GOVT LEVEL] >= 8 Then
     '   HEX TO N/N/N/N/N/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "N", "N")
     Call Populate_the_pacification_Back_color("GL8_1_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/N/N/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "N", "NE")
     Call Populate_the_pacification_Back_color("GL8_2_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/N/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "NE", "NE")
     Call Populate_the_pacification_Back_color("GL8_3_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/N/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NE", "NE", "NE")
     Call Populate_the_pacification_Back_color("GL8_4_hex", CURRENT_TERRAIN)
     '   HEX TO N/N/N/N/NE/NE/NE/NE
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NE", "NE", "NE", "NE")
     Call Populate_the_pacification_Back_color("GL8_5_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/N/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "N", "N", "N")
     Call Populate_the_pacification_Back_color("GL8_6_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/NE/N/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "N", "N")
     Call Populate_the_pacification_Back_color("GL8_7_hex", CURRENT_TERRAIN)
     '   HEX TO NE/NE/NE/NE/NE/NE/NE/N
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "NE", "N")
     Call Populate_the_pacification_Back_color("GL8_8_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "NE", "NE")
     Call Populate_the_pacification_Back_color("GL8_9_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "NE", "SE")
     Call Populate_the_pacification_Back_color("GL8_10_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "NE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_11_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "NE", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_12_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NE", "NE", "NE", "NE", "SE", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_13_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "NE", "NE", "NE")
     Call Populate_the_pacification_Back_color("GL8_14_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "NE", "NE")
     Call Populate_the_pacification_Back_color("GL8_15_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "SE", "NE")
     Call Populate_the_pacification_Back_color("GL8_16_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_17_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SE", "SE", "SE", "SE", "SE", "SE", "SE", "S")
     Call Populate_the_pacification_Back_color("GL8_18_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "SE", "SE", "SE", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_19_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "SE", "SE", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_20_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "SE", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_21_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "SE", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_22_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "SE", "SE")
     Call Populate_the_pacification_Back_color("GL8_23_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "S", "SE")
     Call Populate_the_pacification_Back_color("GL8_24_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "S", "S")
     Call Populate_the_pacification_Back_color("GL8_25_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "S", "S", "S", "S", "S", "S", "S", "SW")
     Call Populate_the_pacification_Back_color("GL8_26_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "S", "S", "S", "S", "S", "S")
     Call Populate_the_pacification_Back_color("GL8_27_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "S", "S", "S", "S", "S")
     Call Populate_the_pacification_Back_color("GL8_28_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "S", "S", "S", "S")
     Call Populate_the_pacification_Back_color("GL8_29_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "S", "S", "S")
     Call Populate_the_pacification_Back_color("GL8_30_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "S", "S")
     Call Populate_the_pacification_Back_color("GL8_31_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "SW", "S")
     Call Populate_the_pacification_Back_color("GL8_32_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "SW", "SW")
     Call Populate_the_pacification_Back_color("GL8_33_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "SW", "SW", "SW", "SW", "SW", "SW", "SW", "NW")
     Call Populate_the_pacification_Back_color("GL8_34_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "SW", "SW", "SW", "SW", "SW", "SW")
     Call Populate_the_pacification_Back_color("GL8_35_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "SW", "SW", "SW", "SW", "SW")
     Call Populate_the_pacification_Back_color("GL8_36_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "SW", "SW", "SW", "SW")
     Call Populate_the_pacification_Back_color("GL8_37_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "SW", "SW", "SW")
     Call Populate_the_pacification_Back_color("GL8_38_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "SW", "SW")
     Call Populate_the_pacification_Back_color("GL8_39_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "NW", "SW")
     Call Populate_the_pacification_Back_color("GL8_40_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "NW", "NW", "NW", "NW", "NW", "NW", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_41_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "NW", "NW", "NW", "NW", "NW", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_42_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "NW", "NW", "NW", "NW", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_43_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "NW", "NW", "NW", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_44_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "NW", "NW", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_45_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "NW", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_46_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "NW", "NW")
     Call Populate_the_pacification_Back_color("GL8_47_hex", CURRENT_TERRAIN)
     Call Get_HexMAP_and_Terrain_of_a_hex(Tribes_Current_Hex, "N", "N", "N", "N", "N", "N", "N", "NW")
     Call Populate_the_pacification_Back_color("GL8_48_hex", CURRENT_TERRAIN)
     fullfield = CStr("GL8_" & count)
     Do
         MYFORM(fullfield).Visible = True
         MYFORM(fullfield) = PACIFICATION_TABLE(fullfield)
         count = count + 1
         fullfield = CStr("GL8_" & count)
         If count > 48 Then
            Exit Do
         End If
     Loop
End If
         
hex_pop = HEX_POPULATION(CLAN, TRIBE, Tribes_Current_Hex)
         
MYFORM![HEX_MAP_POP] = hex_pop

PACIFICATION_TABLE.Close
TRIBESINFO.Close

ERR_Populate_the_Pacification_Screen_Exit_CLOSE:
   Exit Function

ERR_Populate_the_Pacification_Screen_Exit:
If (Err = 3420) Then
   Resume Next
   
Else
  Resume ERR_Populate_the_Pacification_Screen_Exit_CLOSE
End If

End Function

Public Function Populate_the_pacification_Back_color(field, TERRAIN)
Dim lngBlue, lngGreen, lngGrey, lngRed, lngBlack, lngYellow, lngWhite, lngSwamp As Long
Dim fullfield As String

lngBlue = RGB(0, 0, 255)
lngGreen = RGB(0, 200, 0)
lngGrey = RGB(50, 90, 80)
lngRed = RGB(255, 0, 0)
lngBlack = RGB(0, 0, 0)
lngYellow = RGB(255, 255, 0)
lngWhite = RGB(255, 255, 255)
lngSwamp = RGB(190, 100, 140)

fullfield = CStr(field)

If InStr(TERRAIN, "PRAIRIE") Then
    MYFORM(fullfield).BackColor = lngYellow
ElseIf InStr(TERRAIN, "HILL") Then
    MYFORM(fullfield).BackColor = lngGreen
ElseIf InStr(TERRAIN, "FOREST") Then
    MYFORM(fullfield).BackColor = lngGreen
ElseIf InStr(TERRAIN, "SWAMP") Then
    MYFORM(fullfield).BackColor = lngSwamp
ElseIf InStr(TERRAIN, "OCEAN") Then
    MYFORM(fullfield).BackColor = lngBlue
ElseIf InStr(TERRAIN, "LAKE") Then
    MYFORM(fullfield).BackColor = lngBlue
ElseIf InStr(TERRAIN, "ALPS") Then
    MYFORM(fullfield).BackColor = lngGrey
ElseIf InStr(TERRAIN, "MOUNTAIN") Then
    MYFORM(fullfield).BackColor = lngGrey
Else
    MYFORM(fullfield).BackColor = lngYellow
End If
     
MYFORM(fullfield).Visible = True

End Function

Public Function Update_activity_Form(Field_Number)
Dim Number_of_Actives_Assigned As Long
Dim Initial_Field_Number As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![TURNS ACTIVITIES]

Set MODTABLE = TVDBGM.OpenRecordset("MODIFIERS")
MODTABLE.index = "PRIMARYKEY"

count = 1

Number_of_Actives_Assigned = MYFORM![ACTIVES]
Initial_Field_Number = Field_Number

Do Until count >= 14
      If count < 10 Then
         stext1 = "ITEM_0" & CStr(count)
         stext2 = "USE_AMT_0" & CStr(count)
      Else
         stext1 = "ITEM_" & CStr(count)
         stext2 = "USE_AMT_" & CStr(count)
      End If
    
      If IsNull(MYFORM(stext1).Value) Then
         Exit Do
      ElseIf IsNull(MYFORM(stext2).Value) Then
         Exit Do
      End If
      
      sValue1 = MYFORM(stext1).Value
      sValue2 = MYFORM(stext2).Value
      
      If IsNull(sValue2) Then
         Exit Do
      End If

      If count < Field_Number Then
          If MYFORM![ACTIVITY01] = "hunting" Or MYFORM![ACTIVITY01] = "furrier" Then
              If sValue1 = "TRAP" Then
                  MODTABLE.MoveFirst
                  MODTABLE.Seek "=", MYFORM![TRIBENUMBER], "TRAPS"
    
                  If MODTABLE.NoMatch Then
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / 5)
                  Else
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / MODTABLE![AMOUNT])
                  End If
              ElseIf sValue1 = "snare" Then
                  MODTABLE.MoveFirst
                  MODTABLE.Seek "=", MYFORM![TRIBENUMBER], "SNARES"
        
                  If MODTABLE.NoMatch Then
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / 5)
                  Else
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / MODTABLE![AMOUNT])
                  End If
              ElseIf Number_of_Actives_Assigned = 0 Then
                  MYFORM(stext2).Value = 0
                  MYFORM(stext1).Visible = False
                  MYFORM(stext2).Visible = False
              Else
                  Number_of_Actives_Assigned = Number_of_Actives_Assigned - sValue2
              End If
           End If
      Else
          If MYFORM![ACTIVITY01] = "hunting" Or MYFORM![ACTIVITY01] = "furrier" Then
             If Number_of_Actives_Assigned <= 0 Then
                 MYFORM(stext2).Value = 0
                 MYFORM(stext1).Visible = False
                 MYFORM(stext2).Visible = False
                 Field_Number = 14
              ElseIf sValue1 = "trap" Then
                  MODTABLE.MoveFirst
                  MODTABLE.Seek "=", MYFORM![TRIBENUMBER], "TRAPS"
    
                  If MODTABLE.NoMatch Then
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / 5)
                  Else
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / MODTABLE![AMOUNT])
                  End If
              ElseIf sValue1 = "snare" Then
                  MODTABLE.MoveFirst
                  MODTABLE.Seek "=", MYFORM![TRIBENUMBER], "SNARES"
        
                  If MODTABLE.NoMatch Then
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / 5)
                  Else
                      Number_of_Actives_Assigned = Number_of_Actives_Assigned - (sValue2 / MODTABLE![AMOUNT])
                  End If
              Else
                  Number_of_Actives_Assigned = Number_of_Actives_Assigned - sValue2
              End If
          End If
      End If
      
      If count >= Field_Number Then
          If MYFORM![ACTIVITY01] = "hunting" Or MYFORM![ACTIVITY01] = "furrier" Then
              If Number_of_Actives_Assigned > 0 Then
                 Exit Do
              Else
                  Field_Number = 14
             End If
         End If
     End If
     count = count + 1
Loop

If count > Initial_Field_Number Then
    If MYFORM![ACTIVITY01] = "hunting" Or MYFORM![ACTIVITY01] = "furrier" Then
        MYFORM![PROCESS].SetFocus
    End If
End If

CLOSE_Update_activity_form:

End Function
