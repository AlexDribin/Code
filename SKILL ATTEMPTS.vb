Attribute VB_Name = "SKILL ATTEMPTS"
Option Compare Database   'Use database order for string comparisons
Option Explicit

Global CURRENT_GM As String


'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 23/02/96  *  Remove some of the extra dice rolls                            **'
'** 29/06/96  *  Modify Skill Attempts to record attempts                       **'
'** 25/01/17  *  Include Libraries and Universities                             **'
'** 18/12/18  *  Updates to include calla to additional functions               **'
'** 16/01/20  *  Addressed issue with missing skill attempts                    **'
'*===============================================================================*'
 

Function SKILL_1(SCREEN As String)
On Error GoTo ERR_SKILLS

If SCREEN = "YES" Then
  If IsNull(Forms![SKILLS_1]![CLAN NAME]) Then
     Exit Function
  End If
End If
Dim trtab As Recordset           ' TRIBES_GENERAL_INFO
Dim mdtab As Recordset           ' MODIFIERS
Dim crtab As Recordset           ' COMPLETED_RESEARCH
Dim sktab As Recordset           ' SKILLS
Dim vstab As Recordset           ' VALID_SKILLS
Dim rrtab As Recordset           ' TRIBE_RESEARCH
Dim drtab As Recordset           ' DICE_ROLLS
Dim newtab As Recordset          ' RESEARCH
Dim skilltab As Recordset        ' PROCESS_SKILLS
Dim researchtab As Recordset     ' PROCESS_RESEARCH
Dim TRIBESGOODS As Recordset
Dim Skill_Attempts As Recordset
Dim Research_Attempts As Recordset

Dim TVWKSPACE As Workspace       '

Dim CLAN As String
Dim TRIBE As String
Dim errmsg As String
Dim skprimary As String
Dim sksecond As String
Dim sktertiary As String
Dim skgroup1 As String
Dim skgroup2 As String
Dim skgroup3 As String
Dim skmorale As String
Dim skship1 As String
Dim skship2 As String
Dim skship3 As String
Dim crlf As String, wks0 As String, wks1 As String, WKS2 As String
Dim wks3 As String, wks4 As String, wks5 As String, wks6 As String
Dim wks7 As String, wks8 As String, wks9 As String, wks10 As String
Dim wks11 As String, wks12 As String, wks13 As String, wks14 As String
Dim wks15 As String, wks16 As String, wks17 As String, wks18 As String
Dim wks19 As String, wks20 As String, wks21 As String, wks22 As String
Dim wks23 As String, wks24 As String, wks25 As String, wks26 As String
Dim wks27 As String, wks28 As String, wks29 As String, wks30 As String
Dim wks31 As String, wks32 As String, wks33 As String, wks34 As String
Dim wks35 As String, wks36 As String, wks37 As String, wks38 As String
Dim wks39 As String, wks40 As String

Dim roll1 As Long
Dim roll2 As Long
Dim roll3 As Long
Dim cnt1 As Long
Dim cnt2 As Long
Dim cnt3 As Long
Dim sklevel As Long
Dim skok As Long
Dim skcreate As Long
Dim skmod1 As Long
Dim skmod2 As Long
Dim skmod3 As Long
Dim resmod As Long
Dim tmpdl As Long
Dim POSITION As Long
Dim WORDLEN As Long
Dim codetrack As Long
Dim DICE_TRIBE As Long
Dim restopmod As Long
Dim resdevmod As Long
Dim oldrestopmod As Long
Dim oldresdevmod As Long
Dim LINENUMBER As Long
Dim Process_Tertiary As String
Dim Process_Teacher As String
Dim MAP_REFERENCE As String
Dim LITERACY_LEVEL As Long
Dim lib_found As String
Dim uni_found As String
Dim allowed_research_attempts As Long
Dim Research_YN As String
Dim TOTALSILVER As Long
Dim Silver_Spent As Long
Dim PRIMARY_YN As String
Dim SECONDARY_YN As String
Dim TERTIARY_YN As String
Dim Skill_Being_Attempted As String
Dim Research_Being_Attempted As String

Static rtopic(40) As String
Static dlreq(20) As Long
Static newres(40) As String
Static newdlr(40) As Long
Static wkres(40) As String
Static wkdlcur(40) As Long
Static wkdlreq(40) As Long
Static wkchg(40) As Long
Static wkcost(40) As Long
Static rdl(40) As Long
Dim Costs(8) As Long
Dim COST_CLAN(999) As Long
Dim research_only As Boolean

Dim Whole_number As Integer
Dim Decimal_number As Double

Section = "Start"
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

CURRENT_GM = GMTABLE![Name]

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
If SCREEN = "YES" Then
   Set MYFORM = Forms![SKILLS_1]
End If

Set trtab = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
Set mdtab = TVDBGM.OpenRecordset("MODIFIERS")
Set sktab = TVDBGM.OpenRecordset("SKILLS")
Set vstab = TVDB.OpenRecordset("VALID_SKILLS")
Set drtab = TVDBGM.OpenRecordset("DICE_ROLLS")
Set newtab = TVDB.OpenRecordset("RESEARCH")
newtab.index = "TOPIC"
Set crtab = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
crtab.index = "PRIMARYKEY"
Set rrtab = TVDBGM.OpenRecordset("TRIBE_RESEARCH")
rrtab.index = "PRIMARYKEY"
Set Skill_Attempts = TVDBGM.OpenRecordset("SKILL_ATTEMPTS")
Skill_Attempts.index = "PRIMARYKEY"
Skill_Attempts.MoveFirst
Set Research_Attempts = TVDBGM.OpenRecordset("RESEARCH_ATTEMPTS")
Research_Attempts.index = "PRIMARYKEY"
Research_Attempts.MoveFirst
Set SkillsTab = TVDBGM.OpenRecordset("Process_Skills")
SkillsTab.index = "PRIMARYKEY"
SkillsTab.MoveFirst
Set researchtab = TVDBGM.OpenRecordset("Process_Research")
researchtab.index = "TRIBE"
researchtab.MoveFirst

research_only = False

Section = "Begin Skills Table Loop"
Do Until SkillsTab.EOF


Process_Tertiary = "NO"
Process_Teacher = "NO"
Research_YN = "NO"
Silver_Spent = 0
Costs(1) = 1
Costs(2) = 2
Costs(3) = 4
Costs(4) = 7
Costs(5) = 11
Costs(6) = 16
Costs(7) = 22
Costs(8) = 29

count = 1
Do
  COST_CLAN(count) = 0
  count = count + 1
  If count > 999 Then
     Exit Do
  End If
Loop

If research_only = True Then
   'find next non processed research
   Do
     If researchtab![PROCESSED] = "Y" Then
        researchtab.MoveNext
     Else
        CLAN = "0" & Mid(researchtab![TRIBE], 2, 3)
        TRIBE = researchtab![TRIBE]
        GoTo Start_Research
     End If
     If researchtab.EOF Then
        GoTo End_Loop
     End If
   Loop

End If

If SCREEN = "YES" Then
   CLAN = "0" & Mid(MYFORM![TRIBE NAME], 2, 3)
   TRIBE = MYFORM![TRIBE NAME]
   skprimary = MYFORM![PRIMARY SKILL ATTEMPT]
   sksecond = MYFORM![SECONDARY SKILL ATTEMPT]
   sktertiary = MYFORM![TERTIARY SKILL ATTEMPT]
   If sktertiary = "EMPTY" Then
      Process_Tertiary = "NO"
   Else
      Process_Tertiary = "YES"
   End If
   If MYFORM![RESEARCH TOPIC 1] = "EMPTY" Then
      Research_YN = "NO"
   Else
      Research_YN = "YES"
   End If
Else
   CLAN = "0" & Mid(SkillsTab![TRIBE], 2, 3)
   TRIBE = SkillsTab![TRIBE]
   ' verify skill
   If SkillsTab![Order] = 1 Then
      If IsNull(SkillsTab![TOPIC]) Then
         skprimary = "EMPTY"
      Else
         skprimary = SkillsTab![TOPIC]
      End If
      If SkillsTab![PROCESSED] = "Y" Then
         PRIMARY_YN = "Y"
      Else
          PRIMARY_YN = "N"
         SkillsTab.Edit
         SkillsTab![PROCESSED] = "Y"
         SkillsTab.UPDATE
      End If
      SkillsTab.MoveNext
   Else
      ' No first skill
   End If
   If SkillsTab![Order] = 2 Then
      If IsNull(SkillsTab![TOPIC]) Then
         sksecond = "EMPTY"
      Else
         sksecond = SkillsTab![TOPIC]
      End If
      If SkillsTab![PROCESSED] = "Y" Then
         SECONDARY_YN = "Y"
      Else
         SECONDARY_YN = "N"
         SkillsTab.Edit
         SkillsTab![PROCESSED] = "Y"
         SkillsTab.UPDATE
      End If
      SkillsTab.MoveNext
   Else
      ' No second skill
   End If
   If SkillsTab![TRIBE] <> TRIBE Then
      Process_Tertiary = "NO"
      SkillsTab.MovePrevious
      sktertiary = "EMPTY"
   Else
      If IsNull(SkillsTab![TOPIC]) Then
         sktertiary = "EMPTY"
      Else
         sktertiary = SkillsTab![TOPIC]
      End If
      If SkillsTab![PROCESSED] = "Y" Then
         TERTIARY_YN = "Y"
      Else
         TERTIARY_YN = "N"
         Process_Tertiary = "YES"
         SkillsTab.Edit
         SkillsTab![PROCESSED] = "Y"
         SkillsTab.UPDATE
      End If
   End If

End If

codetrack = 0
crlf = Chr(13) & Chr(10)

If codetrack = 1 Then
    wks0 = "Skill Attempts" & crlf & crlf
    wks1 = "Primary  : " & skprimary & crlf
    WKS2 = "Secondary: " & sksecond
    Response = MsgBox((wks0 & wks1 & WKS2), True)
End If

DICE_TRIBE = Unit_Check("DICE", TRIBE)
TRIBE = Unit_Check("TRIBE", TRIBE)

trtab.index = "PRIMARYKEY"
trtab.Seek "=", CLAN, TRIBE
If trtab.NoMatch Then
    errmsg = "Unit " & CLAN & "-" & TRIBE & " Does not exist" & Chr(13) & Chr(10)
    errmsg = errmsg & "Skill attempt failed"
    Response = MsgBox(errmsg, True)
    GoTo End_Loop
End If

If Not IsNull(trtab![GOODS TRIBE]) Then
   GOODS_TRIBE = trtab![GOODS TRIBE]
Else
   GOODS_TRIBE = TRIBE
End If

MAP_REFERENCE = trtab![Current Hex]

Section = "Get Modifiers"
' Determine skill attempt modifiers

mdtab.index = "PRIMARYKEY"
mdtab.Seek "=", TRIBE, "PRIMARY SKILL ATTEMPT"

If mdtab.NoMatch Then
   skmod1 = 0
Else
   skmod1 = mdtab![AMOUNT]
End If

skmod1 = skmod1 + 10

Skill_Attempts.Seek "=", CLAN, TRIBE, skprimary
If Skill_Attempts.NoMatch Then
   'no change
ElseIf Skill_Attempts![ATTEMPTS] >= 11 Then
   skmod1 = 100
Else
   skmod1 = skmod1 + Skill_Attempts![ATTEMPTS]
End If

mdtab.index = "PRIMARYKEY"
mdtab.Seek "=", TRIBE, "SECONDARY SKILL ATTEMPT"

If mdtab.NoMatch Then
   skmod2 = 0
Else
   skmod2 = mdtab![AMOUNT]
End If

skmod2 = skmod2 + 10

Skill_Attempts.Seek "=", CLAN, TRIBE, sksecond
If Skill_Attempts.NoMatch Then
   'no change
ElseIf Skill_Attempts![ATTEMPTS] >= 11 Then
   skmod2 = 100
Else
   skmod2 = skmod2 + Skill_Attempts![ATTEMPTS]
End If

mdtab.index = "PRIMARYKEY"
mdtab.Seek "=", TRIBE, "TERTIARY SKILL ATTEMPT"

If mdtab.NoMatch Then
   skmod3 = 0
Else
   skmod3 = mdtab![AMOUNT]
End If

skmod3 = skmod3 + 10

Skill_Attempts.Seek "=", CLAN, TRIBE, sktertiary
If Skill_Attempts.NoMatch Then
   'no change
ElseIf Skill_Attempts![ATTEMPTS] >= 11 Then
   skmod3 = 100
Else
   skmod3 = skmod3 + Skill_Attempts![ATTEMPTS]
End If

' Determine Research modifiers - not included in current mandate

mdtab.index = "PRIMARYKEY"
mdtab.Seek "=", TRIBE, "RESEARCH TOPIC"

If mdtab.NoMatch Then
   restopmod = 0
Else
   restopmod = mdtab![AMOUNT]
End If
mdtab.index = "PRIMARYKEY"
mdtab.Seek "=", TRIBE, "RESEARCH ATTEMPT"

If mdtab.NoMatch Then
   resdevmod = 0
Else
   resdevmod = mdtab![AMOUNT]
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT 1"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT 2"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT 3"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT 4"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT 5"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH DEVELOPMENT 6"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "LIBRARIAN"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 10
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "HEAD LIBRARIAN"

If Not crtab.NoMatch Then
    resdevmod = resdevmod + 10
End If

' Determine topic modifiers - not included in current mandate

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED REASEARCH TOPIC ATEMPTS 1"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS 1"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS 2"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS 3"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS 4"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS 5"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

crtab.MoveFirst
crtab.Seek "=", TRIBE, "INCREASED RESEARCH TOPIC ATTEMPTS 6"

If Not crtab.NoMatch Then
    restopmod = restopmod + 5
End If

' TESTING MODIFIERS
'If CLAN = "030" Or CLAN = "0330" Then
'   skmod1 = skmod1 + 20
'   skmod2 = skmod2 + 20
'   skmod3 = skmod3 + 20
'   restopmod = restopmod + 20
'   resdevmod = resdevmod + 15
'End If

skmorale = "N"
skgroup1 = "Z"
skgroup2 = "Y"
skgroup3 = "Y"
skok = 0
skcreate = 0
skship1 = "N"
skship2 = "N"
skship3 = "N"
Skill_Being_Attempted = "None"
Research_Being_Attempted = "None"

If codetrack = 1 Then
    wks0 = "Primary - Before Setup" & crlf & crlf
    wks1 = "Skill Name: " & skprimary & crlf
    WKS2 = "Modifier  : " & skmod1 & crlf
    wks3 = "Morale    : " & skmorale & crlf
    wks4 = "Group ID  : " & skgroup1 & crlf
    wks5 = "Naval Flag: " & skship1 & crlf
    wks6 = "Control   : " & skok & crlf & "Setup Flag: " & skcreate
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
End If

If skprimary = sksecond Then
    sksecond = "EMPTY"
End If

If skprimary = sktertiary Then
    sktertiary = "EMPTY"
    Process_Tertiary = "NO"
End If

If sksecond = sktertiary Then
    sktertiary = "EMPTY"
    Process_Tertiary = "NO"
End If

sktab.index = "PRIMARYKEY"
vstab.index = "PRIMARYKEY"
sktab.MoveFirst
vstab.MoveFirst

skprimary_start:

Section = "Primary Skill Attempt"
Skill_Being_Attempted = skprimary
If Not skprimary = "EMPTY" Then
    If PRIMARY_YN = "Y" Then
        GoTo sksecond_start
    End If
    sktab.Seek "=", TRIBE, skprimary
    If sktab.NoMatch Then
        sklevel = 0
        skcreate = 1
    ElseIf sktab![SKILL LEVEL] >= 10 Then
        ' attempting to increase skill above 10
        GoTo sksecond_start
    Else
        sktab.Edit
        sklevel = sktab![SKILL LEVEL]
        skcreate = 0
    End If
    vstab.Seek "=", skprimary
    If vstab.NoMatch Then
        ' needs an update to a comment somewhere to show why the skill was not attempted
        skprimary = "EMPTY"
        skok = 0
        skmorale = "N"
        skgroup1 = "N"
        skship1 = "N"
        GoTo sksecond_start
    Else
        skok = 1
        skmorale = vstab!MORALE
        skgroup1 = vstab!Group
        skship1 = vstab!SHIP
    End If
Else
    GoTo sksecond_start
End If

If codetrack = 1 Then
    wks0 = "Primary - After Setup" & crlf & crlf
    wks1 = "Skill Name: " & skprimary & " (" & sklevel & ")" & crlf
    WKS2 = "Modifier  : " & skmod1 & crlf
    wks3 = "Morale    : " & skmorale & crlf
    wks4 = "Group ID  : " & skgroup1 & crlf
    wks5 = "Naval Flag: " & skship1 & crlf
    wks6 = "Control   : " & skok & crlf & "Setup Flag: " & skcreate
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
End If

'For cnt1 = 1 To 10
'    roll1 = DROLL(6, sklevel, 100, 0, DICE_TRIBE, 0, 0)
'Next


If skok = 1 And sklevel < 10 Then
    roll1 = DROLL(6, sklevel, 100, 0, DICE_TRIBE, 0, 0)
    drtab.AddNew
    drtab![CLAN] = CLAN
    drtab![TRIBE] = TRIBE
    drtab![roll] = roll1
    drtab![Skill] = skprimary
    drtab![level] = sklevel + 1
    drtab.UPDATE
    
    If codetrack = 1 Then
        Response = MsgBox(("Prime roll: " & roll1), True)
    End If
Section = "update sktab for primary skill"
    If roll1 <= (110 - ((sklevel + 1) * 10) + skmod1) Then
        If skcreate = 1 Then
            sktab.AddNew
            sktab!TRIBE = TRIBE
            sktab!Skill = skprimary
            sktab![SKILL LEVEL] = sklevel + 1
            sktab![SUCCESSFUL] = "Y"
            sktab![ATTEMPTED] = "Y"
            sktab.UPDATE
        Else
            sktab.Edit
            sktab![SKILL LEVEL] = sktab![SKILL LEVEL] + 1
            sktab![SUCCESSFUL] = "Y"
            sktab![ATTEMPTED] = "Y"
            sktab.UPDATE
        End If
        Skill_Attempts.Seek "=", CLAN, TRIBE, skprimary
        If Not Skill_Attempts.NoMatch Then
           Skill_Attempts.Delete
        End If
        If codetrack = 1 Then
            wks0 = "Primary Successful" & crlf & crlf
            wks1 = "Skill: " & skprimary & crlf
            WKS2 = "Level: " & (sklevel + 1) & crlf
            Response = MsgBox((wks0 & wks1 & WKS2), True)
        End If
        If skmorale = "Y" Then
            Call CHECK_MORALE(CLAN, TRIBE)
            If codetrack = 1 Then
                Response = MsgBox(("Primary Morale Check!"), True)
            End If
       End If
    ElseIf skcreate = 1 Then
 Section = "update sktab for primary skill unsuccessful"
       sktab.AddNew
        sktab!TRIBE = TRIBE
        sktab!Skill = skprimary
        sktab![SKILL LEVEL] = sklevel
        sktab![SUCCESSFUL] = "N"
        sktab![ATTEMPTED] = "Y"
        sktab.UPDATE
        Skill_Attempts.Seek "=", CLAN, TRIBE, skprimary
        If Skill_Attempts.NoMatch Then
           Skill_Attempts.AddNew
           Skill_Attempts!CLAN = CLAN
           Skill_Attempts!TRIBE = TRIBE
           Skill_Attempts!Skill = skprimary
           Skill_Attempts![ATTEMPTS] = 1
           Skill_Attempts.UPDATE
        Else
           Skill_Attempts.Edit
           Skill_Attempts![ATTEMPTS] = Skill_Attempts![ATTEMPTS] + 1
           Skill_Attempts.UPDATE
        End If
    Else
 Section = "update sktab for primary skill unsuccessful #2"
        Skill_Attempts.Seek "=", CLAN, TRIBE, skprimary
        If Skill_Attempts.NoMatch Then
           Skill_Attempts.AddNew
           Skill_Attempts!CLAN = CLAN
           Skill_Attempts!TRIBE = TRIBE
           Skill_Attempts!Skill = skprimary
           Skill_Attempts![ATTEMPTS] = 1
           Skill_Attempts.UPDATE
        Else
           Skill_Attempts.Edit
           Skill_Attempts![ATTEMPTS] = Skill_Attempts![ATTEMPTS] + 1
           Skill_Attempts.UPDATE
        End If
        sktab.Edit
        sktab![SUCCESSFUL] = "N"
        sktab![ATTEMPTED] = "Y"
        sktab.UPDATE
    End If
End If

sksecond_start:

Section = "Secondary Skill Attempt"

skok = 0
skcreate = 0
skmorale = "N"
Skill_Being_Attempted = sksecond

If codetrack = 1 Then
    wks0 = "Secondary - Before Setup" & crlf & crlf
    wks1 = "Skill Name: " & sksecond & crlf
    WKS2 = "Modifier  : " & skmod2 & crlf
    wks3 = "Morale    : " & skmorale & crlf
    wks4 = "Group ID  : " & skgroup2 & crlf
    wks5 = "Naval Flag: " & skship2 & crlf
    wks6 = "Control   : " & skok & crlf & "Setup Flag: " & skcreate
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
End If

If Not sksecond = "EMPTY" Then
    If SECONDARY_YN = "Y" Then
        GoTo sktertiary_start
    End If
    sktab.Seek "=", TRIBE, sksecond
    If sktab.NoMatch Then
        sklevel = 0
        skcreate = 1
    ElseIf sktab![SKILL LEVEL] >= 10 Then
        ' attempting to increase skill above 10
        GoTo sktertiary_start
    Else
        skcreate = 0
        sktab.Edit
        sklevel = sktab![SKILL LEVEL]
    End If
    vstab.Seek "=", sksecond
    If vstab.NoMatch Then
        ' needs an update to a comment somewhere to show why the skill was not attempted
        sksecond = "EMPTY"
        skok = 0
        skmorale = "N"
        skgroup2 = "N"
        skship2 = "N"
        GoTo sktertiary_start
    Else
        skok = 1
        skmorale = vstab!MORALE
        skgroup2 = vstab!Group
        skship2 = vstab!SHIP
    End If
Else
    GoTo sktertiary_start
End If

If codetrack = 1 Then
    wks0 = "Secondary - After Setup" & crlf & crlf
    wks1 = "Skill Name: " & sksecond & " (" & sklevel & ")" & crlf
    WKS2 = "Modifier  : " & skmod2 & crlf
    wks3 = "Morale    : " & skmorale & crlf
    wks4 = "Group ID  : " & skgroup2 & crlf
    wks5 = "Naval Flag: " & skship2 & crlf
    wks6 = "Control   : " & skok & crlf & "Setup Flag: " & skcreate
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
End If

Section = "Perform Secondary Skill Attempt"
' Perform Skill Attempts

If skok = 1 And sklevel < 10 Then
    roll3 = DROLL(6, sklevel, 100, 0, DICE_TRIBE, 0, 0)
    roll1 = ((110 - ((sklevel + 1) * 10)) / 2) + skmod2
    drtab.AddNew
    drtab![CLAN] = CLAN
    drtab![TRIBE] = TRIBE
    drtab![roll] = roll3
    drtab![Skill] = sksecond
    drtab![level] = sklevel + 1
    drtab.UPDATE
    If codetrack = 1 Then
        Response = MsgBox(("Secondary roll: " & roll3), True)
        Response = MsgBox(("Roll Required: " & roll1), True)
    End If
    If skgroup1 = skgroup2 Then
        If codetrack = 1 Then
            Response = MsgBox("Matching Skill Groups", True)
        End If
        If skship1 = "N" Or skship2 = "N" Then
            If codetrack = 1 Then
                Response = MsgBox("Skills not ship skills", True)
            End If
            roll1 = roll1 / 2
        End If
    End If
    If roll3 <= roll1 Then
Section = "Update sktab for second skill if successful"
        If skcreate = 1 Then
            sktab.AddNew
            sktab!TRIBE = TRIBE
            sktab!Skill = sksecond
            sktab![SKILL LEVEL] = sklevel + 1
            sktab![SUCCESSFUL] = "Y"
            sktab![ATTEMPTED] = "Y"
            sktab.UPDATE
        Else
            sktab.Edit
            sktab![SKILL LEVEL] = sktab![SKILL LEVEL] + 1
            sktab![SUCCESSFUL] = "Y"
            sktab![ATTEMPTED] = "Y"
            sktab.UPDATE
        End If
        Skill_Attempts.Seek "=", CLAN, TRIBE, sksecond
        If Not Skill_Attempts.NoMatch Then
           Skill_Attempts.Delete
        End If
        If codetrack = 1 Then
            wks0 = "Secondary Successful" & crlf & crlf
            wks1 = "Skill: " & sksecond & crlf
            WKS2 = "Level: " & (sklevel + 1) & crlf
            Response = MsgBox((wks0 & wks1 & WKS2), True)
        End If
        If skmorale = "Y" Then
            If codetrack = 1 Then
                Response = MsgBox(("Secondary Morale Check!"), True)
            End If
            Call CHECK_MORALE(CLAN, TRIBE)
        End If
    ElseIf skcreate = 1 Then
 Section = "update sktab for second skill if unsuccessful"
       sktab.AddNew
        sktab!TRIBE = TRIBE
        sktab!Skill = sksecond
        sktab![SKILL LEVEL] = sklevel
        sktab![SUCCESSFUL] = "N"
        sktab![ATTEMPTED] = "Y"
        sktab.UPDATE
        Skill_Attempts.Seek "=", CLAN, TRIBE, sksecond
        If Skill_Attempts.NoMatch Then
           Skill_Attempts.AddNew
           Skill_Attempts!CLAN = CLAN
           Skill_Attempts!TRIBE = TRIBE
           Skill_Attempts!Skill = skprimary
           Skill_Attempts![ATTEMPTS] = 1
           Skill_Attempts.UPDATE
        Else
           Skill_Attempts.Edit
           Skill_Attempts![ATTEMPTS] = Skill_Attempts![ATTEMPTS] + 1
           Skill_Attempts.UPDATE
        End If
    Else
 Section = "update sktab for second skill if unsuccessful # 2"
        sktab.Edit
        sktab![SUCCESSFUL] = "N"
        sktab![ATTEMPTED] = "Y"
        sktab.UPDATE
        Skill_Attempts.Seek "=", CLAN, TRIBE, sksecond
        If Skill_Attempts.NoMatch Then
           Skill_Attempts.AddNew
           Skill_Attempts!CLAN = CLAN
           Skill_Attempts!TRIBE = TRIBE
           Skill_Attempts!Skill = sksecond
           Skill_Attempts![ATTEMPTS] = 1
           Skill_Attempts.UPDATE
        Else
           Skill_Attempts.Edit
           Skill_Attempts![ATTEMPTS] = Skill_Attempts![ATTEMPTS] + 1
           Skill_Attempts.UPDATE
        End If
    End If
End If

Section = "Perform Third Skill Attempt"
Skill_Being_Attempted = sktertiary
sktertiary_start:
If TERTIARY_YN = "Y" Then
   GoTo Start_Research
End If
sktab.Seek "=", TRIBE, sktertiary
If sktab.NoMatch Then
   sklevel = 0
Else
   sklevel = sktab![SKILL LEVEL]
End If

' remove silver if tertiary is used as this is a teacher thing.
' 300 silver * sklevel
 
If SCREEN = "YES" Then
   If Forms![SKILLS_1]![CheckTertiary] = True Then
      'ACTIVITY OTHER THAN TEACHER
      'NO SILVER TO BE REMOVED
      Process_Teacher = "NO"
   Else
      Process_Teacher = "YES"
   End If
Else
   Process_Teacher = "YES"
End If

Section = "Tertiary Subtract silver"
If Process_Teacher = "YES" Then
 If Not sktertiary = "EMPTY" Then
    ' is there enough silver
    ' If yes Then can PROCESS
    ' Else: cant
    
    NumGoods = GET_TRIBES_GOOD_QUANTITY(CLAN, GOODS_TRIBE, "SILVER")
    TOTALSILVER = ((sklevel + 1) * 300)
 
    If NumGoods >= TOTALSILVER Then
       Call UPDATE_TRIBES_GOODS_TABLES(CLAN, GOODS_TRIBE, "SILVER", "SUBTRACT", TOTALSILVER)
       Process_Tertiary = "YES"
    Else
       Process_Tertiary = "NO"
    End If
 Else
    Process_Tertiary = "NO"
 End If
End If

If Process_Tertiary = "YES" Then

   skok = 0
   skcreate = 0
   skmorale = "N"

   If codetrack = 1 Then
      wks0 = "Tertiary - Before Setup" & crlf & crlf
      wks1 = "Skill Name: " & sktertiary & crlf
      WKS2 = "Modifier  : " & skmod3 & crlf
      wks3 = "Morale    : " & skmorale & crlf
      wks4 = "Group ID  : " & skgroup3 & crlf
      wks5 = "Naval Flag: " & skship3 & crlf
      wks6 = "Control   : " & skok & crlf & "Setup Flag: " & skcreate
      Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
   End If

   sktab.Seek "=", TRIBE, sktertiary
   If sktab.NoMatch Then
      sklevel = 0
      skcreate = 1
   Else
      skcreate = 0
      sktab.Edit
      sklevel = sktab![SKILL LEVEL]
   End If
   vstab.Seek "=", sktertiary
   If vstab.NoMatch Then
      ' needs an update to a comment somewhere to show why the skill was not attempted
      sktertiary = "EMPTY"
      skok = 0
      skmorale = "N"
      skgroup2 = "N"
      skship2 = "N"
      GoTo Start_Research
   Else
      skok = 1
      skmorale = vstab!MORALE
      skgroup2 = vstab!Group
      skship2 = vstab!SHIP
   End If

   If codetrack = 1 Then
      wks0 = "Tertiary - After Setup" & crlf & crlf
      wks1 = "Skill Name: " & sktertiary & " (" & sklevel & ")" & crlf
      WKS2 = "Modifier  : " & skmod3 & crlf
      wks3 = "Morale    : " & skmorale & crlf
      wks4 = "Group ID  : " & skgroup3 & crlf
      wks5 = "Naval Flag: " & skship3 & crlf
      wks6 = "Control   : " & skok & crlf & "Setup Flag: " & skcreate
      Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
   End If

   Section = "Perform Third skill attempt"
   If skok = 1 And sklevel < 10 Then
      roll3 = DROLL(6, sklevel, 100, 0, DICE_TRIBE, 0, 0)
      roll1 = ((110 - ((sklevel + 1) * 10)) / 2) + skmod3
      drtab.AddNew
      drtab![CLAN] = CLAN
      drtab![TRIBE] = TRIBE
      drtab![roll] = roll3
      drtab![Skill] = sktertiary
      drtab![level] = sklevel + 1
      drtab.UPDATE
      If codetrack = 1 Then
         Response = MsgBox(("Tertiary roll: " & roll3), True)
         Response = MsgBox(("Roll Required: " & roll1), True)
      End If
      If skgroup1 = skgroup2 Then
         If codetrack = 1 Then
            Response = MsgBox("Matching Skill Groups", True)
         End If
         If skship1 = "N" Or skship2 = "N" Then
            If codetrack = 1 Then
               Response = MsgBox("Skills not ship skills", True)
            End If
            roll1 = roll1 / 2
         End If
      End If
Section = "Update sktab for third skill if successful"
      If roll3 <= roll1 Then
         If skcreate = 1 Then
            sktab.AddNew
            sktab!TRIBE = TRIBE
            sktab!Skill = sktertiary
            sktab![SKILL LEVEL] = sklevel + 1
            sktab![SUCCESSFUL] = "Y"
            sktab![ATTEMPTED] = "Y"
            sktab.UPDATE
         Else
            sktab.Edit
            sktab![SKILL LEVEL] = sktab![SKILL LEVEL] + 1
            sktab![SUCCESSFUL] = "Y"
            sktab![ATTEMPTED] = "Y"
            sktab.UPDATE
         End If
         Skill_Attempts.Seek "=", CLAN, TRIBE, sktertiary
         If Not Skill_Attempts.NoMatch Then
            Skill_Attempts.Delete
         End If
         If codetrack = 1 Then
            wks0 = "Tertiary Successful" & crlf & crlf
            wks1 = "Skill: " & sktertiary & crlf
            WKS2 = "Level: " & (sklevel + 1) & crlf
            Response = MsgBox((wks0 & wks1 & WKS2), True)
         End If
         If skmorale = "Y" Then
            If codetrack = 1 Then
               Response = MsgBox(("Tertiary Morale Check!"), True)
            End If
            Call CHECK_MORALE(CLAN, TRIBE)
         End If
     ElseIf skcreate = 1 Then
Section = "Update sktab for third skill if unsuccessful"
         sktab.AddNew
         sktab!TRIBE = TRIBE
         sktab!Skill = sktertiary
         sktab![SKILL LEVEL] = sklevel
         sktab![SUCCESSFUL] = "N"
         sktab![ATTEMPTED] = "Y"
         sktab.UPDATE
         Skill_Attempts.Seek "=", CLAN, TRIBE, sktertiary
         If Skill_Attempts.NoMatch Then
            Skill_Attempts.AddNew
            Skill_Attempts!CLAN = CLAN
            Skill_Attempts!TRIBE = TRIBE
            Skill_Attempts!Skill = sktertiary
            Skill_Attempts![ATTEMPTS] = 1
            Skill_Attempts.UPDATE
         Else
            Skill_Attempts.Edit
            Skill_Attempts![ATTEMPTS] = Skill_Attempts![ATTEMPTS] + 1
            Skill_Attempts.UPDATE
         End If
      Else
Section = "Update sktab for third skill if successful #2"
         sktab.Edit
         sktab![SUCCESSFUL] = "N"
         sktab![ATTEMPTED] = "Y"
         sktab.UPDATE
         Skill_Attempts.Seek "=", CLAN, TRIBE, sktertiary
         If Skill_Attempts.NoMatch Then
            Skill_Attempts.AddNew
            Skill_Attempts!CLAN = CLAN
            Skill_Attempts!TRIBE = TRIBE
            Skill_Attempts!Skill = sktertiary
            Skill_Attempts![ATTEMPTS] = 1
            Skill_Attempts.UPDATE
         Else
            Skill_Attempts.Edit
            Skill_Attempts![ATTEMPTS] = Skill_Attempts![ATTEMPTS] + 1
            Skill_Attempts.UPDATE
         End If
      End If
   End If
End If

Section = "Begin Research"
Start_Research:
'********************************************************************'
'* Research                                                         *'
'********************************************************************'

' need to check each topic and see if a book is available - if is then use literacy level
' to do calc - literacy * 5%, if library then increase chance by 50%
' Check for university and increase number of permitted topics by number of people - 1 topic per 500 people or part thereof
sktab.index = "PRIMARYKEY"
sktab.MoveFirst
sktab.Seek "=", TRIBE, "LITERACY"
If sktab.NoMatch Then
   LITERACY_LEVEL = 0
Else
   LITERACY_LEVEL = sktab![SKILL LEVEL]
End If

Set CONSTTABLE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
CONSTTABLE.index = "FORTHKEY"
CONSTTABLE.Seek "=", MAP_REFERENCE, CLAN, "LIBRARY"

If CONSTTABLE.NoMatch Then
    lib_found = "NO"
Else
    lib_found = "YES"
End If

CONSTTABLE.Seek "=", MAP_REFERENCE, CLAN, "UNIVERSITY"
If CONSTTABLE.NoMatch Then
    uni_found = "NO"
Else
    uni_found = "YES"
End If

CONSTTABLE.Close

' determine the number of allowed research topics
allowed_research_attempts = trtab![WARRIORS] + trtab![ACTIVES] + trtab![INACTIVES]

dlreq(0) = 5
dlreq(1) = 50
dlreq(2) = 40
dlreq(3) = 30
dlreq(4) = 25
dlreq(5) = 20
dlreq(6) = 20
dlreq(7) = 15
dlreq(8) = 10
dlreq(9) = 10
dlreq(10) = 5
dlreq(11) = 5
dlreq(12) = 5
dlreq(13) = 5
dlreq(14) = 5
dlreq(15) = 5
dlreq(16) = 5
dlreq(17) = 5
dlreq(18) = 3
dlreq(19) = 2
dlreq(20) = 1

SECTION_NAME = "RESEARCH ATTEMPTED"
LINENUMBER = 1
 
count = 1
Do Until count > 40
   newres(count) = "EMPTY"
   wkres(count) = "EMPTY"
   wkcost(count) = 0
   count = count + 1
Loop

'load research into newres stack

If SCREEN = "YES" Then
   newres(1) = MYFORM![RESEARCH TOPIC 1]
   newres(2) = MYFORM![RESEARCH TOPIC 2]
   newres(3) = MYFORM![RESEARCH TOPIC 3]
   newres(4) = MYFORM![RESEARCH TOPIC 4]
   newres(5) = MYFORM![RESEARCH TOPIC 5]
   newres(6) = MYFORM![RESEARCH TOPIC 6]
   newres(7) = MYFORM![RESEARCH TOPIC 7]
   newres(8) = MYFORM![RESEARCH TOPIC 8]
   newres(9) = MYFORM![RESEARCH TOPIC 9]
   newres(10) = MYFORM![RESEARCH TOPIC 10]
   newres(11) = MYFORM![RESEARCH TOPIC 11]
   newres(12) = MYFORM![RESEARCH TOPIC 12]
   newres(13) = MYFORM![RESEARCH TOPIC 13]
   newres(14) = MYFORM![RESEARCH TOPIC 14]
   newres(15) = MYFORM![RESEARCH TOPIC 15]
   newres(16) = MYFORM![RESEARCH TOPIC 16]
   newres(17) = MYFORM![RESEARCH TOPIC 17]
   newres(18) = MYFORM![RESEARCH TOPIC 18]
   newres(19) = MYFORM![RESEARCH TOPIC 19]
   newres(20) = MYFORM![RESEARCH TOPIC 20]
   newres(21) = MYFORM![RESEARCH TOPIC 21]
   newres(22) = MYFORM![RESEARCH TOPIC 22]
   newres(23) = MYFORM![RESEARCH TOPIC 23]
   newres(24) = MYFORM![RESEARCH TOPIC 24]
   newres(25) = MYFORM![RESEARCH TOPIC 25]
   newres(26) = MYFORM![RESEARCH TOPIC 26]
   newres(27) = MYFORM![RESEARCH TOPIC 27]
   newres(28) = MYFORM![RESEARCH TOPIC 28]
   newres(29) = MYFORM![RESEARCH TOPIC 29]
   newres(30) = MYFORM![RESEARCH TOPIC 30]
   newres(31) = MYFORM![RESEARCH TOPIC 31]
   newres(32) = MYFORM![RESEARCH TOPIC 32]
   newres(33) = MYFORM![RESEARCH TOPIC 33]
   newres(34) = MYFORM![RESEARCH TOPIC 34]
   newres(35) = MYFORM![RESEARCH TOPIC 35]
   newres(36) = MYFORM![RESEARCH TOPIC 36]
   newres(37) = MYFORM![RESEARCH TOPIC 37]
   newres(38) = MYFORM![RESEARCH TOPIC 38]
   newres(39) = MYFORM![RESEARCH TOPIC 39]
   newres(40) = MYFORM![RESEARCH TOPIC 40]
Else
   researchtab.Seek "=", TRIBE
   count = 1
   Do Until count > 40
      If researchtab.NoMatch Then
         'ignore
         newres(count) = "EMPTY"
         Exit Do
      ElseIf researchtab![TRIBE] <> TRIBE Then
         newres(count) = "EMPTY"
      ElseIf researchtab![PROCESSED] = "Y" Then
         newres(count) = "EMPTY"
         researchtab.MoveNext
      Else
         newres(count) = researchtab![TOPIC]
         researchtab.Edit
         researchtab![PROCESSED] = "Y"
         researchtab.UPDATE
         researchtab.MoveNext
         
      End If
      count = count + 1
      If researchtab.EOF Then
         Exit Do
      End If
      If researchtab![TRIBE] <> TRIBE Then
         Exit Do
      End If
   Loop
   If researchtab![TRIBE] <> TRIBE Then
      researchtab.MovePrevious
   End If
  
End If

If codetrack > 0 Then
    wks0 = "From Screen(Occ-newres)" & crlf
    wks1 = "1-" & newres(1) & crlf
    WKS2 = "2-" & newres(2) & crlf
    wks3 = "3-" & newres(3) & crlf
    wks4 = "4-" & newres(4) & crlf
    wks5 = "5-" & newres(5) & crlf
    wks6 = "6-" & newres(6) & crlf
    wks7 = "7-" & newres(7) & crlf
    wks8 = "8-" & newres(8) & crlf
    wks9 = "9-" & newres(9) & crlf
    wks10 = "10-" & newres(10) & crlf
    wks11 = "11-" & newres(11) & crlf
    wks12 = "12-" & newres(12) & crlf
    wks13 = "13-" & newres(13) & crlf
    wks14 = "14-" & newres(14) & crlf
    wks15 = "15-" & newres(15) & crlf
    wks16 = "16-" & newres(16) & crlf
    wks17 = "17-" & newres(17) & crlf
    wks18 = "18-" & newres(18) & crlf
    wks19 = "19-" & newres(19) & crlf
    wks20 = "20-" & newres(20) & crlf
    wks21 = "21-" & newres(21) & crlf
    wks22 = "22-" & newres(22) & crlf
    wks23 = "23-" & newres(23) & crlf
    wks24 = "24-" & newres(24) & crlf
    wks25 = "25-" & newres(25) & crlf
    wks26 = "26-" & newres(26) & crlf
    wks27 = "27-" & newres(27) & crlf
    wks28 = "28-" & newres(28) & crlf
    wks29 = "29-" & newres(29) & crlf
    wks30 = "30-" & newres(30) & crlf
    wks31 = "31-" & newres(31) & crlf
    wks32 = "32-" & newres(32) & crlf
    wks33 = "33-" & newres(33) & crlf
    wks34 = "34-" & newres(34) & crlf
    wks35 = "35-" & newres(35) & crlf
    wks36 = "36-" & newres(36) & crlf
    wks37 = "37-" & newres(37) & crlf
    wks38 = "38-" & newres(38) & crlf
    wks39 = "39-" & newres(39) & crlf
    wks40 = "40-" & newres(40) & crlf
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
    Response = MsgBox((wks7 & wks8 & wks9 & wks10 & wks11 & wks12), True)
    Response = MsgBox((wks13 & wks14 & wks15 & wks16 & wks17 & wks18), True)
    Response = MsgBox((wks19 & wks20 & wks21 & wks22 & wks23 & wks24), True)
    Response = MsgBox((wks25 & wks26 & wks27 & wks28 & wks29 & wks30), True)
    Response = MsgBox((wks31 & wks32 & wks33 & wks34 & wks35 & wks36), True)
    Response = MsgBox((wks37 & wks38 & wks39 & wks40), True)
End If

For cnt1 = 0 To 40
 wkchg(cnt1) = 0
Next

' verify the research is a valid research topic
count = 1
Do Until count > 40
   newtab.MoveFirst
   newtab.Seek "=", newres(count)
   If newtab.NoMatch Then
      newres(count) = "EMPTY"
   End If
   count = count + 1
Loop

' get dl levels obtained to date
' if no dl levels, the wkres set to empty

count = 1
Do Until count > 40
   rrtab.Seek "=", TRIBE, newres(count)
   If rrtab.NoMatch Then
      If newres(count) = "EMPTY" Then
         'ignore
      Else
         wkres(count) = "EMPTY"
         wkdlcur(count) = 0
         wkdlreq(count) = 0
      End If
   Else
      rrtab.Edit
      rrtab![RESEARCH ATTEMPTED] = "Y"
      rrtab.UPDATE
      wkres(count) = newres(count)
      wkdlcur(count) = rrtab![DL LEVEL ATTAINED]
      wkdlreq(count) = rrtab![DL LEVEL REQUIRED]
   End If
   count = count + 1
Loop

' should check for completed research topics
count = 1
Do Until count > 40
   crtab.MoveFirst
   crtab.Seek "=", TRIBE, newres(count)
   If crtab.NoMatch Then
      ' great - not repeating
   Else
      newres(count) = "EMPTY"
      wkres(count) = "EMPTY"
   End If
   count = count + 1
Loop

' after this, need to cater for the actual number of topics that can be researched.
' with library - no change
' with university will be the amount in allowed_research_attempts

'allowed_research_attempts = ??
If uni_found = "YES" Then
   Decimal_number = allowed_research_attempts / 500
   Whole_number = Decimal_number
   Decimal_number = Decimal_number - Whole_number
   If Decimal_number > 0 Then
      Whole_number = Whole_number + 1
   End If
   allowed_research_attempts = Whole_number
   If allowed_research_attempts <= 1 Then
      allowed_research_attempts = 2
   End If
Else
   allowed_research_attempts = 2
End If

' should check for sufficient silver if a uni is found
If uni_found = "YES" Then
   count = 1   ' to loop through all research topics
   cnt1 = 0    ' to identify when a research topic should be paid for
   Do Until count > 40
      If newres(count) = "EMPTY" Then
         'ignore
      Else
         cnt1 = cnt1 + 1
         If cnt1 > 2 Then
            COST_CLAN(CLAN) = COST_CLAN(CLAN) + 1
            NumGoods = GET_TRIBES_GOOD_QUANTITY(CLAN, GOODS_TRIBE, "SILVER")
            If cnt1 >= 3 Then
               If cnt1 - 2 > 8 Then
                  TOTALSILVER = Costs(8) * 200
               Else
                  TOTALSILVER = Costs(cnt1 - 2) * 200
               End If
               If TOTALSILVER > NumGoods Then
                  'OutLine = "you attempted research topic - " & newres(cnt1) & " but had insufficient silver"
                  OutLine = "you attempted research topic - " & newres(cnt1)
                  Call WRITE_TURN_ACTIVITY(CLAN, TRIBE, SECTION_NAME, LINENUMBER, OutLine, "No")
                  LINENUMBER = LINENUMBER + 1
                  ' no research attempt
                  COST_CLAN(CLAN) = COST_CLAN(CLAN) - 1
                  newres(count) = "EMPTY"
                  wkres(count) = "EMPTY"
               Else
                  ' delete silver and move to next topic
                  Call UPDATE_TRIBES_GOODS_TABLES(CLAN, GOODS_TRIBE, "SILVER", "SUBTRACT", TOTALSILVER)
               End If
            End If
         End If
      End If
      count = count + 1
   Loop
End If


If codetrack > 0 Then
    wks0 = "Before Research Roll(wkres-wkdlcur-wkdlreq)" & crlf
    wks1 = wkres(1) & "-" & wkdlcur(1) & "-" & wkdlreq(1) & crlf
    WKS2 = wkres(2) & "-" & wkdlcur(2) & "-" & wkdlreq(2) & crlf
    wks3 = wkres(3) & "-" & wkdlcur(3) & "-" & wkdlreq(3) & crlf
    wks4 = wkres(4) & "-" & wkdlcur(4) & "-" & wkdlreq(4) & crlf
    wks5 = wkres(5) & "-" & wkdlcur(5) & "-" & wkdlreq(5) & crlf
    wks6 = wkres(6) & "-" & wkdlcur(6) & "-" & wkdlreq(6)
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
End If
 
If allowed_research_attempts > 40 Then
   allowed_research_attempts = 40
End If

rrtab.index = "PRIMARYKEY"

For cnt1 = 1 To allowed_research_attempts

 Research_Being_Attempted = newres(cnt1)
 roll1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)

    drtab.AddNew
    drtab![CLAN] = CLAN
    drtab![TRIBE] = TRIBE
    drtab![roll] = roll1
    drtab![Skill] = "research attempt"
    drtab![level] = 0
    drtab.UPDATE
 
 rdl(cnt1) = roll1
 
 oldrestopmod = restopmod
 oldresdevmod = resdevmod
 
 Set TRIBESBOOKS = TVDBGM.OpenRecordset("TRIBES_BOOKS")
 TRIBESBOOKS.index = "PRIMARYKEY"
 TRIBESBOOKS.Seek "=", CLAN, TRIBE, newres(cnt1)
 
 If Not TRIBESBOOKS.NoMatch Then
    If lib_found = "yes" Then
       restopmod = ((LITERACY_LEVEL * 5) * 1.5) - 5
       resdevmod = ((LITERACY_LEVEL * 5) * 1.5) - 5
    Else
       restopmod = (LITERACY_LEVEL * 5) - 5
       resdevmod = (LITERACY_LEVEL * 5) - 5
    End If
 End If
 
 Research_Attempts.Seek "=", CLAN, TRIBE, newres(cnt1)
 If Research_Attempts.NoMatch Then
    'no change
 ElseIf Research_Attempts![ATTEMPTS] >= 11 Then
    restopmod = restopmod + 100
    resdevmod = resdevmod + 100
 Else
    restopmod = restopmod + Research_Attempts![ATTEMPTS]
    resdevmod = resdevmod + Research_Attempts![ATTEMPTS]
 End If

 If wkres(cnt1) = "EMPTY" Then
    If Not Left(newres(cnt1), 5) = "EMPTY" Then
       If roll1 <= (restopmod + 5) Then
          wkres(cnt1) = newres(cnt1)
          newtab.MoveFirst
          newtab.Seek "=", newres(cnt1)
          wkdlreq(cnt1) = newtab![DL REQUIRED]
          wkdlcur(cnt1) = 0
          wkchg(cnt1) = 1
          wkcost(cnt1) = TOTALSILVER
          wks0 = "Research Successfully Started!" & crlf & crlf
          wks1 = "Clan : " & CLAN & crlf
          WKS2 = "Tribe: " & TRIBE & crlf & crlf
          wks3 = "Topic: " & wkres(cnt1) & crlf & crlf
          wks4 = "Roll  : " & roll1
          Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4), True)
          Research_Attempts.Seek "=", CLAN, TRIBE, newres(cnt1)
          If Not Research_Attempts.NoMatch Then
             Research_Attempts.Delete
          End If
       Else
          ' WRITE RESEARCH ATTEMPTED
          'If TOTALSILVER = 0 Then
             OutLine = "you attempted research topic - " & newres(cnt1)
          'Else
          '   OutLine = "you attempted research topic - " & newres(cnt1) & " it cost " & TOTALSILVER & " silver"
          'End If
          Call WRITE_TURN_ACTIVITY(CLAN, TRIBE, SECTION_NAME, LINENUMBER, OutLine, "No")
          LINENUMBER = LINENUMBER + 1
          Research_Attempts.Seek "=", CLAN, TRIBE, newres(cnt1)
          If Research_Attempts.NoMatch Then
             Research_Attempts.AddNew
             Research_Attempts!CLAN = CLAN
             Research_Attempts!TRIBE = TRIBE
             Research_Attempts!research = newres(cnt1)
             Research_Attempts![ATTEMPTS] = 1
             Research_Attempts![Cost] = TOTALSILVER
             Research_Attempts.UPDATE
          Else
             Research_Attempts.Edit
             Research_Attempts![ATTEMPTS] = Research_Attempts![ATTEMPTS] + 1
             Research_Attempts![Cost] = TOTALSILVER
             Research_Attempts.UPDATE
          End If
       End If
    End If
Else
   If wkdlcur(cnt1) < 20 Then
      tmpdl = wkdlcur(cnt1) + 1
      If roll1 <= (dlreq(tmpdl) + resdevmod) Then
         wkdlcur(cnt1) = tmpdl
         wkchg(cnt1) = 1
         wkcost(cnt1) = TOTALSILVER
         wks0 = "DL Level Gained!" & crlf & crlf
         wks1 = "Clan : " & CLAN & crlf
         WKS2 = "Tribe: " & TRIBE & crlf & crlf
         wks3 = "Topic: " & wkres(cnt1) & crlf & crlf
         wks4 = "Roll  : " & roll1
         Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4), True)
         Research_Attempts.Seek "=", CLAN, TRIBE, wkres(cnt1)
         If Not Research_Attempts.NoMatch Then
            Research_Attempts.Delete
         End If
      Else
         wkcost(cnt1) = TOTALSILVER
         Research_Attempts.Seek "=", CLAN, TRIBE, wkres(cnt1)
         If Research_Attempts.NoMatch Then
            Research_Attempts.AddNew
            Research_Attempts!CLAN = CLAN
            Research_Attempts!TRIBE = TRIBE
            Research_Attempts!research = wkres(cnt1)
            Research_Attempts![ATTEMPTS] = 1
            Research_Attempts![Cost] = TOTALSILVER
            Research_Attempts.UPDATE
         Else
            Research_Attempts.Edit
            Research_Attempts![ATTEMPTS] = Research_Attempts![ATTEMPTS] + 1
            Research_Attempts![Cost] = TOTALSILVER
            Research_Attempts.UPDATE
         End If
      End If
   End If
   If wkdlcur(cnt1) >= wkdlreq(cnt1) Then
      POSITION = InStr(wkres(cnt1), "(")
      If POSITION > 0 Then
         WORDLEN = POSITION - 1
      Else
         WORDLEN = Len(wkres(cnt1))
      End If
      crtab.AddNew
      crtab!TRIBE = TRIBE
      crtab!TOPIC = (Mid(wkres(cnt1), 1, WORDLEN))
      crtab!COMPLETED_THIS_TURN = "Y"
      crtab.UPDATE
      rrtab.MoveFirst
      rrtab.Seek "=", TRIBE, wkres(cnt1)
      rrtab.Delete
       
       
 ' At this point, update modifiers
 ' Updates to Tribes_General_Info table
       
      If Mid(wkres(cnt1), 1, WORDLEN) = "GOVERNMENT LEVEL 1" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 1
         trtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "GOVERNMENT LEVEL 2" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 2
         trtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "GOVERNMENT LEVEL 3" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 3
         trtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "GOVERNMENT LEVEL 4" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 4
         trtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "GOVERNMENT LEVEL 5" Then
         trtab.Edit
         trtab![GOVT LEVEL] = 5
         trtab.UPDATE
      End If
                  
 ' Updates to Modifiers table
      
      If Mid(wkres(cnt1), 1, WORDLEN) = "TRAPPERS" Then
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "TRAPS"
         If mdtab.NoMatch Then
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "TRAPS"
            mdtab![AMOUNT] = 10
            mdtab.UPDATE
         Else
            mdtab.Edit
            mdtab![AMOUNT] = 10
            mdtab.UPDATE
         End If
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "SNARES"
         If mdtab.NoMatch Then
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "SNARES"
            mdtab![AMOUNT] = 10
            mdtab.UPDATE
         Else
            mdtab.Edit
            mdtab![AMOUNT] = 10
            mdtab.UPDATE
         End If
      End If
                  
      If InStr(wkres(cnt1), "STONES/person") Then
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "STONES QUARRIED"
         If mdtab.NoMatch Then
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "STONES QUARRIED"
            mdtab![AMOUNT] = 5
            mdtab.UPDATE
            mdtab.MoveFirst
            mdtab.Seek "=", TRIBE, "STONES QUARRIED"
         End If
      End If
      If Mid(wkres(cnt1), 1, WORDLEN) = "6 STONES/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 6
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "7 STONES/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 7
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "8 STONES/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 8
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "9 STONES/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 9
         mdtab.UPDATE
      End If
         
      If InStr(wkres(cnt1), "logs/person") Then
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "LOGS"
         If mdtab.NoMatch Then
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "LOGS"
            mdtab![AMOUNT] = 4
            mdtab.UPDATE
            mdtab.MoveFirst
            mdtab.Seek "=", TRIBE, "LOGS"
         End If
      End If
      If Mid(wkres(cnt1), 1, WORDLEN) = "5 logs/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 5
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "6 logs/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 6
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "7 logs/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 7
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "8 logs/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 8
         mdtab.UPDATE
      ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "9 logs/person" Then
         mdtab.Edit
         mdtab![AMOUNT] = 9
         mdtab.UPDATE
      End If
         
      If Mid(wkres(cnt1), 1, WORDLEN) = "Medicine 1" Then
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "POPULATION INCREASE"
         If mdtab.NoMatch Then
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "POPULATION INCREASE"
            mdtab![AMOUNT] = 1
            mdtab.UPDATE
         Else
            mdtab.Edit
            mdtab![AMOUNT] = 1
            mdtab.UPDATE
         End If
      End If
         
      If Mid(wkres(cnt1), 1, WORDLEN) = "Extra Movement 4" Then
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "Tribe Movement"
         If mdtab.NoMatch Then
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "Tribe Movement"
            mdtab![AMOUNT] = 4
            mdtab.UPDATE
         Else
            mdtab.Edit
            mdtab![AMOUNT] = mdtab![AMOUNT] + 4
            mdtab.UPDATE
         End If
      End If
            
      If Mid(wkres(cnt1), 1, WORDLEN) = "Extra Movement 6" Then
         mdtab.MoveFirst
         mdtab.Seek "=", TRIBE, "Tribe Movement"
         If mdtab.NoMatch Then
            ' this should not occur but just in case
            mdtab.AddNew
            mdtab![TRIBE] = TRIBE
            mdtab![Modifier] = "Tribe Movement"
            mdtab![AMOUNT] = 6
            mdtab.UPDATE
         Else
            mdtab.Edit
            mdtab![AMOUNT] = mdtab![AMOUNT] + 2
            mdtab.UPDATE
         End If
      End If
            
' Updates to the Skills Table
' update level 11's etc
      If Right(wkres(cnt1), 2) = "11" Then
         SPACE_POS = InStr(wkres(cnt1), " ")
         If SPACE_POS > 0 Then
            Skill = Left(wkres(cnt1), (SPACE_POS - 1))
         End If
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, Skill
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = 11
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
      ElseIf Right(wkres(cnt1), 2) = "12" Then
         SPACE_POS = InStr(wkres(cnt1), " ")
         If SPACE_POS > 0 Then
            Skill = Left(wkres(cnt1), (SPACE_POS - 1))
         End If
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, Skill
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = 12
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
      ElseIf Right(wkres(cnt1), 2) = "13" Then
         SPACE_POS = InStr(wkres(cnt1), " ")
         If SPACE_POS > 0 Then
            Skill = Left(wkres(cnt1), (SPACE_POS - 1))
         End If
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, Skill
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = 13
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
      ElseIf Right(wkres(cnt1), 2) = "14" Then
         SPACE_POS = InStr(wkres(cnt1), " ")
         If SPACE_POS > 0 Then
            Skill = Left(wkres(cnt1), (SPACE_POS - 1))
         End If
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, Skill
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = 14
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
      ElseIf Right(wkres(cnt1), 2) = "15" Then
         SPACE_POS = InStr(wkres(cnt1), " ")
         If SPACE_POS > 0 Then
            Skill = Left(wkres(cnt1), (SPACE_POS - 1))
         End If
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, Skill
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = 15
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
     End If
      
     If Mid(wkres(cnt1), 1, WORDLEN) = "Astral Navigation 1" Then
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, "Navigation"
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = sktab![SKILL LEVEL] + 2
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
     ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "Astral Navigation 2" Then
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, "Navigation"
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = sktab![SKILL LEVEL] + 4
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
     ElseIf Mid(wkres(cnt1), 1, WORDLEN) = "Astral Navigation 3" Then
        sktab.MoveFirst
        sktab.Seek "=", TRIBE, "Navigation"
        If Not sktab.NoMatch Then
           sktab.Edit
           sktab![SKILL LEVEL] = sktab![SKILL LEVEL] + 6
           sktab![SUCCESSFUL] = "Y"
           sktab![ATTEMPTED] = "Y"
           sktab.UPDATE
        End If
     End If
      
     wks0 = "Research Successfully Completed!" & crlf & crlf
     wks1 = "Clan : " & CLAN & crlf
     WKS2 = "Tribe: " & TRIBE & crlf & crlf
     wks3 = "Topic: " & wkres(cnt1) & crlf & crlf
     wks4 = "DL   : " & wkdlcur(cnt1) & " of " & wkdlreq(cnt1) & " Acheived"
     Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4), True)
     wkres(cnt1) = "EMPTY"
     wkdlcur(cnt1) = 0
     wkdlreq(cnt1) = 0
  End If
End If

restopmod = oldrestopmod
resdevmod = oldresdevmod

Next

count = 1

Do While count < 41

   If Not wkres(count) = "EMPTY" Then
      rrtab.MoveFirst
      rrtab.Seek "=", TRIBE, wkres(count)
 
      If rrtab.NoMatch Then
         rrtab.AddNew
         rrtab![TRIBE] = TRIBE
         rrtab![TOPIC] = wkres(count)
         rrtab![DL LEVEL ATTAINED] = wkdlcur(count)
         rrtab![DL LEVEL REQUIRED] = wkdlreq(count)
         rrtab![RESEARCH ATTEMPTED] = "Y"
         rrtab![RESEARCH ATTAINED] = "Y"
         rrtab![Cost] = wkcost(count)
         rrtab.UPDATE
      Else
         rrtab.Edit
         If rrtab![DL LEVEL ATTAINED] = wkdlcur(count) Then
            rrtab![RESEARCH ATTAINED] = "N"
         Else
            rrtab![RESEARCH ATTAINED] = "Y"
         End If
         rrtab![DL LEVEL ATTAINED] = wkdlcur(count)
         rrtab![DL LEVEL REQUIRED] = wkdlreq(count)
         rrtab![RESEARCH ATTEMPTED] = "Y"
         rrtab![Cost] = wkcost(count)
         rrtab.UPDATE
      End If
   End If
   count = count + 1
Loop

roll1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 1, 0)
If codetrack > 0 Then
    wks0 = "After Research Roll(wkres-wkdlcur-wkdlreq-roll1)" & crlf
    wks1 = wkres(1) & "-" & wkdlcur(1) & "-" & wkdlreq(1) & "-" & rdl(1) & crlf
    WKS2 = wkres(2) & "-" & wkdlcur(2) & "-" & wkdlreq(2) & "-" & rdl(2) & crlf
    wks3 = wkres(3) & "-" & wkdlcur(3) & "-" & wkdlreq(3) & "-" & rdl(3) & crlf
    wks4 = wkres(4) & "-" & wkdlcur(4) & "-" & wkdlreq(4) & "-" & rdl(4) & crlf
    wks5 = wkres(5) & "-" & wkdlcur(5) & "-" & wkdlreq(5) & "-" & rdl(5) & crlf
    wks6 = wkres(6) & "-" & wkdlcur(6) & "-" & wkdlreq(6) & "-" & rdl(6)
    Response = MsgBox((wks0 & wks1 & WKS2 & wks3 & wks4 & wks5 & wks6), True)
    Response = MsgBox(("Final Roll:" & roll1), True)
End If

End_Loop:
If SCREEN = "YES" Then
   Exit Do
End If

If research_only = True Then
   ' dont movenext
Else
   SkillsTab.MoveNext
   If SkillsTab![TRIBE] = 9999 Then
      SkillsTab.MoveNext
   End If
End If

If SkillsTab.EOF Then
   If researchtab.EOF Then
      If research_only = True Then
         Exit Do
      Else
         research_only = True
      End If
   ElseIf research_only = False Then
      research_only = True
      researchtab.MoveFirst
      SkillsTab.MoveFirst
   Else
      'research_only = True
      'working through research
     SkillsTab.MoveFirst
     researchtab.MoveNext
   End If
ElseIf research_only = True Then
   If researchtab.EOF Then
      Exit Do
   Else
      researchtab.MoveNext
   End If
Else
   ' do nothing as working through skills
End If

Loop

rrtab.Close
crtab.Close
mdtab.Close
SkillsTab.Close
researchtab.Close
Research_Attempts.Close
newtab.Close

EXIT_FORMS ("SKILLS_1")

OPEN_FORMS ("SKILLS_1")

Forms![SKILLS_1].Refresh

ERR_SKILLS_CLOSE:
 Exit Function

ERR_SKILLS:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

ElseIf (Err = 3420) Then
   Stop
   Resume Next

ElseIf (Err = 6) Then
   Stop
   Resume Next

Else
   MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
          "Error Number: " & Err.NUMBER & vbCrLf & _
          "Error Tribe: " & TRIBE & vbCrLf & _
          "Error Skill: " & Skill_Being_Attempted & vbCrLf & _
          "Error Research: " & Research_Being_Attempted & vbCrLf & _
          "Error Section: " & Section & vbCrLf & _
          "Error Description: " & Err.Description & _
          Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
          , vbOKOnly + vbCritical, "An Error has Occured!"
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
'   Stop
'   Resume Next
   Resume ERR_SKILLS_CLOSE
End If


End Function



