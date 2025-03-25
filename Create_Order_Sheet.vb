Attribute VB_Name = "Create_Order_Sheet"
 Option Compare Database

Function create_workbook(CurrentClan As String, Spreadsheet As String)
'Function create_workbook()

Dim TVWKSPACE As Workspace
Dim TVDB As DAO.Database
Dim RSTemp As Recordset
Dim qdf As QueryDef
Dim objexcel As Excel.Application
Dim wbexcel As Excel.Workbook
Dim wsexcel As Excel.Worksheet
Dim ValidGoodsRange As Range
Dim DocDir, ValidItems As String
Dim i, j, startown, lastrow, ActivityCount, ActivityWDupsCount, UnitCount, NonGCount, MissionCount, ValidDirCount As Integer

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GM = GMTABLE![Name]
FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set Globaltable = TVDBGM.OpenRecordset("Global")
Globaltable.index = "PRIMARYKEY"
Globaltable.MoveFirst

TVDirect = Mid(Globaltable![CURRENT TURN], 1, 2) & Right(Globaltable![CURRENT TURN], 3)

Set objexcel = New Excel.Application
objexcel.DisplayAlerts = False
objexcel.Visible = False
objexcel.ScreenUpdating = False
'objexcel.Workbooks.Add

DocDir = CurDir$ & "\Documents\"
DIRECTPATH = DocDir & TVDirect & "\"

If Dir(DIRECTPATH, vbDirectory) = "" Then
    MkDir (DIRECTPATH)
End If

If Spreadsheet = "Excel" Then
   Set wbexcel = objexcel.Workbooks.Open(DocDir & "Base_Orders.xlsx")
Else
   Set wbexcel = objexcel.Workbooks.Open(DocDir & "Base_Orders.xls")
End If
'Set ValidGoodsRange = wbexcel.Names.ITEM("ValidGoods")

With wbexcel

    With .Sheets("Clan")
        .Range("A1:W23").ClearContents
        .Range("A1:R23").ClearFormats
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "Unit"
        .Range("B1").Value = "GT"
        .Range("C1").Value = "Warrior"
        .Range("D1").Value = "Active"
        .Range("E1").Value = "Inactive"
        .Range("F1").Value = "Slave"
        .Range("G1").Value = "Eaters"
        .Range("H1").Value = "Provs"
        .Range("I1").Value = "Months"
        .Range("J1").Value = "Hirelings"
        .Range("K1").Value = "Mercs"
        .Range("L1").Value = "Locals"
        .Range("M1").Value = "Auxiliaries"
        .Range("N1").Value = "Workers"
        .Range("O1").Value = "Used"
        .Range("P1").Value = "Remains"
        .Range("Q1").Value = "Cattle"
        .Range("R1").Value = "Dog"
        .Range("S1").Value = "Elephant"
        .Range("T1").Value = "Goat"
        .Range("U1").Value = "Horse"
        .Range("V1").Value = "Camel"
        .Range("W1").Value = "Herders"
        
        'get the query "Orders_Wrksht_CLAN"
        Set qdf = CurrentDb.QueryDefs("Orders_Wrksht_CLAN")

        'Supply the parameter value
        qdf.Parameters("ClanNo") = "0" & CurrentClan

        Set RSTemp = qdf.OpenRecordset()
        
'== following commented out as clan sheet is now handled by saved query with a parameter===
'        Set RSTemp = TVDB.OpenRecordset("Select " _
'        & "tc.tribe, iif(IsNull([goods tribe]),tc.tribe,ti.[goods tribe]), tc.warriors, tc.actives, tc.inactives, tc.slave, " _
'        & "'', tg1.item_number, '', ti.hirelings, ti.mercenaries, ti.locals, ti.auxiliaries," _
'        & "'','','',tg2.item_number, tg3.item_number, tg4.item_number, tg5.item_number, tg6.item_number, tg7.item_number " _
'        & "from (((((((TRIBE_CHECKING as TC " _
'        & "left outer join (select * from TRIBES_GENERAL_INFO where clan='0" & CurrentClan & "') as TI on TC.tribe=TI.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='PROVS') as TG1 on TC.tribe=TG1.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='CATTLE') as TG2 on TC.tribe=TG2.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='DOG') as TG3 on TC.tribe=TG3.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='ELEPHANT') as TG4 on TC.tribe=TG4.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='GOAT') as TG5 on TC.tribe=TG5.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='HORSE') as TG6 on TC.tribe=TG6.tribe) " _
'        & "left outer join (select * from TRIBES_GOODS where clan='0" & CurrentClan & "' and item='CAMEL') as TG7 on TC.tribe=TG7.tribe " _
'        & "where TC.clan='0" & CurrentClan & "' ORDER BY tc.tribe")
'===============================================================================
        .Range("A2").CopyFromRecordset RSTemp
        lastrow = RSTemp.RecordCount + 1
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("G2").Formula = "=SUMIF($B$2:$B$" & lastrow & ",$A2,C$2:C$" & lastrow & ")+" _
        & "SUMIF($B$2:$B$" & lastrow & ",$A2,D$2:D$" & lastrow & ")+" _
        & "SUMIF($B$2:$B$" & lastrow & ",$A2,E$2:E$" & lastrow & ")+" _
        & "SUMIF($B$2:$B$" & lastrow & ",$A2,F$2:F$" & lastrow & ")"
        .Range("G2").Copy
        .Range("G3:G" & lastrow).PasteSpecial xlPasteAll
        .Range("I2").Formula = "=IF(OR(G2=0,TRIM(G2)=""""),"""",ROUND(H2/G2,1))"
        .Range("I2").Copy
        .Range("I3:I" & lastrow).PasteSpecial xlPasteAll
        .Range("N2").Formula = "=C2+D2+F2+J2+K2+L2+M2"
        .Range("N2").Copy
        .Range("N3:N" & lastrow).PasteSpecial xlPasteAll
        .Range("O2").Formula = "=SUMIF(Tribes_Activities!$A$2:$A$1009,A2,Tribes_Activities!$E$2:$E$1009)"
        .Range("O2").Copy
        .Range("O3:O" & lastrow).PasteSpecial xlPasteAll
        .Range("P2").Formula = "=N2-O2"
        .Range("P2").Copy
        .Range("P3:P" & lastrow).PasteSpecial xlPasteAll
        .Range("W2").Formula = "=ROUNDUP(Q2/10,0)+ROUNDUP(R2/10,0)+ROUNDUP(S2/5,0)+ROUNDUP((T2)/20,0)+ROUNDUP(U2/10,0)+ROUNDUP(V2/10,0)"
        .Range("W2").Copy
        .Range("W3:W" & lastrow).PasteSpecial xlPasteAll
        .Range("A1:W1").ColumnWidth = 8
        .Range("C1:W1").HorizontalAlignment = xlRight
        .Range("C2:P" & lastrow).Interior.ColorIndex = 34
        .Range("Q2:W" & lastrow).Interior.ColorIndex = 38
        .Range("A2:W" & lastrow).Borders.LineStyle = xlContinuous
        .Select
        .Range("A2").Activate
    End With
    With .Sheets("Tribe_Movement")
        Set RSTemp = TVDB.OpenRecordset("Select tribe from TRIBE_CHECKING where clan='0" & CurrentClan & "'")
        UnitCount = RSTemp.RecordCount
        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select tribe from TRIBE_CHECKING where clan='0" & CurrentClan & "' and ucase(mid(tribe,5,1)) <> 'G' ORDER BY tribe")
        .Range("A2").CopyFromRecordset RSTemp
        NonGCount = RSTemp.RecordCount
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("D2:D" & NonGCount + 1).Value = "Still"
        .Range("E2:AQ" & NonGCount + 1).Value = "Empty"
        .Range("AR2:AR" & NonGCount + 1).Value = "N"
    End With
    With .Sheets("Scout_Movement")
        For i = 0 To 1
            .Range("A" & CStr(i * 8 + 2) & ":A" & CStr(i * 8 + 10)).Value = "'" & CStr(i) & CurrentClan
        Next i
        .Range("G2:O17").Value = "Empty"
        .Range("p2:p17").Value = "N"
    End With
    
    With .Sheets("Tribes_Activities")
    End With
    
    With .Sheets("Skill_Attempts")
    End With
    
    With .Sheets("Research_Attempts")
    End With
   
    With .Sheets("Clan_Goods")
        .Range("A1:F1000").ClearContents
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "Tribe"
        .Range("B1").Value = "Item_Type"
        .Range("C1").Value = "Item"
        .Range("D1").Value = "Number"
        Set RSTemp = Nothing
        Set RSTemp = TVDBGM.OpenRecordset("Select Tribe, Item_Type, Item, Item_Number from TRIBES_GOODS where clan='0" & CurrentClan & "'")
        .Range("A2").CopyFromRecordset RSTemp
        lastrow = RSTemp.RecordCount + 1
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:C1000").Columns.AutoFit
    End With
    
'    '=============Previous Turn's activities=================
'    'get the query "Orders_Wrksht_CLAN"
'    Set qdf = CurrentDb.QueryDefs("ClanPrevActq")
'
'    'Supply the parameter value
'    qdf.Parameters("ClanNo") = "0" & CurrentClan
'    Debug.Print " CC " & CurrentClan
'
'    Set RSTemp = qdf.OpenRecordset()
'
'    With .Sheets("ClanPrevActivities")
'        .Range("A1").Value = "CLAN"
'        .Range("B1").Value = "UNIT"
'        .Range("C1").Value = "SECTION"
'        .Range("D1").Value = "LINE_NUMBER"
'        .Range("E1").Value = "LINE_DETAIL"
'        .Range("A2").CopyFromRecordset RSTemp
'        .Range("A:E").Columns.AutoFit
'    End With
'    RSTemp.Close
'    Set RSTemp = Nothing
'    '=============/Previous Turn's activities=================
    
'    With .Sheets("Clan_Research")
'        .Range("A1:F1000").ClearContents
'        .Range("A1").EntireRow.Font.Bold = True
'        .Range("A1").EntireRow.Interior.ColorIndex = 15
'        .Range("A1").Value = "Tribe"
'        .Range("B1").Value = "Topic"
'        Set RSTemp = Nothing
'        Set RSTemp = TVDBGM.OpenRecordset("Select Tribe, Topic from TRIBES_RESEARCH where tribe='0" & CurrentClan & "'")
'        .Range("A2").CopyFromRecordset RSTemp
'        lastrow = RSTemp.RecordCount + 1
'        RSTemp.Close
'        Set RSTemp = Nothing
'        .Range("A1:C1000").Columns.AutoFit
'    End With
    
    With .Sheets("Valid Activity")
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "ACTIVITY"
        .Range("B1").Value = "ITEM"
        .Range("C1").Value = "TYPE"
        .Range("D1").Value = "SHORTNAME"
        .Range("G1").Value = "Valid_Activity"
        .Range("I1").Value = "Activity"
        .Range("J1").Value = "Item"
        Set RSTemp = TVDB.OpenRecordset("Select activity, item, type, shortname from ACTIVITIES order by activity, item")
        .Range("A2").CopyFromRecordset RSTemp
        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select distinct activity from ACTIVITIES")
        .Range("G2").CopyFromRecordset RSTemp
        ActivityCount = RSTemp.RecordCount
        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select distinct activity, item from ACTIVITIES order by activity, item")
        .Range("I2").CopyFromRecordset RSTemp
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:J1000").Columns.AutoFit
    End With
    With .Sheets("Valid_Implements")
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "ACTIVITY"
        .Range("B1").Value = "ITEM"
        .Range("C1").Value = "IMPLEMENT"
        .Range("F1").Value = "ACTIVITY"
        .Range("I1").Value = "ACTIVITY"
        .Range("J1").Value = "ITEM"
        Set RSTemp = TVDB.OpenRecordset("Select activity, item, implement from IMPLEMENTS order by activity, item")
        .Range("A2").CopyFromRecordset RSTemp
        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select distinct activity from IMPLEMENTS")
        .Range("F2").CopyFromRecordset RSTemp
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("F1:F1000").Columns.AutoFit
        Set RSTemp = TVDB.OpenRecordset("Select distinct activity, item from IMPLEMENTS order by activity, item")
        .Range("I2").CopyFromRecordset RSTemp
        ActivityWDupsCount = RSTemp.RecordCount
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:J1000").Columns.AutoFit
   End With
   With .Sheets("Valid Goods")
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "Goods"
        .Range("B1").Value = "Table"
        .Range("C1").Value = "Shortname"
        Set RSTemp = TVDB.OpenRecordset("Select goods, table, shortname from VALID_GOODS order by goods")
        .Range("A2").CopyFromRecordset RSTemp
        lastrow = RSTemp.RecordCount + 1
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:C1000").Columns.AutoFit
    End With
'    ValidGoodsRange.RefersTo = "='Valid Goods'!$A$2:$A$" & lastrow
    With .Sheets("Valid Units")
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "Unit"
        .Range("B1").Value = "Description"
        .Range("F1").Value = "Direction"
        .Range("G1").Value = "Screen_Related"
        .Range("H1").Value = "Description"
        .Range("J1").Value = "Mission"
        Set RSTemp = TVDB.OpenRecordset("Select tribe, `tribe name` from TRIBES_GENERAL_INFO where clan = '0263' order by tribe")
        .Range("A2").CopyFromRecordset RSTemp
        startown = RSTemp.RecordCount + 2
        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select tribe, `tribe name` from TRIBES_GENERAL_INFO where clan = '0" & CurrentClan & "' order by tribe")
        .Range("A" & startown).CopyFromRecordset RSTemp
        startown = startown + RSTemp.RecordCount + 2
        RSTemp.Close
'        Set RSTemp = Nothing
'        Set RSTemp = TVDB.OpenRecordset("Select tribe, `tribe name` from TRIBES_GENERAL_INFO order by tribe")
'        .Range("A" & startown).CopyFromRecordset RSTemp
'        startown = startown + RSTemp.RecordCount + 2
'        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select direction, screen_related, description from VALID_DIRECTIONS order by direction")
        .Range("F2").CopyFromRecordset RSTemp
        ValidDirCount = RSTemp.RecordCount
        RSTemp.Close
        Set RSTemp = Nothing
        Set RSTemp = TVDB.OpenRecordset("Select mission from VALID_SCOUTING_MISSIONS order by mission")
        .Range("J2").CopyFromRecordset RSTemp
        MissionCount = RSTemp.RecordCount
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:J1000").Columns.AutoFit
    End With
    With .Sheets("Valid_Skills")
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        For i = 0 To 9
            .Range("D1").Offset(0, i).Value = "'" & CStr(i) & CurrentClan
        Next i
        Set RSTemp = TVDB.OpenRecordset("Select " _
        & "vs.skill, vs.group, vs.shortname, t0.lvl, t1.lvl, t2.lvl, t3.lvl, t4.lvl, t5.lvl, t6.lvl, t7.lvl, t8.lvl, t9.lvl " _
        & "from (((((((((VALID_SKILLS as vs " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='0" & CurrentClan & "') as t0 on vs.skill=t0.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='1" & CurrentClan & "') as t1 on vs.skill=t1.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='2" & CurrentClan & "') as t2 on vs.skill=t2.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='3" & CurrentClan & "') as t3 on vs.skill=t3.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='4" & CurrentClan & "') as t4 on vs.skill=t4.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='5" & CurrentClan & "') as t5 on vs.skill=t5.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='6" & CurrentClan & "') as t6 on vs.skill=t6.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='7" & CurrentClan & "') as t7 on vs.skill=t7.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='8" & CurrentClan & "') as t8 on vs.skill=t8.skill) " _
        & "left outer join (select skill, `skill level` as lvl from SKILLS where tribe='9" & CurrentClan & "') as t9 on vs.skill=t9.skill")
        .Range("A2").CopyFromRecordset RSTemp
        lastrow = RSTemp.RecordCount + 1
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:C1000").Columns.AutoFit
        .Range("D1:M1").HorizontalAlignment = xlRight
        .Range("A2:M" & lastrow).Borders.LineStyle = xlContinuous
    End With
    With .Sheets("Valid_Research")
        .Range("A1").EntireRow.Font.Bold = True
        .Range("A1").EntireRow.Interior.ColorIndex = 15
        .Range("A1").Value = "Topic"
        Set RSTemp = TVDB.OpenRecordset("Select distinct topic from RESEARCH order by topic")
        .Range("A2").CopyFromRecordset RSTemp
        RSTemp.Close
        Set RSTemp = Nothing
        .Range("A1:A1000").Columns.AutoFit
    End With
    With .Sheets("Skill_Attempts")
        .Range("E2").Formula = "=IF(ISNA(VLOOKUP($C2,Valid_Skills!$A$2:$M$200,2,FALSE)),"""",VLOOKUP($C2,Valid_Skills!$A$2:$M$200,2,FALSE)&"" - ""&VLOOKUP($C2,Valid_Skills!$A$2:$M$200,3+MATCH($A2,Valid_Skills!$D$1:$M$1),FALSE)+1)"
        .Range("E2").Copy
        .Range("E3:E31").PasteSpecial xlPasteAll
        .Select
        .Range("A2").Activate
    End With

End With

               
wbexcel.Sheets("Clan").Select
If Spreadsheet = "Excel" Then
   wbexcel.SaveAs fileName:=DIRECTPATH & "0" & CurrentClan & "_Orders.xlsx", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
Else
   wbexcel.SaveAs fileName:=DIRECTPATH & "0" & CurrentClan & "_Orders.xls", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
End If

wbexcel.Close
objexcel.Quit

'TVDB.Close

End Function
