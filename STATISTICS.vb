Attribute VB_Name = "STATISTICS"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'*===============================================================================*'
 
' MODULE NAME IS STATISTICS

Global CLANS(3) As String
Global count As Long
Global COUNT1 As Long
Global NumChars As Long
Global CountTribes As Long
Global TOTALCLANS As Long
Global TotalTribes As Long
Global TURN As String
Global x As String
Global Continue As String

Function Number_of_Clans()
Dim TRIBESTABLE As Recordset
Dim LastClan As String
Dim CountClans As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESTABLE = TVDBGM.OpenRecordset("tribes_general_info")
TRIBESTABLE.index = "PRIMARYKEY"
TRIBESTABLE.MoveFirst

CLANS(1) = ""
CLANS(2) = ""
CLANS(3) = ""
count = 1
NumChars = 0
CountClans = 1
LastClan = TRIBESTABLE![CLAN]
TRIBESTABLE.MoveNext

Do Until TRIBESTABLE.EOF
   If TRIBESTABLE![CLAN] = LastClan Then
      TRIBESTABLE.MoveNext
   Else
      If TRIBESTABLE![CLAN] = "0000" Or TRIBESTABLE![CLAN] = "000" Or TRIBESTABLE![CLAN] >= "A" Then
         TRIBESTABLE.MoveNext
      Else
         CountClans = CountClans + 1
         If NumChars = 0 Then
            CLANS(count) = TRIBESTABLE![CLAN]
            NumChars = NumChars + Len(TRIBESTABLE![CLAN])
         Else
            CLANS(count) = CLANS(count) & ", " & TRIBESTABLE![CLAN]
            NumChars = NumChars + 2 + Len(TRIBESTABLE![CLAN])
         End If
         If NumChars > 150 Then
            NumChars = 0
            count = count + 1
         End If
         LastClan = TRIBESTABLE![CLAN]
         TRIBESTABLE.MoveNext
      End If
   End If
   
Loop

Number_of_Clans = CountClans

End Function

Function Number_of_Tribes()
Dim TRIBESTABLE As Recordset
Dim LastClan As String
Dim CountClans As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESTABLE = TVDBGM.OpenRecordset("tribes_general_info")
TRIBESTABLE.index = "PRIMARYKEY"
TRIBESTABLE.MoveFirst

CountTribes = 1
TRIBESTABLE.MoveNext

Do Until TRIBESTABLE.EOF
   If TRIBESTABLE![TRIBE] > "0000" And TRIBESTABLE![TRIBE] < "A" Then
      CountTribes = CountTribes + 1
   End If
   TRIBESTABLE.MoveNext
   
Loop

Number_of_Tribes = CountTribes


End Function

Function PRINT_GOODS_STATS()
On Error Resume Next

Dim GStatsTable As Recordset
Dim Globaltable As Recordset
Dim TribeTable As Recordset
Dim ITEM1 As String
Dim MOST1 As Long
Dim TOTAL1 As Long
Dim AVGTOTAL1 As Long
Dim ITEM2 As String
Dim MOST2 As Long
Dim TOTAL2 As Long
Dim AVGTOTAL2 As Long
Dim ROW As Long
Dim COLUMN As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GM = GMTABLE![Name]

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

Set GStatsTable = TVDBGM.OpenRecordset("GOODS_stats")
GStatsTable.index = "PRIMARYKEY"
GStatsTable.MoveFirst

Set Globaltable = TVDBGM.OpenRecordset("global")
Globaltable.index = "PRIMARYKEY"

TOTALCLANS = Number_of_Clans()
TotalTribes = Number_of_Tribes()

Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = False
'wrdApp.Visible = True
      
DIRECTPATH = CurDir$ & "\documents\Goods Stats\"

CURRENT_DIRECTORY = Dir(DIRECTPATH, vbDirectory)
If IsNull(CURRENT_DIRECTORY) Or CURRENT_DIRECTORY = "" Then
   MkDir (DIRECTPATH)
End If

fileName = DIRECTPATH & "Goods Stats.doc"
Kill fileName

Set wrdDoc = wrdApp.Documents.Add

wrdApp.ActiveDocument.SaveAs DIRECTPATH & "Goods Stats.doc"
wrdApp.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.BottomMargin = CentimetersToPoints(1.5)
wrdApp.ActiveDocument.PageSetup.PaperSize = wdPaperA4
wrdApp.ActiveDocument.Range.ParagraphFormat.SpaceAfter = 0

With wrdApp.Selection
       
   wrdApp.Selection.Font.Name = "Times New Roman"
   wrdApp.Selection.Font.Size = 10
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(1), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(2.5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(3.5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(4.5), Alignment:=wdAlignTabLeft
       
End With

wrdDoc.Save
wrdDoc.Activate

wrdApp.Selection.TypeText "GOODS STATISTICS" & vbCr
wrdApp.Selection.TypeText "Turn : " & Globaltable![CURRENT TURN]
wrdApp.Selection.TypeText vbTab & "Total number of Clans : " & TOTALCLANS
wrdApp.Selection.TypeText vbTab & "Total number of Tribes : " & TotalTribes
wrdApp.Selection.TypeText vbNewLine & vbNewLine
wrdApp.Selection.TypeText "A list of Current Clans in the Game : "
wrdApp.Selection.TypeText CLANS(1) & vbNewLine & CLANS(2) & vbNewLine & CLANS(3)
wrdApp.Selection.TypeText vbNewLine & vbNewLine

' CREATE THE TABLE
Dim myRange As Range
Set myRange = wrdApp.ActiveDocument.Content
myRange.Collapse Direction:=wdCollapseEnd

wrdApp.ActiveDocument.Tables.Add Range:=myRange, numrows:=70, numcolumns:=9


' FORMAT THE TABLE
wrdApp.Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=CentimetersToPoints(3), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=CentimetersToPoints(1), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(6).SetWidth ColumnWidth:=CentimetersToPoints(3), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(7).SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(8).SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(9).SetWidth ColumnWidth:=CentimetersToPoints(2), RulerStyle:=wdAdjustNone

' GO TO THE FIRST LINE OF THE TABLE
ROW = 1
COLUMN = 1
wrdApp.Selection.Tables(1).Cell(ROW, 1).Select

' LINE 01 OF THE TABLE
wrdApp.Selection.TypeText "Item"
wrdApp.Selection.Tables(1).Cell(ROW, 2).Select
wrdApp.Selection.TypeText "Total"
wrdApp.Selection.Tables(1).Cell(ROW, 3).Select
wrdApp.Selection.TypeText "Most"
wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
wrdApp.Selection.TypeText "Avg Total"
wrdApp.Selection.Tables(1).Cell(ROW, 6).Select
wrdApp.Selection.TypeText "Item"
wrdApp.Selection.Tables(1).Cell(ROW, 7).Select
wrdApp.Selection.TypeText "Total"
wrdApp.Selection.Tables(1).Cell(ROW, 8).Select
wrdApp.Selection.TypeText "Most"
wrdApp.Selection.Tables(1).Cell(ROW, 9).Select
wrdApp.Selection.TypeText "Avg Total"

' LINE 02 OF THE TABLE
ROW = 2
wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
wrdApp.Selection.TypeText "per Clan"
wrdApp.Selection.Tables(1).Cell(ROW, 9).Select
wrdApp.Selection.TypeText "per Clan"

' LINE 04 OF THE TABLE
ROW = 3
COLUMN = 1
GStatsTable.MoveFirst
Do Until GStatsTable.EOF
   ROW = ROW + 1
   ITEM1 = GStatsTable![GOODS]
   MOST1 = GStatsTable![NUMBER]
   TOTAL1 = GStatsTable![max]
   AVGTOTAL1 = MOST1 / TOTALCLANS
   wrdApp.Selection.Tables(1).Cell(ROW, 1).Select
   wrdApp.Selection.TypeText ITEM1
   wrdApp.Selection.Tables(1).Cell(ROW, 2).Select
   wrdApp.Selection.TypeText MOST1
   wrdApp.Selection.Tables(1).Cell(ROW, 3).Select
   wrdApp.Selection.TypeText TOTAL1
   wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
   wrdApp.Selection.TypeText AVGTOTAL1
  
   GStatsTable.MoveNext
   If GStatsTable.EOF Then
      Exit Do
   End If
   ITEM2 = GStatsTable![GOODS]
   MOST2 = GStatsTable![NUMBER]
   TOTAL2 = GStatsTable![max]
   AVGTOTAL2 = MOST2 / TOTALCLANS
   wrdApp.Selection.Tables(1).Cell(ROW, 6).Select
   wrdApp.Selection.TypeText ITEM2
   wrdApp.Selection.Tables(1).Cell(ROW, 7).Select
   wrdApp.Selection.TypeText MOST2
   wrdApp.Selection.Tables(1).Cell(ROW, 8).Select
   wrdApp.Selection.TypeText TOTAL2
   wrdApp.Selection.Tables(1).Cell(ROW, 9).Select
   wrdApp.Selection.TypeText AVGTOTAL2
   
   GStatsTable.MoveNext
   If GStatsTable.EOF Then
      Exit Do
   End If
Loop

wrdApp.Documents.Save
wrdApp.Documents.Close

Exit Function

End Function

Function PRINT_SKILL_STATS()
On Error Resume Next

Dim SStatsTable As Recordset
Dim Globaltable As Recordset
Dim TribeTable As Recordset
Dim SKILL1 As String
Dim MOST1 As Long
Dim TOTAL1 As Long
Dim TOTALTRIBES1 As Long
Dim AVGTOTAL1 As Long
Dim SKILL2 As String
Dim MOST2 As Long
Dim TOTAL2 As Long
Dim AVGTOTAL2 As Long
Dim TOTALTRIBES2 As Long
Dim GMNAME As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GM = GMTABLE![Name]
GMNAME = GMTABLE![Name]

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

Set SStatsTable = TVDBGM.OpenRecordset("skills_stats")
SStatsTable.index = "PRIMARYKEY"

Set Globaltable = TVDBGM.OpenRecordset("global")
Globaltable.index = "PRIMARYKEY"

TOTALCLANS = Number_of_Clans()
TotalTribes = Number_of_Tribes()

Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = False
'wrdApp.Visible = True
      
DIRECTPATH = CurDir$ & "\documents\Goods Stats\"

CURRENT_DIRECTORY = Dir(DIRECTPATH, vbDirectory)
If IsNull(CURRENT_DIRECTORY) Or CURRENT_DIRECTORY = "" Then
   MkDir (DIRECTPATH)
End If

fileName = DIRECTPATH & "Skills Stats.doc"
Kill fileName

Set wrdDoc = wrdApp.Documents.Add

wrdApp.ActiveDocument.SaveAs DIRECTPATH & "Skills Stats.doc"
wrdApp.ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.RightMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.BottomMargin = CentimetersToPoints(1)
wrdApp.ActiveDocument.PageSetup.PaperSize = wdPaperA4
wrdApp.ActiveDocument.Range.ParagraphFormat.SpaceAfter = 0

With wrdApp.Selection
       
   wrdApp.Selection.Font.Name = "Times New Roman"
   wrdApp.Selection.Font.Size = 10
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(1), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(2.5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(3.5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=CentimetersToPoints(4.5), Alignment:=wdAlignTabLeft
       
End With

wrdDoc.Save
wrdDoc.Activate

wrdApp.Selection.TypeText "SKILLS STATISTICS{ENTER}"
wrdApp.Selection.TypeText "Turn : " & Globaltable![CURRENT TURN]
wrdApp.Selection.TypeText vbTab & "Total number of Clans : " & TOTALCLANS
wrdApp.Selection.TypeText vbTab & "Total number of Tribes : " & TotalTribes
wrdApp.Selection.TypeText vbNewLine & vbNewLine

' CREATE THE TABLE
Dim myRange As Range
Set myRange = wrdApp.ActiveDocument.Content
myRange.Collapse Direction:=wdCollapseEnd

wrdApp.ActiveDocument.Tables.Add Range:=myRange, numrows:=60, numcolumns:=9

' FORMAT THE TABLE
wrdApp.Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=CentimetersToPoints(4), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(6).SetWidth ColumnWidth:=CentimetersToPoints(4), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(7).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(8).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone
wrdApp.Selection.Tables(1).Columns(9).SetWidth ColumnWidth:=CentimetersToPoints(1.5), RulerStyle:=wdAdjustNone

' LINE 01 OF THE TABLE
ROW = 1
COLUMN = 1
wrdApp.Selection.Tables(1).Cell(ROW, 1).Select
wrdApp.Selection.TypeText "Skill"
wrdApp.Selection.Tables(1).Cell(ROW, 2).Select
wrdApp.Selection.TypeText "Most"
wrdApp.Selection.Tables(1).Cell(ROW, 3).Select
wrdApp.Selection.TypeText "Total"
wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
wrdApp.Selection.TypeText "Avg Total"
wrdApp.Selection.Tables(1).Cell(ROW, 6).Select
wrdApp.Selection.TypeText "Skill"
wrdApp.Selection.Tables(1).Cell(ROW, 7).Select
wrdApp.Selection.TypeText "Most"
wrdApp.Selection.Tables(1).Cell(ROW, 8).Select
wrdApp.Selection.TypeText "Total"
wrdApp.Selection.Tables(1).Cell(ROW, 9).Select
wrdApp.Selection.TypeText "Avg Total"

' LINE 02 OF THE TABLE
ROW = 2
wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
wrdApp.Selection.TypeText "per Clan"
wrdApp.Selection.Tables(1).Cell(ROW, 9).Select
wrdApp.Selection.TypeText "per Clan"

' LINE 04 OF THE TABLE
ROW = 3
COLUMN = 1

SStatsTable.MoveFirst
Do Until SStatsTable.EOF
   ROW = ROW + 1
   SKILL1 = SStatsTable![Skill]
   MOST1 = SStatsTable![SKILL LEVEL]
   TOTAL1 = SStatsTable![TOTAL LEVELS]
   TOTALTRIBES1 = SStatsTable![NUMBER OF TRIBES]
   AVGTOTAL1 = TOTAL1 / TOTALCLANS
   wrdApp.Selection.Tables(1).Cell(ROW, 1).Select
   wrdApp.Selection.TypeText SKILL1
   wrdApp.Selection.Tables(1).Cell(ROW, 2).Select
   wrdApp.Selection.TypeText MOST1
   wrdApp.Selection.Tables(1).Cell(ROW, 3).Select
   wrdApp.Selection.TypeText TOTAL1
   wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
   wrdApp.Selection.TypeText AVGTOTAL1
   
   SStatsTable.MoveNext
   If SStatsTable.EOF Then
      Exit Do
   End If
   SKILL2 = SStatsTable![Skill]
   MOST2 = SStatsTable![SKILL LEVEL]
   TOTAL2 = SStatsTable![TOTAL LEVELS]
   TOTALTRIBES2 = SStatsTable![NUMBER OF TRIBES]
   AVGTOTAL2 = TOTAL2 / TOTALCLANS
   wrdApp.Selection.Tables(1).Cell(ROW, 6).Select
   wrdApp.Selection.TypeText SKILL2
   wrdApp.Selection.Tables(1).Cell(ROW, 7).Select
   wrdApp.Selection.TypeText MOST2
   wrdApp.Selection.Tables(1).Cell(ROW, 8).Select
   wrdApp.Selection.TypeText TOTAL2
   wrdApp.Selection.Tables(1).Cell(ROW, 9).Select
   wrdApp.Selection.TypeText AVGTOTAL2
   
   SStatsTable.MoveNext
   If SStatsTable.EOF Then
      Exit Do
   End If
Loop

wrdApp.Documents.Save
wrdApp.Documents.Close
wrdApp.Quit

Exit Function

End Function

Function CLAN_STATISTICS(STATS)
Dim TRIBESTABLE As Recordset, SKILLSTABLE As Recordset, SStatsTable As Recordset
Dim CStatsTable As Recordset, GStatsTable As Recordset

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM CLAN_STATS;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM SKILLS_STATS;")
qdfCurrent.Execute

Set qdfCurrent = TVDBGM.CreateQueryDef("", "DELETE * FROM GOODS_STATS;")
qdfCurrent.Execute

DoCmd.OpenQuery "STATS - TRIBES GOODS"

Set SKILLSTABLE = TVDBGM.OpenRecordset("skills")
SKILLSTABLE.index = "PRIMARYKEY"
SKILLSTABLE.MoveFirst

Set SStatsTable = TVDBGM.OpenRecordset("skills_stats")
SStatsTable.index = "PRIMARYKEY"

Set TRIBESTABLE = TVDBGM.OpenRecordset("tribes_GENERAL_INFO")
TRIBESTABLE.index = "PRIMARYKEY"
TRIBESTABLE.MoveFirst

Set CStatsTable = TVDBGM.OpenRecordset("Clan_stats")
CStatsTable.index = "PRIMARYKEY"

Set GStatsTable = TVDBGM.OpenRecordset("goods_stats")
GStatsTable.index = "PRIMARYKEY"

Do While Not SKILLSTABLE.EOF
   If SKILLSTABLE![TRIBE] = "0000" Or SKILLSTABLE![TRIBE] = "000" Or SKILLSTABLE![TRIBE] >= "A" Then
      SKILLSTABLE.MoveNext
   Else
      SStatsTable.Seek "=", SKILLSTABLE![Skill]
   
      If SStatsTable.NoMatch Then
         SStatsTable.AddNew
         SStatsTable![Skill] = SKILLSTABLE![Skill]
         SStatsTable![SKILL LEVEL] = SKILLSTABLE![SKILL LEVEL]
         SStatsTable![NUMBER OF TRIBES] = 0
         SStatsTable.UPDATE
         SStatsTable.Seek "=", SKILLSTABLE![Skill]
      End If
   
      SStatsTable.Edit
   
      If SStatsTable![SKILL LEVEL] = SKILLSTABLE![SKILL LEVEL] Then
         SStatsTable![NUMBER OF TRIBES] = SStatsTable![NUMBER OF TRIBES] + 1
         SStatsTable![TOTAL LEVELS] = SStatsTable![TOTAL LEVELS] + SKILLSTABLE![SKILL LEVEL]
      ElseIf SKILLSTABLE![SKILL LEVEL] > SStatsTable![SKILL LEVEL] Then
         SStatsTable![SKILL LEVEL] = SKILLSTABLE![SKILL LEVEL]
         SStatsTable![TOTAL LEVELS] = SStatsTable![TOTAL LEVELS] + SKILLSTABLE![SKILL LEVEL]
         SStatsTable![NUMBER OF TRIBES] = 1
      Else
         SStatsTable![TOTAL LEVELS] = SStatsTable![TOTAL LEVELS] + SKILLSTABLE![SKILL LEVEL]
      End If

      SStatsTable.UPDATE
      SKILLSTABLE.MoveNext
 
      If SKILLSTABLE.EOF Then
         Exit Do
      End If
      If SKILLSTABLE![TRIBE] = "ZZZZ" Then
         Exit Do
      End If
   End If
Loop

TCLANNUMBER = TRIBESTABLE![CLAN]
Continue = "YES"

Do While Continue = "YES"
      
   If TRIBESTABLE![CLAN] = "0000" Or TRIBESTABLE![CLAN] = "000" Or TRIBESTABLE![CLAN] >= "A" Then
      TRIBESTABLE.MoveNext
   Else
      If TRIBESTABLE![SLAVE] > 0 Then
         CStatsTable.Seek "=", TCLANNUMBER, "SLAVE"
           
         If CStatsTable.NoMatch Then
            CStatsTable.AddNew
             CStatsTable![CLAN] = TCLANNUMBER
            CStatsTable![GOOD] = "SLAVE"
            CStatsTable![NUMBER] = TRIBESTABLE![SLAVE]
            CStatsTable.UPDATE
         Else
            CStatsTable.Edit
            CStatsTable![NUMBER] = CStatsTable![NUMBER] + TRIBESTABLE![SLAVE]
            CStatsTable.UPDATE
         End If
      End If

      CStatsTable.index = "primarykey"
      CStatsTable.Seek "=", TCLANNUMBER, "PEOPLE"
      If CStatsTable.NoMatch Then
         CStatsTable.AddNew
         CStatsTable![CLAN] = TCLANNUMBER
         CStatsTable![GOOD] = "PEOPLE"
         CStatsTable![NUMBER] = TRIBESTABLE![WARRIORS] + TRIBESTABLE![ACTIVES] + TRIBESTABLE![INACTIVES]
         CStatsTable.UPDATE
      Else
         CStatsTable.Edit
         CStatsTable![NUMBER] = CStatsTable![NUMBER] + TRIBESTABLE![WARRIORS]
         CStatsTable![NUMBER] = CStatsTable![NUMBER] + TRIBESTABLE![ACTIVES]
         CStatsTable![NUMBER] = CStatsTable![NUMBER] + TRIBESTABLE![INACTIVES]
         CStatsTable.UPDATE
      End If
    
      TRIBESTABLE.MoveNext
   End If
   
   If TRIBESTABLE.EOF Then
      Exit Do
   End If
   
   TCLANNUMBER = TRIBESTABLE![CLAN]
   
   If Left(TRIBESTABLE![CLAN], 3) = "999" Then
      Continue = "NO"
   End If

Loop

DoCmd.OpenQuery "STATS - GOODS STATS"

If STATS = "YES" Then
   Call PRINT_GOODS_STATS

   Call PRINT_SKILL_STATS
End If

End Function

