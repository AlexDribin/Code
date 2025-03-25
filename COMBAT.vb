Attribute VB_Name = "COMBAT"
Option Compare Database   'Use database order for string comparisons
Option Explicit

Function ARCHERY_PHASE()
Dim ITEM(17) As String
Dim leftflank(17) As Long
Dim center(17) As Long
Dim rightflank(17) As Long
Dim CLANNUMBER As String
Dim TRIBENUMBER As String

Dim count As Long

DoCmd.Hourglass True

Set MYFORM = Forms![ARCHERY PHASE]

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

CLANNUMBER = MYFORM![CLAN NUMBER]
TRIBENUMBER = MYFORM![TRIBE NUMBER]

Set TRIBESINFO = TVDBGM.OpenRecordset("tribes - general info")
TRIBESINFO.index = "primarykey"
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", CLANNUMBER, TRIBENUMBER
TRIBESINFO.Edit





End Function

Function CAVALRY_PHASE()

End Function

Function LOOTING_PHASE()

End Function

Function MELEE_PHASE()

End Function

