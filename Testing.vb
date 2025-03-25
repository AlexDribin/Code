Attribute VB_Name = "Testing"
Declare Function GetProfileString Lib "Kernel" (ByVal SName$, ByVal KName$, ByVal Def$, ByVal ret$, ByVal Size%) As Long
Option Compare Database   'Use database order for string comparisons
Option Explicit

Function CHECK_DIRECT()
CURRENT_DIRECTORY = CurDir$
MsgBox "THE DIRECTORY = " & CURRENT_DIRECTORY
End Function

Function check_dll()
Dim dice_sides As Long
Dim level As Long
Dim roll_type As Long
Dim reset_roll As Long
Dim TRIBE As Long
Dim PRESET As Long
Dim MODIFY As Long

dice_sides = 100
level = 1
roll_type = 6
reset_roll = 0
TRIBE = 30
PRESET = 1
MODIFY = 0
x = DROLL(roll_type, level, dice_sides, reset_roll, TRIBE, PRESET, MODIFY)

MsgBox (x)

End Function

Function DeclareDemo()
Dim SName As String
Dim KName As String
Dim ret As String
Dim Success As String

    SName = "Intl" ' WIN.INI section name.
    KName = "sCountry" ' WIN.INI country code.
    ret = String(255, 0)  ' Initialize return string.
' Call Windows Kernel DLL.
    Success = GetProfileString(SName, KName, "", ret, Len(ret))
    If Success Then ' Evaluate results.
        Msg = "'" & KName & "' = " & ret
    Else
        Msg = "There is no country code in your WIN.INI file."
    End If
    MsgBox Msg          ' Display message.
End Function

Public Function test()
Dim SEQ_NUMBER As Long
Dim TRIBE As String
Dim CONSTRUCTION As String
Dim LOGS As Long
Dim STONES As Long
Dim COAL As Long
Dim BRASS As Long
Dim BRONZE As Long
Dim COPPER As Long
Dim IRON As Long
Dim LEAD As Long
Dim CLOTH As Long
Dim LEATHER As Long
Dim ROPES As Long
Dim LOGS_H As Long
Dim QUERY As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst
Current_Turn = globalinfo![CURRENT TURN]
TURN_NUMBER = "TURN" & Left(globalinfo![CURRENT TURN], 2)
globalinfo.Close

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst

Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.MoveFirst

Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst

Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXMAPCONST.index = "PRIMARYKEY"

Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTAB.index = "PRIMARYKEY"
COMPRESTAB.MoveFirst

Set ConstructionTable = TVDBGM.OpenRecordset("Under_Construction")
ConstructionTable.index = "PRIMARYKEY"
ConstructionTable.MoveFirst

Set ConstructionTable2 = TVDB.OpenRecordset("Under_Construction_TEMP")
ConstructionTable2.index = "PRIMARYKEY"

Do Until TRIBESGOODS.EOF
   If IsNull(TRIBESGOODS![ITEM_NUMBER]) Then
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = 0
      TRIBESGOODS.UPDATE
   End If
   TRIBESGOODS.MoveNext
   If TRIBESGOODS.EOF Then
      Exit Do
   End If
Loop

End Function


Public Function Write_A_Flat_File()
Dim globalinfo As Recordset
Dim TribesTurnsActivity As Recordset

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst
Current_Turn = globalinfo![CURRENT TURN]
TURN_NUMBER = "TURN" & Left(globalinfo![CURRENT TURN], 2)
globalinfo.Close

Set TribesTurnsActivity = TVDBGM.OpenRecordset("Tribes_Turns_Activity")
TribesTurnsActivity.index = "ActivityOrder"
TribesTurnsActivity.MoveFirst

Open "Testfile.txt" For Output Shared As #1 Len = 200
Write #1, "Tribes_Turns_Activity"

Do While Not TribesTurnsActivity.EOF
Write #1, TribesTurnsActivity![CLAN]; Tab(11); TribesTurnsActivity![TRIBE]; Tab(21); TribesTurnsActivity![Order]; Tab(23); TribesTurnsActivity![ACTIVITY]; Tab(73); _
TribesTurnsActivity![ITEM]; Tab(123); TribesTurnsActivity![DISTINCTION]; Tab(143); TribesTurnsActivity![ACTIVES]; Tab(148); TribesTurnsActivity![JOINT]
TribesTurnsActivity.MoveNext
Loop

Write #1, "Tribes_Turns_Activity finished"

Close #1

TribesTurnsActivity.Close

End Function

Public Function Read_A_Flat_File()
Dim globalinfo As Recordset
Dim TribesTurnsActivity As Recordset
Dim Record_Data As String
Dim Record_Format As String
Dim CLAN As String
Dim TRIBE As String
Dim Order As Double
Dim ACTIVITY As String
Dim ITEM As String
Dim DISTINCTION As String
Dim ACTIVES As Double
Dim JOINT As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
DoCmd.Hourglass True

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst
Current_Turn = globalinfo![CURRENT TURN]
TURN_NUMBER = "TURN" & Left(globalinfo![CURRENT TURN], 2)
globalinfo.Close

Set TribesTurnsActivity = TVDBGM.OpenRecordset("Tribes_Turns_Activity")
TribesTurnsActivity.index = "ActivityOrder"
TribesTurnsActivity.MoveFirst

Close #1
Open "Testfile.txt" For Input As #1
Input #1, Record_Format

Do While Not TribesTurnsActivity.EOF
     Input #1, CLAN, TRIBE, Order, ACTIVITY, ITEM, DISTINCTION, ACTIVES, JOINT
     If Mid(Record_Data, 1, 21) = "Tribes_Turns_Activity" Then
         Record_Format = Mid(Record_Data, 1, 21)
     Else
          TribesTurnsActivity.AddNew
          TribesTurnsActivity![CLAN] = CLAN
          TribesTurnsActivity![TRIBE] = TRIBE
          TribesTurnsActivity![Order] = Order
          TribesTurnsActivity![ACTIVITY] = ACTIVITY
          TribesTurnsActivity![ITEM] = ITEM
          TribesTurnsActivity![DISTINCTION] = DISTINCTION
          TribesTurnsActivity![ACTIVES] = ACTIVES
          TribesTurnsActivity![JOINT] = JOINT
          TribesTurnsActivity.UPDATE
        
     End If
Loop

Write #1, "Tribes_Turns_Activity finished"

Close #1

TribesTurnsActivity.Close


End Function

