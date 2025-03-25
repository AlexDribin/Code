Attribute VB_Name = "UPDATE TRIBES"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'*                           VERSION 3.1t                                         *'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 28/01/96  *  Add function for updating Activities Screen                    **'
'** 20/01/25  *  Fixed human capacity bug (Alex D)                              **'
'** 22/01/25  *  Fixed  Animal/wagon capacity calculation (Alex D)              **'
'** 10/09/24  *  Empty passenges spaces on fleet calculation fixed (AlexD)      **'
'** 06/03/25  *  Fixed calculation of mounted capacity (AlexD)                  **'
'*===============================================================================*'
Global ShowUpdateTribesVersion As Long

Function UPDATE_ACTIVITIES()
Dim ActivitySeqTable As Recordset
ReDim Description(9) As String
ReDim NUMBER(9) As Double
Dim Counter As Long

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![Activities]

  Description(1) = Forms![Activities]![DESCRIPTION 1]
  NUMBER(1) = Forms![Activities]![VALUE 1]
  Description(2) = Forms![Activities]![DESCRIPTION 2]
  NUMBER(2) = Forms![Activities]![VALUE 2]
  Description(3) = Forms![Activities]![DESCRIPTION 3]
  NUMBER(3) = Forms![Activities]![VALUE 3]
  Description(4) = Forms![Activities]![DESCRIPTION 4]
  NUMBER(4) = Forms![Activities]![VALUE 4]
  Description(5) = Forms![Activities]![DESCRIPTION 5]
  NUMBER(5) = Forms![Activities]![VALUE 5]
  Description(6) = Forms![Activities]![DESCRIPTION 6]
  NUMBER(6) = Forms![Activities]![VALUE 6]
  Description(7) = Forms![Activities]![DESCRIPTION 7]
  NUMBER(7) = Forms![Activities]![VALUE 7]
  Description(8) = Forms![Activities]![DESCRIPTION 8]
  NUMBER(8) = Forms![Activities]![VALUE 8]
  Description(9) = Forms![Activities]![DESCRIPTION 9]
  NUMBER(9) = Forms![Activities]![VALUE 9]

Set activitytable = TVDB.OpenRecordset("ACTIVITY")
activitytable.index = "PRIMARYKEY"

Set ActivitiesTable = TVDB.OpenRecordset("ACTIVITIES")
ActivitiesTable.index = "primarykey"

Set ActivitySeqTable = TVDB.OpenRecordset("Activity_Sequence")
ActivitySeqTable.index = "primarykey"

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"

ActivitySeqTable.MoveFirst
ActivitySeqTable.Seek "=", MYFORM![ACTIVITY TYPE]

If ActivitySeqTable.NoMatch Then
   ActivitySeqTable.AddNew
   ActivitySeqTable![ACTIVITY] = MYFORM![ACTIVITY TYPE]
   ActivitySeqTable![Sequence] = 99
   ActivitySeqTable.UPDATE
End If

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", MYFORM![ACTIVITY TYPE], MYFORM![ITEM TYPE], MYFORM![TYPE]
If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![ACTIVITY] = MYFORM![ACTIVITY TYPE]
   ActivitiesTable![ITEM] = MYFORM![ITEM TYPE]
   ActivitiesTable![TYPE] = MYFORM![TYPE]
   ActivitiesTable![SHORTNAME] = MYFORM![SHORTNAME]
   ActivitiesTable![SKILL LEVEL] = MYFORM![SKILL LEVEL]
   ActivitiesTable![SECOND SKILL] = MYFORM![SECOND SKILL]
   ActivitiesTable![SECOND SKILL LEVEL] = MYFORM![SECOND SKILL LEVEL]
   ActivitiesTable![THIRD SKILL] = MYFORM![THIRD SKILL]
   ActivitiesTable![THIRD SKILL LEVEL] = MYFORM![THIRD SKILL LEVEL]
   ActivitiesTable![FORTH SKILL] = MYFORM![FORTH SKILL]
   ActivitiesTable![FORTH SKILL LEVEL] = MYFORM![FORTH SKILL LEVEL]
   ActivitiesTable![NUMBER OF ITEMS] = MYFORM![NUMBER OF ITEMS]
   ActivitiesTable![PEOPLE] = MYFORM![PEOPLE]
   ActivitiesTable![research] = MYFORM![research]
   ActivitiesTable![GOODS_USED] = MYFORM![GOODS_USED]
   ActivitiesTable![GOOD_PRODUCED] = MYFORM![GOOD_PRODUCED]
   ActivitiesTable.UPDATE

Else
   ActivitiesTable.Edit
   ActivitiesTable![SHORTNAME] = MYFORM![SHORTNAME]
   ActivitiesTable![SKILL LEVEL] = MYFORM![SKILL LEVEL]
   ActivitiesTable![SECOND SKILL] = MYFORM![SECOND SKILL]
   ActivitiesTable![SECOND SKILL LEVEL] = MYFORM![SECOND SKILL LEVEL]
   ActivitiesTable![THIRD SKILL] = MYFORM![THIRD SKILL]
   ActivitiesTable![THIRD SKILL LEVEL] = MYFORM![THIRD SKILL LEVEL]
   ActivitiesTable![FORTH SKILL] = MYFORM![FORTH SKILL]
   ActivitiesTable![FORTH SKILL LEVEL] = MYFORM![FORTH SKILL LEVEL]
   ActivitiesTable![NUMBER OF ITEMS] = MYFORM![NUMBER OF ITEMS]
   ActivitiesTable![PEOPLE] = MYFORM![PEOPLE]
   ActivitiesTable![research] = MYFORM![research]
   ActivitiesTable![GOODS_USED] = Forms![Activities]![GOODS_USED]
   ActivitiesTable![GOOD_PRODUCED] = Forms![Activities]![GOOD_PRODUCED]
   ActivitiesTable.UPDATE
End If

Counter = 1

Do Until Counter > 9
  If Not IsNull(Description(Counter)) And Not (Description(Counter) = "") Then
     activitytable.MoveFirst
     activitytable.Seek "=", MYFORM![ACTIVITY TYPE], MYFORM![ITEM TYPE], MYFORM![TYPE], Description(Counter)
     If activitytable.NoMatch Then
        VALIDGOODS.MoveFirst
        VALIDGOODS.Seek "=", Description(Counter)
        If VALIDGOODS.NoMatch Then
           Msg = Description(Counter) & " is not a valid good.  Please check valid goods table."
           MsgBox (Msg)
        Else
           activitytable.AddNew
           activitytable![ACTIVITY] = MYFORM![ACTIVITY TYPE]
           activitytable![ITEM] = MYFORM![ITEM TYPE]
           activitytable![TYPE] = MYFORM![TYPE]
           activitytable![GOOD] = Description(Counter)
           activitytable![NUMBER] = NUMBER(Counter)
           activitytable.UPDATE
        End If
     Else
        activitytable.Edit
        activitytable![NUMBER] = NUMBER(Counter)
        activitytable.UPDATE
     End If
  End If
  Counter = Counter + 1

Loop

activitytable.Close
ActivitySeqTable.Close

Call EXIT_FORMS("ACTIVITIES")
Call OPEN_FORMS("ACTIVITIES")

End Function

Function UPDATE_HEX_MAP_FORM()
Dim MAP As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set MYFORM = Forms![HEX_MAP]
MAP = MYFORM![MAP]

Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"
hexmaptable.MoveFirst
hexmaptable.Seek "=", MYFORM![MAP]

If Not hexmaptable.NoMatch Then
   hexmaptable.Edit
   If IsNull(MYFORM![CITY]) Then
      hexmaptable![CITY] = Null
   Else
      hexmaptable![CITY] = "Y"
   End If
   If IsNull(MYFORM![ORE TYPE]) Then
      hexmaptable![ORE TYPE] = Null
   Else
      hexmaptable![ORE TYPE] = "Y"
   End If
   If MYFORM![Feature_One] = "QUARRY" _
   Or MYFORM![Feature_Two] = "QUARRY" _
   Or MYFORM![Feature_Three] = "QUARRY" Then
      hexmaptable![QUARRYING] = "y"
   Else
      hexmaptable![QUARRYING] = "n"
   End If
   If MYFORM![Feature_One] = "SPRING" _
   Or MYFORM![Feature_Two] = "SPRING" _
   Or MYFORM![Feature_Three] = "SPRING" Then
      hexmaptable![SPRINGS] = "y"
   Else
      hexmaptable![SPRINGS] = "n"
   End If
   If MYFORM![Feature_One] = "SALMON" _
   Or MYFORM![Feature_Two] = "SALMON" _
   Or MYFORM![Feature_Three] = "SALMON" Then
      hexmaptable![SALMON RUN] = "y"
   Else
      hexmaptable![SALMON RUN] = "n"
   End If
   If MYFORM![Feature_One] = "FISH" _
   Or MYFORM![Feature_Two] = "FISH" _
   Or MYFORM![Feature_Three] = "FISH" Then
      hexmaptable![FISH AREA] = "y"
   Else
      hexmaptable![FISH AREA] = "n"
   End If
   If MYFORM![Feature_One] = "WHALE" _
   Or MYFORM![Feature_Two] = "WHALE" _
   Or MYFORM![Feature_Three] = "WHALE" Then
      hexmaptable![WHALE AREA] = "y"
   Else
      hexmaptable![WHALE AREA] = "n"
   End If
   
   hexmaptable![Borders] = Left(MYFORM![North_Border], 2) & Left(MYFORM![North_East_Border], 2)
   hexmaptable![Borders] = hexmaptable![Borders] & Left(MYFORM![South_East_Border], 2)
   hexmaptable![Borders] = hexmaptable![Borders] & Left(MYFORM![South_Border], 2)
   hexmaptable![Borders] = hexmaptable![Borders] & Left(MYFORM![South_West_Border], 2)
   hexmaptable![Borders] = hexmaptable![Borders] & Left(MYFORM![North_West_Border], 2)
   hexmaptable![ROADS] = Left(MYFORM![ROAD(N)], 1) & Left(MYFORM![ROAD(NE)], 1)
   hexmaptable![ROADS] = hexmaptable![ROADS] & Left(MYFORM![ROAD(SE)], 1)
   hexmaptable![ROADS] = hexmaptable![ROADS] & Left(MYFORM![ROAD(S)], 1)
   hexmaptable![ROADS] = hexmaptable![ROADS] & Left(MYFORM![ROAD(SW)], 1)
   hexmaptable![ROADS] = hexmaptable![ROADS] & Left(MYFORM![ROAD(NW)], 1)
   hexmaptable.UPDATE

End If

If Not IsNull(MYFORM![CITY]) Then
   Set HEXMAPCITY = TVDBGM.OpenRecordset("HEX_MAP_CITY")
   HEXMAPCITY.index = "PRIMARYKEY"
   HEXMAPCITY.MoveFirst
   HEXMAPCITY.Seek "=", MYFORM![MAP]
   
   If HEXMAPCITY.NoMatch Then
      HEXMAPCITY.AddNew
      HEXMAPCITY![MAP] = MYFORM![MAP]
      HEXMAPCITY![CITY] = MYFORM![CITY]
      HEXMAPCITY.UPDATE
   ElseIf IsNull(MYFORM![CITY]) Then
      HEXMAPCITY.Delete
   Else
      HEXMAPCITY.Edit
      HEXMAPCITY![CITY] = MYFORM![CITY]
      HEXMAPCITY.UPDATE
   End If
   HEXMAPCITY.Close
End If

If Not IsNull(MYFORM![ORE TYPE]) Then
   Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
   HEXMAPMINERALS.index = "PRIMARYKEY"
   HEXMAPMINERALS.MoveFirst
   HEXMAPMINERALS.Seek "=", MYFORM![MAP]

   If HEXMAPMINERALS.NoMatch Then
      If IsNull(MYFORM![ORE TYPE]) And _
         IsNull(MYFORM![SECOND ORE]) And _
         IsNull(MYFORM![THIRD ORE]) Then
         ' IGNORE
      Else
         HEXMAPMINERALS.AddNew
         HEXMAPMINERALS![MAP] = MYFORM![MAP]
         HEXMAPMINERALS![ORE_TYPE] = MYFORM![ORE TYPE]
'         HEXMAPMINERALS![SECOND_ORE] = MYFORM![SECOND ORE]
'         HEXMAPMINERALS![THIRD_ORE] = MYFORM![THIRD ORE]
'         HEXMAPMINERALS![SECOND_MINING] = MYFORM![SECOND MINING]
'         HEXMAPMINERALS![THIRD_MINING] = MYFORM![THIRD MINING]
         HEXMAPMINERALS.UPDATE
      End If
   Else
      HEXMAPMINERALS.Edit
      HEXMAPMINERALS![ORE_TYPE] = MYFORM![ORE TYPE]
'      HEXMAPMINERALS![SECOND_ORE] = MYFORM![SECOND ORE]
'      HEXMAPMINERALS![THIRD_ORE] = MYFORM![THIRD ORE]
'      HEXMAPMINERALS![SECOND_MINING] = MYFORM![SECOND MINING]
'      HEXMAPMINERALS![THIRD_MINING] = MYFORM![THIRD MINING]
      HEXMAPMINERALS.UPDATE
   End If
   HEXMAPMINERALS.Close
End If

hexmaptable.Close
EXIT_FORMS ("HEX_MAP")
OPEN_FORMS ("HEX_MAP")
'DoCmd.FindRecord MAP, acEntire
'Forms![HEX_MAP]![MAP].SetFocus

End Function


Public Function UPDATE_TRIBES_GOODS()
On Error GoTo ERR_UPDATE_TRIBES_GOODS
TRIBE_STATUS = "Update Tribes Goods"

Dim TRIBESGOODS As Recordset
Dim TribesModifiers As Recordset
Dim CLANNUMBER As String
Dim TRIBENUMBER As String
ReDim Description(60) As String
ReDim NUMBER(60) As Long
Dim Counter As Long
Dim ITEM_TYPE As String

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

Description(1) = MYFORM![DESCRIPTION 1]
NUMBER(1) = MYFORM![VALUE 1]
Description(2) = MYFORM![DESCRIPTION 2]
NUMBER(2) = MYFORM![VALUE 2]
Description(3) = MYFORM![DESCRIPTION 3]
NUMBER(3) = MYFORM![VALUE 3]
Description(4) = MYFORM![DESCRIPTION 4]
NUMBER(4) = MYFORM![VALUE 4]
Description(5) = MYFORM![DESCRIPTION 5]
NUMBER(5) = MYFORM![VALUE 5]
Description(6) = MYFORM![DESCRIPTION 6]
NUMBER(6) = MYFORM![VALUE 6]
Description(7) = MYFORM![DESCRIPTION 7]
NUMBER(7) = MYFORM![VALUE 7]
Description(8) = MYFORM![DESCRIPTION 8]
NUMBER(8) = MYFORM![VALUE 8]
Description(9) = MYFORM![DESCRIPTION 9]
NUMBER(9) = MYFORM![VALUE 9]
Description(10) = MYFORM![DESCRIPTION 10]
NUMBER(10) = MYFORM![VALUE 10]
Description(11) = MYFORM![DESCRIPTION 11]
NUMBER(11) = MYFORM![VALUE 11]
Description(12) = MYFORM![DESCRIPTION 12]
NUMBER(12) = MYFORM![VALUE 12]
Description(13) = MYFORM![DESCRIPTION 13]
NUMBER(13) = MYFORM![VALUE 13]
Description(14) = MYFORM![DESCRIPTION 14]
NUMBER(14) = MYFORM![VALUE 14]
Description(15) = MYFORM![DESCRIPTION 15]
NUMBER(15) = MYFORM![VALUE 15]
Description(16) = MYFORM![DESCRIPTION 16]
NUMBER(16) = MYFORM![VALUE 16]
Description(17) = MYFORM![DESCRIPTION 17]
NUMBER(17) = MYFORM![VALUE 17]
Description(18) = MYFORM![DESCRIPTION 18]
NUMBER(18) = MYFORM![VALUE 18]
Description(19) = MYFORM![DESCRIPTION 19]
NUMBER(19) = MYFORM![VALUE 19]
Description(20) = MYFORM![DESCRIPTION 20]
NUMBER(20) = MYFORM![VALUE 20]
Description(21) = MYFORM![DESCRIPTION 21]
NUMBER(21) = MYFORM![VALUE 21]
Description(22) = MYFORM![DESCRIPTION 22]
NUMBER(22) = MYFORM![VALUE 22]
Description(23) = MYFORM![DESCRIPTION 23]
NUMBER(23) = MYFORM![VALUE 23]
Description(24) = MYFORM![DESCRIPTION 24]
NUMBER(24) = MYFORM![VALUE 24]
Description(25) = MYFORM![DESCRIPTION 25]
NUMBER(25) = MYFORM![VALUE 25]
Description(26) = MYFORM![DESCRIPTION 26]
NUMBER(26) = MYFORM![VALUE 26]
Description(27) = MYFORM![DESCRIPTION 27]
NUMBER(27) = MYFORM![VALUE 27]
Description(28) = MYFORM![DESCRIPTION 28]
NUMBER(28) = MYFORM![VALUE 28]
Description(29) = MYFORM![DESCRIPTION 29]
NUMBER(29) = MYFORM![VALUE 29]
Description(30) = MYFORM![DESCRIPTION 30]
NUMBER(30) = MYFORM![VALUE 30]
Description(31) = MYFORM![DESCRIPTION 31]
NUMBER(31) = MYFORM![VALUE 31]
Description(32) = MYFORM![DESCRIPTION 32]
NUMBER(32) = MYFORM![VALUE 32]
Description(33) = MYFORM![DESCRIPTION 33]
NUMBER(33) = MYFORM![VALUE 33]
Description(34) = MYFORM![DESCRIPTION 34]
NUMBER(34) = MYFORM![VALUE 34]
Description(35) = MYFORM![DESCRIPTION 35]
NUMBER(35) = MYFORM![VALUE 35]
Description(36) = MYFORM![DESCRIPTION 36]
NUMBER(36) = MYFORM![VALUE 36]
Description(37) = MYFORM![DESCRIPTION 37]
NUMBER(37) = MYFORM![VALUE 37]
Description(38) = MYFORM![DESCRIPTION 38]
NUMBER(38) = MYFORM![VALUE 38]
Description(39) = MYFORM![DESCRIPTION 39]
NUMBER(39) = MYFORM![VALUE 39]
Description(40) = MYFORM![DESCRIPTION 40]
NUMBER(40) = MYFORM![VALUE 40]
Description(41) = MYFORM![DESCRIPTION 41]
NUMBER(41) = MYFORM![VALUE 41]
Description(42) = MYFORM![DESCRIPTION 42]
NUMBER(42) = MYFORM![VALUE 42]
Description(43) = MYFORM![DESCRIPTION 43]
NUMBER(43) = MYFORM![VALUE 43]
Description(44) = MYFORM![DESCRIPTION 44]
NUMBER(44) = MYFORM![VALUE 44]
Description(45) = MYFORM![DESCRIPTION 45]
NUMBER(45) = MYFORM![VALUE 45]
Description(46) = MYFORM![DESCRIPTION 46]
NUMBER(46) = MYFORM![VALUE 46]
Description(47) = MYFORM![DESCRIPTION 47]
NUMBER(47) = MYFORM![VALUE 47]
Description(48) = MYFORM![DESCRIPTION 48]
NUMBER(48) = MYFORM![VALUE 48]
Description(49) = MYFORM![DESCRIPTION 49]
NUMBER(49) = MYFORM![VALUE 49]
Description(50) = MYFORM![DESCRIPTION 50]
NUMBER(50) = MYFORM![VALUE 50]
Description(51) = MYFORM![DESCRIPTION 51]
NUMBER(51) = MYFORM![VALUE 51]
Description(52) = MYFORM![DESCRIPTION 52]
NUMBER(52) = MYFORM![VALUE 52]
Description(53) = MYFORM![DESCRIPTION 53]
NUMBER(53) = MYFORM![VALUE 53]
Description(54) = MYFORM![DESCRIPTION 54]
NUMBER(54) = MYFORM![VALUE 54]
Description(55) = MYFORM![DESCRIPTION 55]
NUMBER(55) = MYFORM![VALUE 55]
Description(56) = MYFORM![DESCRIPTION 56]
NUMBER(56) = MYFORM![VALUE 56]
Description(57) = MYFORM![DESCRIPTION 57]
NUMBER(57) = MYFORM![VALUE 57]
Description(58) = MYFORM![DESCRIPTION 58]
NUMBER(58) = MYFORM![VALUE 58]
Description(59) = MYFORM![DESCRIPTION 59]
NUMBER(59) = MYFORM![VALUE 59]
Description(60) = MYFORM![DESCRIPTION 60]
NUMBER(60) = MYFORM![VALUE 60]
  
If ITEM_TYPE = "MODIFIERS" Then
   Set TribesModifiers = TVDBGM.OpenRecordset("MODIFIERS")
   TribesModifiers.index = "PRIMARYKEY"

   Counter = 1

   Do Until Counter > 60
     If Not IsNull(Description(Counter)) And Not (Description(Counter) = "") Then
        TribesModifiers.MoveFirst
        TribesModifiers.Seek "=", MYFORM![TRIBENUMBER], Description(Counter)
        If TribesModifiers.NoMatch Then
           TribesModifiers.AddNew
           TribesModifiers![TRIBE] = MYFORM![TRIBENUMBER]
           TribesModifiers![Modifier] = Description(Counter)
           TribesModifiers![AMOUNT] = NUMBER(Counter)
           TribesModifiers.UPDATE
        Else
           TribesModifiers.Edit
           TribesModifiers![AMOUNT] = NUMBER(Counter)
           TribesModifiers.UPDATE
        End If
     End If
     Counter = Counter + 1

   Loop

   TribesModifiers.Close

   Call EXIT_FORMS("TRIBES - MODIFIERS")

Else
   Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_GOODS")
   TRIBESGOODS.index = "PRIMARYKEY"

   Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
   VALIDGOODS.index = "primarykey"
  
   Counter = 1

   Do Until Counter > 60
     If Not IsNull(Description(Counter)) And Not (Description(Counter) = "") Then
        TRIBESGOODS.MoveFirst
        TRIBESGOODS.Seek "=", MYFORM![CLANNUMBER], MYFORM![TRIBENUMBER], MYFORM![ITEM_TYPE], Description(Counter)
        If TRIBESGOODS.NoMatch Then
           VALIDGOODS.MoveFirst
           VALIDGOODS.Seek "=", Description(Counter)
           If VALIDGOODS.NoMatch Then
              Msg = Description(Counter) & " is not a valid good.  Please check valid goods table."
              MsgBox (Msg)
           Else
              TRIBESGOODS.AddNew
              TRIBESGOODS![CLAN] = MYFORM![CLANNUMBER]
              TRIBESGOODS![TRIBE] = MYFORM![TRIBENUMBER]
              TRIBESGOODS![ITEM_TYPE] = MYFORM![ITEM_TYPE]
              TRIBESGOODS![ITEM] = Description(Counter)
              TRIBESGOODS![ITEM_NUMBER] = NUMBER(Counter)
              TRIBESGOODS.UPDATE
           End If
        Else
           TRIBESGOODS.Edit
           TRIBESGOODS![ITEM_NUMBER] = NUMBER(Counter)
           TRIBESGOODS.UPDATE
        End If
     End If
     Counter = Counter + 1

   Loop

   TRIBESGOODS.Close

Call EXIT_FORMS("TRIBES - GOODS")
End If

Call OPEN_FORM_TRIBES_GOODS(ITEM_TYPE)


ERR_UPDATE_TRIBES_GOODS_CLOSE:
   Exit Function


ERR_UPDATE_TRIBES_GOODS:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
   Msg = "Error # " & Err & " " & Error$
   MsgBox (Msg)
   Resume ERR_UPDATE_TRIBES_GOODS_CLOSE
End If


End Function




'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 16/10/17  *  Changing function to calculate Capacity in two passes          **'
'**           *  First pass will do the tribe/element/fleet                     **'
'**           *  The second will do the Goods Tribes                            **'
'**           *  Capacity holds tribe/elements capacity                         **'
'**           *  Weight holds the tribe/elements weight in goods                **'
'**           *  Walking_Capacity is the GT/tribe/elements movement capacity    **'
'**           *  should be zero in a GT scenario                                **'
'**           *  Use Valid_Goods table for weights and capacity                 **'
'**           *  GT_MOUNTED_CAPACITY & GT_WALKING_CAPACITY are new              **'
'*===============================================================================*'
 
Public Function Determine_Capacities(Capacity_Type, DC_CLANNUMBER, DC_TRIBENUMBER)
On Error GoTo Determine_Capacities_Error
TRIBE_STATUS = "Determine Capacities"

Dim CAPACITY As Long         ' will be loaded into Capacity on Tribes_General_Info - treat as mounted_capacity
Dim Walking_Capacity As Long ' will be loaded into Walking_Capacity on Tribes_General_Info
Dim Mounted_Capacity As Long ' will be loaded into Mounted_Capacity on Tribes_General_Info
Dim TotalPeople As Long
Dim TotalPalaquins As Long
Dim TotalBackpacks As Long
Dim TotalCattle As Long
Dim TotalHorses As Long
Dim TotalLHorses As Long
Dim TotalHHorses As Long
Dim TotalCamels As Long
Dim TotalElephants As Long
Dim TotalSaddleBags As Long
Dim TotalWagons As Long
Dim TotalScouts As Long
Dim TotalBackpacksUsed As Long
Dim CamelCapacity As Long
Dim HorseCapacity As Long
Dim ElephantCapacity As Long
Dim Required_Crew As Long
Dim Max_Crew As Long
Dim AVAILABLE_PEOPLE As Long
Dim EXCESS_PEOPLE As Long
Dim Wagons_Carry As Long
Dim WagonsNumber As Long
Dim MSG1 As String
Dim MSG2 As String



If (ShowUpdateTribesVersion = 0) Then
   MSG1 = "UPDATE TRIBES module version is 3.1t "
   Response = MsgBox(MSG1, True)
   ShowUpdateTribesVersion = 1
End If

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Forms![TRIBEVIBES]![Status] = "Determining Carrying Capacity for groups"
Forms![TRIBEVIBES].Repaint

' Group Movement
' Number of People @ 30lbs
' Number of Backpacks @ 30lbs (if not more than People)
' Slaves are people
' Number of Wagons @ 2000lbs
' Number of Ridden Horses/Camels @ 100lbs
' Number of Unridden Horses/Camels @ 300 lbs
' Elephants: carry 1000 unridden (800 ridden by 1 person - or 3 people may ride with no gear)
' Number of Palaquins @ 300lbs req 4 people plus loss of people carry capacity.
' Number of Saddlebags @ 100lbs

CAPACITY = 0
Walking_Capacity = 0

If Capacity_Type = "Group" Then
   Set TRIBESTABLE = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
   TRIBESTABLE.index = "PRIMARYKEY"
   TRIBESTABLE.MoveFirst
   TRIBESTABLE.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER

   If TRIBESTABLE![Village] = "FLEET" Then
      FLEET = "YES"
   End If
   
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
  
AVAILABLE_PEOPLE = 0
If TRIBESTABLE.NoMatch Then
   TotalPeople = 0
Else
   TotalPeople = TRIBESTABLE![WARRIORS] + TRIBESTABLE![ACTIVES] + TRIBESTABLE![INACTIVES]
   TotalPeople = TotalPeople + TRIBESTABLE![SLAVE]
   AVAILABLE_PEOPLE = TotalPeople
End If

Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_GOODS")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "FINISHED", "BACKPACK"

If Not TRIBESGOODS.NoMatch Then
   TotalBackpacks = TRIBESGOODS![ITEM_NUMBER]
Else
   TotalBackpacks = 0
End If

TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "FINISHED", "PALAQUIN"
VALIDGOODS.Seek "=", "PALAQUIN"

If Not TRIBESGOODS.NoMatch Then
   TotalPalaquins = TRIBESGOODS![ITEM_NUMBER]
Else
   TotalPalaquins = 0
End If

If Not VALIDGOODS.NoMatch Then
   Walking_Capacity = Walking_Capacity + (TotalPeople - (TotalPalaquins * 4)) * 30 + TotalPalaquins * VALIDGOODS![CARRIES]
Else
   Walking_Capacity = Walking_Capacity + TotalPeople * 30
   MSG1 = "Palaquins are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If


VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "BACKPACK"

If Not VALIDGOODS.NoMatch Then
   If TotalBackpacks <= TotalPeople Then
      Walking_Capacity = Walking_Capacity + (TotalBackpacks * VALIDGOODS![CARRIES])
      TotalBackpacksUsed = TotalBackpacks
   Else
      Walking_Capacity = Walking_Capacity + (TotalPeople * VALIDGOODS![CARRIES])
      TotalBackpacksUsed = TotalPeople
   End If
Else
   MSG1 = "Backpacks are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "FINISHED", "WAGON"

If Not TRIBESGOODS.NoMatch Then
   TotalWagons = TRIBESGOODS![ITEM_NUMBER]
Else
   TotalWagons = 0
End If

WagonsNumber = TotalWagons

VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "WAGON"
If Not VALIDGOODS.NoMatch Then
      Wagons_Carry = VALIDGOODS![CARRIES]
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "ANIMAL", "ELEPHANT"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "ELEPHANT"

If Not VALIDGOODS.NoMatch Then
   ElephantCapacity = VALIDGOODS![CARRIES]
   If Not TRIBESGOODS.NoMatch Then
      TotalElephants = TRIBESGOODS![ITEM_NUMBER]
      If TotalWagons > 0 Then
         If TotalWagons > TotalElephants Then
            Walking_Capacity = Walking_Capacity + (TotalElephants * Wagons_Carry)
            TotalWagons = TotalWagons - TotalElephants
            TotalElephants = 0
         Else
            Walking_Capacity = Walking_Capacity + (TotalWagons * Wagons_Carry)
            Walking_Capacity = Walking_Capacity + ((TotalElephants - TotalWagons) * ElephantCapacity)
            ' Elephants must be able to carry all wagons to maintain mounted movement
            Mounted_Capacity = Mounted_Capacity + ((TotalElephants - TotalWagons) * ElephantCapacity)
            TotalElephants = TotalElephants - TotalWagons
            TotalWagons = 0
         End If
      Else
         Walking_Capacity = Walking_Capacity + (TotalElephants * ElephantCapacity)
         'Mounted_Capacity = Mounted_Capacity + (TotalElephants * ElephantCapacity)
      End If
   Else
      TotalElephants = 0
   End If
Else
   ElephantCapacity = 0
   MSG1 = "ELEPHANT are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "ANIMAL", "CATTLE"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "WAGON"

If Not TRIBESGOODS.NoMatch Then
   TotalCattle = TRIBESGOODS![ITEM_NUMBER]
Else
   TotalCattle = 0
End If

If Not VALIDGOODS.NoMatch Then
   If (TotalCattle / 2) >= TotalWagons Then
      Walking_Capacity = Walking_Capacity + (TotalWagons * Wagons_Carry)
      TotalWagons = 0
   Else
      Walking_Capacity = Walking_Capacity + ((TotalCattle / 2) * Wagons_Carry)
      TotalWagons = TotalWagons - (TotalCattle / 2)
   End If
Else
   MSG1 = "Wagons are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "ANIMAL", "HORSE"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "HORSE"

If Not TRIBESGOODS.NoMatch Then
   TotalHorses = TRIBESGOODS![ITEM_NUMBER]
Else
   TotalHorses = 0
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "ANIMAL", "HORSE/HEAVY"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "HORSE/HEAVY"

If Not VALIDGOODS.NoMatch Then
   If Not TRIBESGOODS.NoMatch Then
      TotalHHorses = TRIBESGOODS![ITEM_NUMBER]
      Walking_Capacity = Walking_Capacity + (TotalLHorses * VALIDGOODS![CARRIES])
   Else
      TotalHHorses = 0
   End If
Else
   MSG1 = "HORSE/HEAVY are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "ANIMAL", "HORSE/LIGHT"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "HORSE/LIGHT"

If Not TRIBESGOODS.NoMatch Then
   TotalLHorses = TRIBESGOODS![ITEM_NUMBER]
   Walking_Capacity = Walking_Capacity + (TotalLHorses * VALIDGOODS![CARRIES])
Else
   TotalLHorses = 0
End If

VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "HORSE"

If Not VALIDGOODS.NoMatch Then
   HorseCapacity = VALIDGOODS![CARRIES]
   If (TotalHorses / 2) >= TotalWagons Then
      Walking_Capacity = Walking_Capacity + (TotalWagons * Wagons_Carry)
      TotalHorses = TotalHorses - TotalWagons * 2
      Walking_Capacity = Walking_Capacity + (TotalHorses * HorseCapacity)
      'Mounted_Capacity = Mounted_Capacity + ((TotalHorses * HorseCapacity) / 2)
      TotalWagons = 0
   Else
      Walking_Capacity = Walking_Capacity + ((TotalHorses / 2) * Wagons_Carry)
      TotalWagons = TotalWagons - (TotalHorses / 2)
      TotalHorses = TotalHorses - (TotalHorses / 2) * 2
      Walking_Capacity = Walking_Capacity + (TotalHorses * HorseCapacity)
      'Mounted_Capacity = Mounted_Capacity + ((TotalHorses * HorseCapacity) / 2)
   End If
Else
   MSG1 = "HORSE/LIGHT are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "ANIMAL", "CAMEL"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "CAMEL"

If Not VALIDGOODS.NoMatch Then
   If Not TRIBESGOODS.NoMatch Then
      CamelCapacity = VALIDGOODS![CARRIES]
      TotalCamels = TRIBESGOODS![ITEM_NUMBER]
      Walking_Capacity = Walking_Capacity + (TotalCamels * CamelCapacity)
      'Mounted_Capacity = Mounted_Capacity + (TotalCamels * CamelCapacity)
   Else
      TotalCamels = 0
   End If
Else
   CamelCapacity = 0
   MSG1 = "CAMEL are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If

TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "FINISHED", "SADDLEBAG"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "SADDLEBAG"

If Not TRIBESGOODS.NoMatch Then
   TotalSaddleBags = TRIBESGOODS![ITEM_NUMBER]
Else
   TotalSaddleBags = 0
End If

If Not VALIDGOODS.NoMatch Then
   If TotalSaddleBags <= (TotalHorses + TotalCamels * 2) Then
      Walking_Capacity = Walking_Capacity + (TotalSaddleBags * VALIDGOODS![CARRIES])
      Mounted_Capacity = Mounted_Capacity + (TotalSaddleBags * VALIDGOODS![CARRIES])
   Else
      Walking_Capacity = Walking_Capacity + ((TotalHorses + TotalCamels * 2) * VALIDGOODS![CARRIES])
      Mounted_Capacity = Mounted_Capacity + ((TotalHorses + TotalCamels * 2) * VALIDGOODS![CARRIES])
   End If
Else
   MSG1 = "SADDLEBAG are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If


VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", "PALAQUIN"
If Not VALIDGOODS.NoMatch Then
   If TotalPalaquins <= TotalPeople Then
      Walking_Capacity = Walking_Capacity + (TotalPalaquins * VALIDGOODS![CARRIES])
      TotalPalaquins = 0
   Else
      Walking_Capacity = Walking_Capacity + (TotalPeople * VALIDGOODS![CARRIES])
      TotalPalaquins = 0
   End If
Else
   MSG1 = "PALAQUINS are not a Valid_Good"
   MSG2 = "Please add it in"
   Response = MsgBox(MSG1 & MSG2, True)
End If


' Mounted capacity calculation
'ridden horses    capacity is 100lb of  300lb
'ridden camels    capacity is 400lb of  500lb
'ridden elephants capacity is 800lb of 1000lb

If TotalPeople <= TotalHorses Then
        TotalHorses = TotalHorses - TotalPeople
        Mounted_Capacity = Mounted_Capacity + (TotalPeople * HorseCapacity) / 3
        TotalPeople = 0
Else
        TotalPeople = TotalPeople - TotalHorses
        Mounted_Capacity = Mounted_Capacity + (TotalHorses * HorseCapacity) / 3
        TotalHorses = 0
End If

If TotalPeople <= TotalCamels Then
        TotalCamels = TotalCamels - TotalPeople
        Mounted_Capacity = Mounted_Capacity + (TotalPeople * CamelCapacity) * 0.8
        TotalPeople = 0
Else
        TotalPeople = TotalPeople - TotalCamels
        Mounted_Capacity = Mounted_Capacity + (TotalCamels * CamelCapacity) * 0.8
        TotalCamels = 0
End If

If TotalPeople <= TotalElephants Then
        TotalElephants = TotalElephants - TotalPeople
        Mounted_Capacity = Mounted_Capacity + (TotalPeople * ElephantCapacity) * 0.8
        TotalPeople = 0
Else
        TotalPeople = TotalPeople - TotalCamels
        Mounted_Capacity = Mounted_Capacity + (TotalElephants * ElephantCapacity) * 0.8
        TotalElephants = 0
End If


Mounted_Capacity = Mounted_Capacity + (TotalHorses * HorseCapacity) _
       + (TotalCamels * CamelCapacity) + (TotalElephants * ElephantCapacity)




If FLEET = "YES" Then
    ' if fleet then calc people space as cargo space.
    ' coaster   - crew 3,   max 20
    ' fisher    - crew 6,   max 8
    ' sm galley - crew 48,  max 65
    ' md galley - crew 72,  max 100
    ' lg galley - crew 120, max 150
    ' trader    - crew 12,  max 80
    ' longship  - crew 10,  max 100
    ' merchant  - crew 10,  max 60
    ' warship   - crew 10,  max 60
    ' WHAT ABOUT ACTUAL CARGO SPACE????????? -- NOT CATERED FOR????
     
    Set SHIPSTABLE = TVDB.OpenRecordset("VALID_SHIPS")
    SHIPSTABLE.index = "PRIMARYKEY"
    SHIPSTABLE.MoveFirst
    
    Required_Crew = 0
    Max_Crew = 0
    Mounted_Capacity = 0
    
    Do Until SHIPSTABLE.EOF
       TRIBESGOODS.MoveFirst
       TRIBESGOODS.Seek "=", DC_CLANNUMBER, DC_TRIBENUMBER, "SHIP", SHIPSTABLE![VESSEL]
       If Not TRIBESGOODS.NoMatch Then
          If TRIBESGOODS![ITEM_NUMBER] > 0 Then
             Mounted_Capacity = Mounted_Capacity + (SHIPSTABLE![Cargo_Space] * TRIBESGOODS![ITEM_NUMBER])
             Required_Crew = Required_Crew + (SHIPSTABLE![Crew_Required] * TRIBESGOODS![ITEM_NUMBER])
             Max_Crew = Max_Crew + (SHIPSTABLE![Max_Crew] * TRIBESGOODS![ITEM_NUMBER])
          End If
       End If
     
       SHIPSTABLE.MoveNext
    Loop
' Take into account empty passenger spaces Alex D 10.09 24
    If (AVAILABLE_PEOPLE < Max_Crew) Then
       Mounted_Capacity = Mounted_Capacity + (Max_Crew - AVAILABLE_PEOPLE) * 500
    End If
    
'    EXCESS_PEOPLE = 0
'
'    If AVAILABLE_PEOPLE > Required_Crew Then
'       EXCESS_PEOPLE = EXCESS_PEOPLE + (AVAILABLE_PEOPLE - Required_Crew)
'   End If
'
'    If EXCESS_PEOPLE >= 0 Then
'       Mounted_Capacity = Mounted_Capacity + (EXCESS_PEOPLE * 500)
'   End If
End If


'If Not TRIBESTABLE.NoMatch Then ' meaning it is a match
'   If Not IsNull(TRIBESTABLE![GOODS TRIBE]) And Not (TRIBESTABLE![CLAN] = TRIBESTABLE![TRIBE]) Then
'      TRIBESTABLE.Seek "=", CLANNUMBER, TRIBESTABLE![GOODS TRIBE]
'      If TRIBESTABLE.NoMatch Then
'         'whoops
'      ElseIf Not TRIBESTABLE![GOODS TRIBE] = TRIBESTABLE![TRIBE] Then
'         TRIBESTABLE.Edit
'         TRIBESTABLE![CAPACITY] = TRIBESTABLE![CAPACITY] + CAPACITY
'         TRIBESTABLE![Walking_Capacity] = TRIBESTABLE![Walking_Capacity] + Walking_Capacity
'         TRIBESTABLE.UPDATE
'      End If
'   Else
'      TRIBESTABLE.Edit
'      TRIBESTABLE![CAPACITY] = Mounted_Capacity
'      TRIBESTABLE![Walking_Capacity] = Walking_Capacity
'      TRIBESTABLE.UPDATE
'   End If
'End If
      TRIBESTABLE.Edit
      TRIBESTABLE![CAPACITY] = Mounted_Capacity
      TRIBESTABLE![Walking_Capacity] = Walking_Capacity
      TRIBESTABLE.UPDATE

TRIBESTABLE.Close

End If

Determine_Capacities_Error_CLOSE:
Forms![TRIBEVIBES]![Status] = " "
Forms![TRIBEVIBES].Repaint
Exit Function


Determine_Capacities_Error:
If (Err = 6) Then
   Resume Next

Else
   Call A999_ERROR_HANDLING
   Resume Determine_Capacities_Error_CLOSE
End If

End Function

'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 16/10/17  *  Adding in maintenance log                                      **'
'**           *  Note: if weight is not in valid goods table then it will not   **'
'**           *  be picked up                                                   **'
'**           *                                                                 **'
'*===============================================================================*'
Public Function Determine_Weights(DW_CLANNUMBER, DW_TRIBENUMBER)
Dim WEIGHT As Double
Dim VALIDGOODS As Recordset
Dim TRIBESGOODS As Recordset
Dim count As Integer

Set TVWKSPACE = DBEngine.Workspaces(0)
Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb")
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Forms![TRIBEVIBES]![Status] = "Determining Current Weight for group"
Forms![TRIBEVIBES].Repaint
    
Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_GOODS")
TRIBESGOODS.index = "CLANTRIBE"
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", DW_CLANNUMBER, DW_TRIBENUMBER

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "PRIMARYKEY"

If Not TRIBESGOODS.NoMatch Then
   Do Until Not TRIBESGOODS![TRIBE] = DW_TRIBENUMBER
      VALIDGOODS.MoveFirst
      VALIDGOODS.Seek "=", TRIBESGOODS![ITEM]
      If Not VALIDGOODS.NoMatch Then
         If InStr(TRIBENUMBER, "F") Then
            If VALIDGOODS![TABLE] = "ANIMAL" Then
               If IsNull(VALIDGOODS![WEIGHT]) Then
                  WEIGHT = WEIGHT + 0
               Else
                  WEIGHT = WEIGHT + TRIBESGOODS![ITEM_NUMBER] * VALIDGOODS![WEIGHT]
               End If
            Else
               'DONT ADD WEIGHT OF ANIMALS UNLESS IN FLEET
            End If
         Else
            If VALIDGOODS![TABLE] = "ANIMAL" Then
               'DONT ADD WEIGHT OF ANIMALS UNLESS IN FLEET
            ElseIf IsNull(VALIDGOODS![WEIGHT]) Then
               WEIGHT = WEIGHT + 0
            Else
               WEIGHT = WEIGHT + TRIBESGOODS![ITEM_NUMBER] * VALIDGOODS![WEIGHT]
            End If
         End If
      End If
      TRIBESGOODS.MoveNext
      If TRIBESGOODS.EOF Then
         Exit Do
      End If
   Loop
End If

   Set TRIBESTABLE = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
   TRIBESTABLE.index = "PRIMARYKEY"
   TRIBESTABLE.MoveFirst
   TRIBESTABLE.Seek "=", DW_CLANNUMBER, DW_TRIBENUMBER

If Not TRIBESTABLE.NoMatch Then
   TRIBESTABLE.Edit
   TRIBESTABLE![WEIGHT] = WEIGHT
   TRIBESTABLE.UPDATE
End If

TRIBESTABLE.Close
Forms![TRIBEVIBES]![Status] = " "
Forms![TRIBEVIBES].Repaint

End Function
