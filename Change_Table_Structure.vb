Attribute VB_Name = "Change_Table_Structure"
Option Compare Database
Option Explicit




Public Function Delete_old_HEX_MAP_fields()
On Error GoTo ERR_Delete_old_HEX_MAP_fields
TRIBE_STATUS = "Delete old hexmap fields"

Dim COLUMN_NAME As String
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
TVDBGM.Execute "DROP INDEX Secondarykey ON HEX_MAP;"
COLUMN_NAME = "PL CLAN"
TVDBGM.Execute "alter table Scout_Movement drop column [MOVEMENT];"

TVDBGM.Execute "alter table HEX_MAP drop column [PL CLAN];"

TVDBGM.Execute "alter table HEX_MAP drop column [PL TRIBE];"

TVDBGM.Execute "alter table HEX_MAP drop column [PACIFICATION LEVEL];"

TVDBGM.Execute "alter table HEX_MAP drop column [POPULATION];"

TVDBGM.Execute "alter table HEX_MAP drop column [POP INCREASED];"

TVDBGM.Execute "alter table HEX_MAP drop column [SECOND ORE];"

TVDBGM.Execute "alter table HEX_MAP drop column [THIRD ORE];"

TVDBGM.Execute "alter table HEX_MAP drop column [FORTH ORE];"

TVDBGM.Execute "alter table HEX_MAP drop column [MINING];"

TVDBGM.Execute "alter table HEX_MAP drop column [SECOND MINING];"

TVDBGM.Execute "alter table HEX_MAP drop column [THIRD MINING];"

TVDBGM.Execute "alter table HEX_MAP drop column [FORTH MINING];"

TVDBGM.Close

    
ERR_Delete_old_HEX_MAP_fields_close:
   Exit Function

ERR_Delete_old_HEX_MAP_fields:
' 3010 Table already exists
' 3080 Field already exists on table
' 3375 Index already exists on table

If (Err = 3010) Then
   MsgBox "HEX_MAP Table Exists"
   Resume ERR_Delete_old_HEX_MAP_fields_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   'MsgBox (MSG)
   
   Resume Next
End If

End Function
Public Function Add_New_Field_To_Table1()
On Error GoTo ERR_Add_New_Field_To_Table
TRIBE_STATUS = "Add new fields to table1"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Adding Fields to the TRADING_POST_GOODS Table"
Forms![TRIBEVIBES].Repaint
    
' NEED TO REMOVE INDEXES FROM TABLE
TVDBGM.Execute "alter table TRADING_POST_GOODS DROP CONSTRAINT PRIMARYKEY;"
TVDBGM.Execute "alter table TRADING_POST_GOODS DROP CONSTRAINT TRIBE;"


' NEED TO ADD NEW FIELDS
    
TVDBGM.Execute "alter table TRADING_POST_GOODS add column TYPE_OF_TRADING_POST TEXT(5);"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column HEX_MAP_ID TEXT(2);"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column BUY_RESET_WAIT DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column NORMAL_BUY_LIMIT DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column TURNS_SINCE_LAST_BUY DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column BUY_THIS_TURN TEXT(1);"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column BUY_TOTAL DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column SELL_RESET_WAIT DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column NORMAL_SELL_LIMIT DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column TURNS_SINCE_LAST_SELL DOUBLE;"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column SELL_THIS_TURN TEXT(1);"
TVDBGM.Execute "alter table TRADING_POST_GOODS add column SELL_TOTAL DOUBLE;"

' NEED TO REBUILD EXISTING RECORDS.

Set TRIBESINFO = TVDBGM.OpenRecordset("TRADING_POST_GOODS")
TRIBESINFO.MoveFirst

Do While Not TRIBESINFO.EOF
   If Left(TRIBESINFO![TRIBE], 1) = "0" Or Left(TRIBESINFO![TRIBE], 1) = "1" _
   Or Left(TRIBESINFO![TRIBE], 1) = "2" Or Left(TRIBESINFO![TRIBE], 1) = "3" _
   Or Left(TRIBESINFO![TRIBE], 1) = "4" Or Left(TRIBESINFO![TRIBE], 1) = "5" _
   Or Left(TRIBESINFO![TRIBE], 1) = "6" Or Left(TRIBESINFO![TRIBE], 1) = "7" _
   Or Left(TRIBESINFO![TRIBE], 1) = "8" Or Left(TRIBESINFO![TRIBE], 1) = "9" Then
      TRIBESINFO.Edit
      TRIBESINFO![TYPE_OF_TRADING_POST] = "TRIBE"
      TRIBESINFO![HEX_MAP_ID] = "BA"
      TRIBESINFO.UPDATE
   ElseIf TRIBESINFO![TRIBE] = "GM SALE" Then
      TRIBESINFO.Edit
      TRIBESINFO![TYPE_OF_TRADING_POST] = "SALE"
      TRIBESINFO![HEX_MAP_ID] = "BA"
      TRIBESINFO![BUY_RESET_WAIT] = 0
      TRIBESINFO![NORMAL_BUY_LIMIT] = TRIBESINFO![BUY LIMIT]
      TRIBESINFO![TURNS_SINCE_LAST_BUY] = 0
      TRIBESINFO![BUY_THIS_TURN] = "N"
      TRIBESINFO![BUY_TOTAL] = 0
      TRIBESINFO![SELL_RESET_WAIT] = 0
      TRIBESINFO![NORMAL_SELL_LIMIT] = TRIBESINFO![SELL LIMIT]
      TRIBESINFO![TURNS_SINCE_LAST_SELL] = 0
      TRIBESINFO![SELL_THIS_TURN] = "N"
      TRIBESINFO![SELL_TOTAL] = 0
      TRIBESINFO.UPDATE
   ElseIf TRIBESINFO![TRIBE] = "FAIR" Then
      TRIBESINFO.Edit
      TRIBESINFO![TYPE_OF_TRADING_POST] = "FAIR"
      TRIBESINFO![HEX_MAP_ID] = "BA"
      TRIBESINFO![BUY_RESET_WAIT] = 0
      TRIBESINFO![NORMAL_BUY_LIMIT] = TRIBESINFO![BUY LIMIT]
      TRIBESINFO![TURNS_SINCE_LAST_BUY] = 0
      TRIBESINFO![BUY_THIS_TURN] = "N"
      TRIBESINFO![BUY_TOTAL] = 0
      TRIBESINFO![SELL_RESET_WAIT] = 0
      TRIBESINFO![NORMAL_SELL_LIMIT] = TRIBESINFO![SELL LIMIT]
      TRIBESINFO![TURNS_SINCE_LAST_SELL] = 0
      TRIBESINFO![SELL_THIS_TURN] = "N"
      TRIBESINFO![SELL_TOTAL] = 0
      TRIBESINFO.UPDATE
   Else
      TRIBESINFO.Edit
      TRIBESINFO![TYPE_OF_TRADING_POST] = "CITY"
      TRIBESINFO![HEX_MAP_ID] = "BA"
      TRIBESINFO![BUY_RESET_WAIT] = 0
      TRIBESINFO![NORMAL_BUY_LIMIT] = TRIBESINFO![BUY LIMIT]
      TRIBESINFO![TURNS_SINCE_LAST_BUY] = 0
      TRIBESINFO![BUY_THIS_TURN] = "N"
      TRIBESINFO![BUY_TOTAL] = 0
      TRIBESINFO![SELL_RESET_WAIT] = 0
      TRIBESINFO![NORMAL_SELL_LIMIT] = TRIBESINFO![SELL LIMIT]
      TRIBESINFO![TURNS_SINCE_LAST_SELL] = 0
      TRIBESINFO![SELL_THIS_TURN] = "N"
      TRIBESINFO![SELL_TOTAL] = 0
      TRIBESINFO.UPDATE
   End If
   TRIBESINFO.MoveNext
   If TRIBESINFO.EOF Then
      Exit Do
   End If
Loop

TRIBESINFO.Close

' NEED TO REBUILD INDEXES.

TVDBGM.Execute "CREATE INDEX PrimaryKey ON TRADING_POST_GOODS " _
        & "(TYPE_OF_TRADING_POST,HEX_MAP_ID,TRIBE,GOOD) WITH PRIMARY;"

TVDBGM.Execute "CREATE INDEX TRIBESGOOD ON TRADING_POST_GOODS " _
        & "(TRIBE,GOOD);"

TVDBGM.Execute "CREATE INDEX TRIBE ON TRADING_POST_GOODS " _
        & "(TRIBE);"

TVDBGM.Execute "CREATE INDEX HEX_MAP_ID ON TRADING_POST_GOODS " _
        & "(HEX_MAP_ID,GOOD);"

Forms![TRIBEVIBES]![Status] = "Adding Fields to the Hex Map Table"
Forms![TRIBEVIBES].Repaint
   
TVDBGM.Close

ERR_Add_New_Field_To_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Add_New_Field_To_Table:
If (Err = 3380) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "Field exists on table"
   End If
   Resume Next
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function


Public Function Create_Process_Tribes_Activity_Table()
On Error GoTo ERR_Create_Process_Tribes_Activity_Table
TRIBE_STATUS = "Create Process Tribes Activity Table"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating Process_Tribes_Activity_Table"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE Process_Tribes_Activity " _
  & "(CLAN TEXT(10),TRIBE TEXT(10), ORDER DOUBLE, ACTIVITY TEXT (50), ITEM TEXT(50)," _
  & "DISTINCTION TEXT(20),PEOPLE DOUBLE ,SLAVES DOUBLE,SPECIALISTS DOUBLE,JOINT TEXT(1)," _
  & "OWNING_CLAN TEXT(10),OWNING_TRIBE TEXT(10),NUMBER_OF_SEEKING_GROUPS DOUBLE," _
  & "WHALE_SIZE TEXT(1),MINING_DIRECTION TEXT(6),Processed TEXT(1));"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Process_Tribes_Activity " _
        & "(CLAN,TRIBE,ORDER) WITH PRIMARY;"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Activity", "Process_Tribes_Activity"
' COPY DATA FROM OTHER TABLES

ERR_Create_Process_Tribes_Activity_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Process_Tribes_Activity_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR_Create_Process_Tribes_Activity_Table Exists"
   End If
   Resume ERR_Create_Process_Tribes_Activity_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If



End Function

Public Function Create_Seeking_Returns_Table()
On Error GoTo ERR_Create_Seeking_Returns_Table
TRIBE_STATUS = "Create Seeking Returns Table"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating Seeking_Returns_Table"
Forms![TRIBEVIBES].Repaint

TVDBGM.Execute "CREATE TABLE Seeking_Returns_Table " _
  & "(ITEM TEXT(50),MODIFIER DOUBLE);"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Seeking_Returns_Table " _
        & "(ITEM) WITH PRIMARY;"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Seeking_Returns_Table", "Seeking_Returns_Table"
' ADD DATA TO TABLE

Set TRIBESINFO = TVDBGM.OpenRecordset("Seeking_Returns_Table")
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "BARK"
TRIBESINFO![Modifier] = 3
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "CAMEL"
TRIBESINFO![Modifier] = 67
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "CATTLE"
TRIBESINFO![Modifier] = 35
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "DOG"
TRIBESINFO![Modifier] = 87
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "ELEPHANT"
TRIBESINFO![Modifier] = 77
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "GOAT"
TRIBESINFO![Modifier] = 5
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "HERB"
TRIBESINFO![Modifier] = 2.5
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "HIVE"
TRIBESINFO![Modifier] = 75
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "HONEY"
TRIBESINFO![Modifier] = 7
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "HORSE"
TRIBESINFO![Modifier] = 67
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "LIMESTONE"
TRIBESINFO![Modifier] = 30
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "LOG"
TRIBESINFO![Modifier] = 45
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "RECRUITS"
TRIBESINFO![Modifier] = 45
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "SALTPETER"
TRIBESINFO![Modifier] = 30
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "SHEEP"
TRIBESINFO![Modifier] = 5
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "SLAVE"
TRIBESINFO![Modifier] = 45
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "SPICE"
TRIBESINFO![Modifier] = 55
TRIBESINFO.UPDATE
TRIBESINFO.AddNew
TRIBESINFO![ITEM] = "WAX"
TRIBESINFO![Modifier] = 15
TRIBESINFO.UPDATE
TRIBESINFO.Close

ERR_Create_Seeking_Returns_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Seeking_Returns_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR_Create_Seeking_Returns_Table"
   End If
   Resume ERR_Create_Seeking_Returns_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function


Public Function Add_New_Field_To_Table()
On Error GoTo ERR_Add_New_Field_To_Table
TRIBE_STATUS = "Add new field to table"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
'Forms![TRIBEVIBES]![Status] = "Adding Fields to the Tribes_Processing Table"
'Forms![TRIBEVIBES].Repaint
    
' NEED TO ADD NEW FIELDS
    
TVDBGM.Execute "alter table Tribes_Processing add column Number_Of_Slaves_Overseen DOUBLE;"
Dim tdfTribes_Processing As TableDef
Set tdfTribes_Processing = TVDBGM.TableDefs!Tribes_Processing
tdfTribes_Processing.Fields!Number_Of_Slaves_Overseen.DefaultValue = 0

TVDBGM.Execute "alter table Completed_Research add column Completed_This_Turn text(1);"
Dim tdfCompleted_Research As TableDef
Set tdfCompleted_Research = TVDBGM.TableDefs!Completed_Research
tdfCompleted_Research.Fields!COMPLETED_THIS_TURN.DefaultValue = "N"

TVDBGM.Execute "alter table Tribes_Specialists add column Number_Of_Turns_Training DOUBLE;"
Dim tdfTribes_Specialists As TableDef
Set tdfTribes_Specialists = TVDBGM.TableDefs!Tribes_Specialists
tdfTribes_Specialists.Fields!NUMBER_OF_TURNS_TRAINING.DefaultValue = 0
tdfTribes_Specialists.Fields!SPECIALISTS_USED.DefaultValue = 0

TVDBGM.Execute "alter table Under_Construction add column MILLSTONE DOUBLE;"
Dim tdfUnder_Construction As TableDef
Set tdfUnder_Construction = TVDBGM.TableDefs!Under_Construction
tdfUnder_Construction.Fields!MILLSTONE.DefaultValue = 0

TVDBGM.Execute "alter table Tribes_General_Info add column Walking_Capacity DOUBLE;"

Dim tdfTribes_General_Info As TableDef
Set tdfTribes_General_Info = TVDBGM.TableDefs!TRIBES_GENERAL_INFO
tdfTribes_General_Info.Fields!Walking_Capacity.DefaultValue = 0

Dim tdfHEXMAP_Permanent_FARMING As TableDef
Set tdfHEXMAP_Permanent_FARMING = TVDBGM.TableDefs!HEXMAP_Permanent_FARMING
tdfHEXMAP_Permanent_FARMING.Fields!ITEM_NUMBER.DefaultValue = 0
tdfHEXMAP_Permanent_FARMING.Fields!HARVESTED.DefaultValue = 0

' NEED TO REBUILD INDEXES.

Forms![TRIBEVIBES]![Status] = "Adding Fields to the Tribes_Processing Table"
Forms![TRIBEVIBES].Repaint
   
TVDBGM.Close

ERR_Add_New_Field_To_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Add_New_Field_To_Table:
If (Err = 3380) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "Field exists on table"
   End If
   Resume Next
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function

Public Function Create_Hexmap_Farming_Table()
On Error GoTo ERR_Create_Process_Tribes_Activity_Table

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating HEXMAP_FARMING"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE HEXMAP_FARMING " _
  & "(HEXMAP TEXT(7),CLAN TEXT(10), TRIBE TEXT(10), TURN TEXT(6), ITEM TEXT (20), ITEM_NUMBER DOUBLE);"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON HEXMAP_FARMING " _
        & "(HEXMAP,CLAN,TRIBE,TURN,ITEM);"

TVDBGM.Execute "CREATE INDEX TRIBE ON HEXMAP_FARMING " _
        & "(HEXMAP,CLAN,TRIBE);"

TVDBGM.Execute "CREATE INDEX HEXMAP ON HEXMAP_FARMING " _
        & "(HEXMAP);"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEXMAP_FARMING", "HEXMAP_FARMING"
' COPY DATA FROM OTHER TABLES

Forms![TRIBEVIBES]![Status] = "Creating HEXMAP_Permanent_FARMING"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE HEXMAP_Permanent_FARMING " _
  & "(HEXMAP TEXT(7),CLAN TEXT(10), TRIBE TEXT(10), ITEM TEXT (20), ITEM_NUMBER DOUBLE);"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON HEXMAP_Permanent_FARMING " _
        & "(HEXMAP,CLAN,TRIBE,TURN,ITEM);"

TVDBGM.Execute "CREATE INDEX TRIBE ON HEXMAP_Permanent_FARMING " _
        & "(HEXMAP,CLAN,TRIBE);"

TVDBGM.Execute "CREATE INDEX HEXMAP ON HEXMAP_Permanent_FARMING " _
        & "(HEXMAP);"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEXMAP_Permanent_FARMING", "HEXMAP_Permanent_FARMING"
' COPY DATA FROM OTHER TABLES

ERR_Create_Process_Tribes_Activity_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Process_Tribes_Activity_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR_HEXMAP_FARMING Exists"
   End If
   Resume ERR_Create_Process_Tribes_Activity_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If




End Function

Public Function Create_Games_Weather_Table()
On Error GoTo ERR_Create_Process_Tribes_Activity_Table

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating GAMES_WEATHER"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE GAMES_WEATHER " _
  & "(WEATHER_ZONE TEXT(10),TURN TEXT(6), WEATHER TEXT (50));"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON GAMES_WEATHER " _
        & "(WEATHER_ZONE,TURN) WITH PRIMARY;"

TVDBGM.Execute "CREATE INDEX WEATHER_ZONE ON GAMES_WEATHER " _
        & "(WEATHER_ZONE);"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GAMES_WEATHER", "GAMES_WEATHER"
' COPY DATA FROM OTHER TABLES

ERR_Create_Process_Tribes_Activity_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Process_Tribes_Activity_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR GAMES_WEATHER Exists"
   End If
   Resume ERR_Create_Process_Tribes_Activity_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If


End Function

Public Function Create_GM_Costs_Table()

On Error GoTo ERR_Create_GM_Costs_Table

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating GM_Costs_Table"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE GM_Costs_Table " _
  & "(GROUP TEXT(10) ,COST SINGLE);"


' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "GM_Costs_Table", "GM_Costs_Table"
' COPY DATA FROM OTHER TABLES

Call Populate_GM_Costs_Table

ERR_Create_GM_Costs_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_GM_Costs_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR GM_Costs_Table Exists"
   End If
   Resume ERR_Create_GM_Costs_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function

Public Function Create_Hexmap_Permanent_Farming_Table()
On Error GoTo ERR_Create_Process_Tribes_Activity_Table

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating HEXMAP_Permanent_FARMING"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE HEXMAP_Permanent_FARMING " _
  & "(HEXMAP TEXT(7),CLAN TEXT(10), TRIBE TEXT(10), ITEM TEXT (20), ITEM_NUMBER DOUBLE, HARVESTED double);"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON HEXMAP_Permanent_FARMING " _
        & "(HEXMAP,CLAN,TRIBE,ITEM);"

TVDBGM.Execute "CREATE INDEX TRIBE ON HEXMAP_Permanent_FARMING " _
        & "(HEXMAP,CLAN,TRIBE);"

TVDBGM.Execute "CREATE INDEX HEXMAP ON HEXMAP_Permanent_FARMING " _
        & "(HEXMAP);"

Dim tdfHEXMAP_Permanent_FARMING As TableDef
Set tdfHEXMAP_Permanent_FARMING = TVDBGM.TableDefs!HEXMAP_Permanent_FARMING
tdfHEXMAP_Permanent_FARMING.Fields!ITEM_NUMBER.DefaultValue = 0
tdfHEXMAP_Permanent_FARMING.Fields!HARVESTED.DefaultValue = 0

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "HEXMAP_Permanent_FARMING", "HEXMAP_Permanent_FARMING"
' COPY DATA FROM OTHER TABLES

ERR_Create_Process_Tribes_Activity_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Process_Tribes_Activity_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR_HEXMAP_FARMING Exists"
   End If
   Resume ERR_Create_Process_Tribes_Activity_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function

Public Function Create_Turn_Info_Reqd_Next_Turn()
On Error GoTo ERR_Create_Turn_Info_Reqd_Next_Turn

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating Turn_Info_Reqd_Next_Turn"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE Turn_Info_Reqd_Next_Turn " _
  & "(CLAN TEXT(10), TRIBE TEXT(10), ITEM TEXT (20), ITEM_NUMBER DOUBLE);"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Turn_Info_Reqd_Next_Turn " _
        & "(CLAN,TRIBE,ITEM) WITH PRIMARY;"

Dim tdfTurn_Info_Reqd_Next_Turn As TableDef
Set tdfTurn_Info_Reqd_Next_Turn = TVDBGM.TableDefs!Turn_Info_Reqd_Next_Turn
tdfTurn_Info_Reqd_Next_Turn.Fields!ITEM_NUMBER.DefaultValue = 0

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Turn_Info_Reqd_Next_Turn", "Turn_Info_Reqd_Next_Turn"
' COPY DATA FROM OTHER TABLES

ERR_Create_Turn_Info_Reqd_Next_Turn_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Turn_Info_Reqd_Next_Turn:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR Turn_Info_Reqd_Next_Turn Exists"
   End If
   Resume ERR_Create_Turn_Info_Reqd_Next_Turn_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If


End Function

Public Function Create_Table_For_Perm_Messages()
On Error GoTo ERR_Create_Table_For_Perm_Messages

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating Permanent_Messages_Table"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE Permanent_Messages_Table " _
  & "(CLAN TEXT(10), TRIBE TEXT(10));"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Permanent_Messages_Table " _
        & "(CLAN,TRIBE) WITH PRIMARY;"

TVDBGM.Execute "alter table Permanent_Messages_Table add column MESSAGE MEMO;"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Permanent_Messages_Table", "Permanent_Messages_Table"
' COPY DATA FROM OTHER TABLES

ERR_Create_Table_For_Perm_Messages_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Table_For_Perm_Messages:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR Turn_Info_Reqd_Next_Turn Exists"
   End If
   Resume ERR_Create_Table_For_Perm_Messages_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If
End Function

Public Function Create_Pacification_Table()
On Error GoTo ERR_Create_Pacification_Table

Dim fullfield As String
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating Pacification_Table"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE Pacification_Table " _
  & "(CLAN TEXT(10), TRIBE TEXT(10),Primary_Hex DOUBLE,GL1_1 DOUBLE, GL1_2 DOUBLE,GL1_3 DOUBLE,GL1_4 DOUBLE,GL1_5 DOUBLE,GL1_6 DOUBLE, " _
  & "GL2_1 DOUBLE,GL2_2 DOUBLE,GL2_3 DOUBLE,GL2_4 DOUBLE,GL2_5 DOUBLE,GL2_6 DOUBLE,GL2_7 DOUBLE,GL2_8 DOUBLE,GL2_9 DOUBLE,GL2_10 DOUBLE,GL2_11 DOUBLE,GL2_12 DOUBLE, " _
  & "GL3_1 DOUBLE,GL3_2 DOUBLE,GL3_3 DOUBLE,GL3_4 DOUBLE,GL3_5 DOUBLE,GL3_6 DOUBLE,GL3_7 DOUBLE,GL3_8 DOUBLE,GL3_9 DOUBLE,GL3_10 DOUBLE, " _
  & "GL3_11 DOUBLE,GL3_12 DOUBLE,GL3_13 DOUBLE,GL3_14 DOUBLE,GL3_15 DOUBLE,GL3_16 DOUBLE,GL3_17 DOUBLE,GL3_18 DOUBLE,GL4_1 DOUBLE,GL4_2 DOUBLE, " _
  & "GL4_3 DOUBLE,GL4_4 DOUBLE,GL4_5 DOUBLE,GL4_6 DOUBLE,GL4_7 DOUBLE,GL4_8 DOUBLE,GL4_9 DOUBLE,GL4_10 DOUBLE,GL4_11 DOUBLE,GL4_12 DOUBLE,GL4_13 DOUBLE, " _
  & "GL4_14 DOUBLE,GL4_15 DOUBLE,GL4_16 DOUBLE,GL4_17 DOUBLE,GL4_18 DOUBLE,GL4_19 DOUBLE,GL4_20 DOUBLE,GL4_21 DOUBLE,GL4_22 DOUBLE,GL4_23 DOUBLE,GL4_24 DOUBLE, " _
  & "GL5_1 DOUBLE,GL5_2 DOUBLE,GL5_3 DOUBLE,GL5_4 DOUBLE,GL5_5 DOUBLE,GL5_6 DOUBLE,GL5_7 DOUBLE,GL5_8 DOUBLE,GL5_9 DOUBLE,GL5_10 DOUBLE,GL5_11 DOUBLE, " _
  & "GL5_12 DOUBLE,GL5_13 DOUBLE,GL5_14 DOUBLE,GL5_15 DOUBLE,GL5_16 DOUBLE,GL5_17 DOUBLE,GL5_18 DOUBLE,GL5_19 DOUBLE,GL5_20 DOUBLE,GL5_21 DOUBLE,GL5_22 DOUBLE, " _
  & "GL5_23 DOUBLE,GL5_24 DOUBLE,GL5_25 DOUBLE,GL5_26 DOUBLE,GL5_27 DOUBLE,GL5_28 DOUBLE,GL5_29 DOUBLE,GL5_30 DOUBLE,GL6_1 DOUBLE,GL6_2 DOUBLE,GL6_3 DOUBLE, " _
  & "GL6_4 DOUBLE,GL6_5 DOUBLE,GL6_6 DOUBLE,GL6_7 DOUBLE,GL6_8 DOUBLE,GL6_9 DOUBLE,GL6_10 DOUBLE,GL6_11 DOUBLE,GL6_12 DOUBLE,GL6_13 DOUBLE,GL6_14 DOUBLE, " _
  & "GL6_15 DOUBLE,GL6_16 DOUBLE,GL6_17 DOUBLE,GL6_18 DOUBLE,GL6_19 DOUBLE,GL6_20 DOUBLE,GL6_21 DOUBLE,GL6_22 DOUBLE,GL6_23 DOUBLE,GL6_24 DOUBLE,GL6_25 DOUBLE, " _
  & "GL6_26 DOUBLE,GL6_27 DOUBLE,GL6_28 DOUBLE,GL6_29 DOUBLE,GL6_30 DOUBLE,GL6_31 DOUBLE,GL6_32 DOUBLE,GL6_33 DOUBLE,GL6_34 DOUBLE,GL6_35 DOUBLE,GL6_36 DOUBLE, " _
  & "GL7_1 DOUBLE,GL7_2 DOUBLE,GL7_3 DOUBLE,GL7_4 DOUBLE,GL7_5 DOUBLE,GL7_6 DOUBLE,GL7_7 DOUBLE,GL7_8 DOUBLE,GL7_9 DOUBLE,GL7_10 DOUBLE,GL7_11 DOUBLE, " _
  & "GL7_12 DOUBLE,GL7_13 DOUBLE,GL7_14 DOUBLE,GL7_15 DOUBLE,GL7_16 DOUBLE,GL7_17 DOUBLE,GL7_18 DOUBLE,GL7_19 DOUBLE,GL7_20 DOUBLE,GL7_21 DOUBLE,GL7_22 DOUBLE, " _
  & "GL7_23 DOUBLE,GL7_24 DOUBLE,GL7_25 DOUBLE,GL7_26 DOUBLE,GL7_27 DOUBLE,GL7_28 DOUBLE,GL7_29 DOUBLE,GL7_30 DOUBLE,GL7_31 DOUBLE,GL7_32 DOUBLE,GL7_33 DOUBLE, " _
  & "GL7_34 DOUBLE,GL7_35 DOUBLE,GL7_36 DOUBLE,GL7_37 DOUBLE,GL7_38 DOUBLE,GL7_39 DOUBLE,GL7_40 DOUBLE,GL7_41 DOUBLE,GL7_42 DOUBLE,GL8_1 DOUBLE,GL8_2 DOUBLE, " _
  & "GL8_3 DOUBLE,GL8_4 DOUBLE,GL8_5 DOUBLE,GL8_6 DOUBLE,GL8_7 DOUBLE,GL8_8 DOUBLE,GL8_9 DOUBLE,GL8_10 DOUBLE,GL8_11 DOUBLE,GL8_12 DOUBLE,GL8_13 DOUBLE, " _
  & "GL8_14 DOUBLE,GL8_15 DOUBLE,GL8_16 DOUBLE,GL8_17 DOUBLE,GL8_18 DOUBLE,GL8_19 DOUBLE,GL8_20 DOUBLE,GL8_21 DOUBLE,GL8_22 DOUBLE,GL8_23 DOUBLE,GL8_24 DOUBLE, " _
  & "GL8_25 DOUBLE,GL8_26 DOUBLE,GL8_27 DOUBLE,GL8_28 DOUBLE,GL8_29 DOUBLE,GL8_30 DOUBLE,GL8_31 DOUBLE,GL8_32 DOUBLE,GL8_33 DOUBLE,GL8_34 DOUBLE,GL8_35 DOUBLE, " _
  & "GL8_36 DOUBLE,GL8_37 DOUBLE,GL8_38 DOUBLE,GL8_39 DOUBLE,GL8_40 DOUBLE,GL8_41 DOUBLE,GL8_42 DOUBLE,GL8_43 DOUBLE,GL8_44 DOUBLE,GL8_45 DOUBLE,GL8_46 DOUBLE, " _
  & "GL8_47 DOUBLE,GL8_48 DOUBLE);"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Pacification_Table " _
        & "(CLAN,TRIBE) WITH PRIMARY;"

Dim tbfPacification_Table As TableDef
Set tbfPacification_Table = TVDBGM.TableDefs!PACIFICATION_TABLE
tbfPacification_Table.Fields!primary_hex.DefaultValue = 0
count = 1
fullfield = CStr("GL1_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL1_" & count)
     If count > 6 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL2_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL2_" & count)
     If count > 12 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL3_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL3_" & count)
     If count > 18 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL4_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL4_" & count)
     If count > 24 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL5_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL5_" & count)
     If count > 30 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL6_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL6_" & count)
     If count > 36 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL7_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL7_" & count)
     If count > 42 Then
        Exit Do
     End If
Loop
count = 1
fullfield = CStr("GL8_" & count)
Do
     tbfPacification_Table.Fields(fullfield).DefaultValue = 0
     count = count + 1
     fullfield = CStr("GL8_" & count)
     If count > 48 Then
        Exit Do
     End If
Loop


' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Pacification_Table", "Pacification_Table"
' COPY DATA FROM OTHER TABLES

ERR_Create_Pacification_Table_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Pacification_Table:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR Turn_Info_Reqd_Next_Turn Exists"
   End If
   Resume ERR_Create_Pacification_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function

Public Function Create_Activity_Copies()
On Error GoTo ERR_Create_Activity_Copies

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
Forms![TRIBEVIBES]![Status] = "Creating Create_Activity_Copies"
Forms![TRIBEVIBES].Repaint
    
TVDBGM.Execute "CREATE TABLE Process_Tribes_Item_allocation_Copy " _
  & "(CLAN TEXT(10), TRIBE TEXT(10), ACTIVITY TEXT(50), ITEM TEXT(50), ITEM_USED TEXT(50), QUANTITY DOUBLE, PROCESSED TEXT(1));"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Process_Tribes_Item_allocation_Copy " _
        & "(CLAN,TRIBE,ACTIVITY, ITEM, ITEM_USED) WITH PRIMARY;"

TVDBGM.Execute "CREATE TABLE Process_Tribes_Activity_Copy " _
  & "(CLAN TEXT(10), TRIBE TEXT(10), ORDER DOUBLE, ACTIVITY TEXT(50), ITEM TEXT(50), DISTINCTION TEXT(20), PEOPLE DOUBLE, SLAVES DOUBLE, SPECIALISTS DOUBLE," _
  & "JOINT TEXT(1), OWNING_CLAN TEXT(10), OWNING_TRIBE TEXT(10), NUMBER_OF_SEEKING_GROUPS DOUBLE, WHALE_SIZE TEXT(1), MINING_DIRECTION TEXT (6), PROCESSED TEXT (1));"

TVDBGM.Execute "CREATE INDEX PrimaryKey ON Process_Tribes_Activity_Copy " _
        & "(CLAN,TRIBE, ORDER) WITH PRIMARY;"

' MsgBox "THE NEW TABLE IS CREATED"

' ATTACH TABLE
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Activity_Copy", "Process_Tribes_Activity_Copy"
   DoCmd.TransferDatabase A_ATTACH, "MICROSOFT ACCESS", FILEGM, A_TABLE, "Process_Tribes_Item_allocation_Copy", "Process_Tribes_Item_allocation_Copy"
' COPY DATA FROM OTHER TABLES

ERR_Create_Activity_Copies_close:
   Forms![TRIBEVIBES]![Status] = ""
   Forms![TRIBEVIBES].Repaint
   Exit Function

ERR_Create_Activity_Copies:
If (Err = 3010) Then
   If GMTABLE![Name] = "JEFF" Then
      MsgBox "ERR Turn_Info_Reqd_Next_Turn Exists"
   End If
   Resume ERR_Create_Activity_Copies_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If

End Function
