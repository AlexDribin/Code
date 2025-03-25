Attribute VB_Name = "Update_Tables"
Option Compare Database
Option Explicit


Public Function Update_Tribes_Transfers_Table(Update_Type As String, New_Record As String)
' Update_Type will identify where the data is being supplied from eg from a Form or from
' a spreadsheet
' New_Record will identify if this is to be an update of an existing record or not.

On Error GoTo ERR_Update_Tribes_Transfers_Table
TRIBE_STATUS = "Update Tribes Transfers Table"

Dim tvtable As Recordset
Dim NEWTABLE As Recordset
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
    
If Update_Type = "Form" Then
   If New_Record = "YES" Then
   
   
   Else
   
   End If
ElseIf Update_Type = "Spreadsheet" Then

Else
   ' fuck it bombed.

End If

ERR_Update_Tribes_Transfers_Table_close:
   Exit Function

ERR_Update_Tribes_Transfers_Table:
If (Err = 3010) Then
   MsgBox "ERR_Update_Tribes_Transfers_Table Exists"
   Resume ERR_Update_Tribes_Transfers_Table_close
Else
   Dim errorstring As String
   errorstring = Err.Description
   Msg = "err = " & Err & " " & errorstring
    
   MsgBox (Msg)
   Resume Next
End If


End Function


Function UPDATE_TRIBES_GOODS_TABLES(CLAN As String, TRIBE As String, ITEM As String, MOVE_TYPE As String, MOVE_QUANTITY As Long)
On Error GoTo ERR_TABLES
TRIBE_STATUS = "Update Tribes Goods Tables"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"

VALID_GOODS:
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", ITEM
       
If VALIDGOODS.NoMatch Then
   Msg = "The Item requested - " & ITEM & " was not found in the Valid_Goods Table" & Chr(13) & Chr(10)
   Msg = Msg & "You will need to manually " & MOVE_TYPE & " " & MOVE_QUANTITY & " into the goods table" & Chr(13) & Chr(10)
   Msg = Msg & "for Clan " & CLAN & ", Tribe " & TRIBE & Chr(13) & Chr(10)
   Msg = Msg & "You will also need to update the Valid_Goods table with the Good " & ITEM
  MsgBox (Msg)
End If

War:
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", CLAN, TRIBE, VALIDGOODS![TABLE], ITEM
   If MOVE_TYPE = "ADD" Then
      If TRIBESGOODS.NoMatch Then
         TRIBESGOODS.AddNew
         TRIBESGOODS![CLAN] = CLAN
         TRIBESGOODS![TRIBE] = TRIBE
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
       
VALIDGOODS.Close
      
ERR_close:
   Exit Function

ERR_TABLES:
If (Err = 91) Or (Err = 3420) Then
   Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.MoveFirst
   
   Resume War
   
Else
  Msg = "Error # " & Err & " " & Error$
  Msg = Msg & "The Tribe is most likely missing the following good to produce the item"
  Msg = Msg & " requested.  You can probably ignore this."
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  Resume ERR_close

End If

End Function

Public Function VERIFY_QUANTITY(ITEM As String, MOVE_QUANTITY As Long)
On Error GoTo ERR_VERIFY
TRIBE_STATUS = "Verify Quantity"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

VALIDS:
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", ITEM
       
If VALIDGOODS.NoMatch Then
   MsgBox (ITEM)
   MsgBox (MOVE_QUANTITY)
End If

VAL_WAR:
   
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", TCLANNUMBER, GOODS_TRIBE, VALIDGOODS![TABLE], ITEM
   If TRIBESGOODS.NoMatch Then
      VERIFY_QUANTITY = "NO"
   ElseIf TRIBESGOODS![ITEM_NUMBER] >= MOVE_QUANTITY Then
      VERIFY_QUANTITY = "YES"
   Else
      TEMPITEM = TRIBESGOODS![ITEM_NUMBER]
      VERIFY_QUANTITY = "NO"
   End If
       

ERR_VAL_CLOSE:
   Exit Function


ERR_VERIFY:
If (Err = 91) Or (Err = 3420) Then
   If VALIDGOODS![TABLE] = "GENERAL" Or VALIDGOODS![TABLE] = "HUMANS" Then
      Resume VALIDS

   Else
      Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
      TRIBESGOODS.index = "PRIMARYKEY"
      TRIBESGOODS.MoveFirst
   
      Resume VAL_WAR
   
   End If
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  TQuantity(Index1) = 0
  TNUMOCCURS = 0
  Resume ERR_VAL_CLOSE
End If

End Function

Public Function UPDATE_TABLE(TABLE_NAME As String, Update_Type As String, New_Record As String)
On Error GoTo ERR_UPDATE_TABLE
TRIBE_STATUS = "Update Table"

Dim CITYTRADINGPOST As Recordset
Dim UPDATETABLE As Recordset
Dim QUANTITY As Double
Dim TRADINGSELLQUANTITY As Double
Dim TRADINGBUYQUANTITY As Double
Dim fullfield1 As String
Dim fullfield2 As String
Dim fullfield3 As String
Dim Current_Mineral As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set UPDATETABLE = TVDBGM.OpenRecordset(TABLE_NAME)
UPDATETABLE.index = "primarykey"
       
If TABLE_NAME = "TURNS_TRADING_POST_ACTIVITY" Then

   Set CITYTRADINGPOST = TVDBGM.OpenRecordset("TRADING_POST_GOODS")
   If Forms![TRADING POST]![CITY] = "GM SALE" _
   Or Forms![TRADING POST]![CITY] = "FAIR" Then
      CITYTRADINGPOST.index = "HEX_MAP_ID"
      CITYTRADINGPOST.MoveFirst
      CITYTRADINGPOST.Seek "=", Left(Forms![TRADING POST]![Current Hex], 2), Forms![TRADING POST]![ITEM]
   Else
      CITYTRADINGPOST.index = "TRIBESGOOD"
      CITYTRADINGPOST.MoveFirst
      CITYTRADINGPOST.Seek "=", Forms![TRADING POST]![CITY], Forms![TRADING POST]![ITEM]
   End If
   
   If Forms![TRADING POST]![QUANTITY] = "" Then
      Exit Function
   End If
       
   TRADINGSELLQUANTITY = CITYTRADINGPOST![SELL LIMIT]
   TRADINGBUYQUANTITY = CITYTRADINGPOST![BUY LIMIT]
   QUANTITY = Forms![TRADING POST]![QUANTITY]
   If Forms![TRADING POST]![TRADE_TYPE] = "SELL" Then
      If TRADINGBUYQUANTITY <= QUANTITY Then
         QUANTITY = TRADINGBUYQUANTITY
         CITYTRADINGPOST.Edit
         CITYTRADINGPOST![SELL_THIS_TURN] = "Y"
         CITYTRADINGPOST![SELL_TOTAL] = CITYTRADINGPOST![SELL_TOTAL] + QUANTITY
         CITYTRADINGPOST.UPDATE
      End If
   Else
      If TRADINGSELLQUANTITY < QUANTITY Then
         QUANTITY = CITYTRADINGPOST![SELL LIMIT]
         CITYTRADINGPOST.Edit
         CITYTRADINGPOST![BUY_THIS_TURN] = "Y"
         CITYTRADINGPOST![BUY_TOTAL] = CITYTRADINGPOST![BUY_TOTAL] + QUANTITY
         CITYTRADINGPOST.UPDATE
      End If
   End If
   UPDATETABLE.AddNew
   UPDATETABLE![CLAN] = Forms![TRADING POST]![CLANNUMBER]
   UPDATETABLE![TRIBE] = Forms![TRADING POST]![TRIBENUMBER]
   UPDATETABLE![TRADE_TYPE] = Forms![TRADING POST]![TRADE_TYPE]
   UPDATETABLE![ITEM] = Forms![TRADING POST]![ITEM]
   UPDATETABLE![QUANTITY] = QUANTITY
   UPDATETABLE![PRICE] = Forms![TRADING POST]![PRICE]
   UPDATETABLE![PROCESSED] = "N"
   UPDATETABLE.UPDATE
   'refresh the screen
   Forms![TRADING POST]![ITEM] = ""
   Forms![TRADING POST]![QUANTITY] = ""
   Forms![TRADING POST]![PRICE] = ""
   CITYTRADINGPOST.Close
ElseIf TABLE_NAME = "Process_TRIBES_TRANSFERS" Then
   If Update_Type = "Form" Then
      If IsNull(Forms![TRANSFER_GOODS]![QUANTITY]) Or Forms![TRANSFER_GOODS]![QUANTITY] = "" Then
         Exit Function
      End If
      If New_Record = "YES" Then
         UPDATETABLE.AddNew
         UPDATETABLE![From_Clan] = Forms![TRANSFER_GOODS]![FROM CLAN]
         UPDATETABLE![From_Tribe] = Forms![TRANSFER_GOODS]![FROM TRIBE]
         UPDATETABLE![To_Clan] = Forms![TRANSFER_GOODS]![TO CLAN]
         UPDATETABLE![To_Tribe] = Forms![TRANSFER_GOODS]![TO TRIBE]
         UPDATETABLE![ITEM] = Forms![TRANSFER_GOODS]![ITEM]
         UPDATETABLE![QUANTITY] = Forms![TRANSFER_GOODS]![QUANTITY]
         UPDATETABLE![ABSORBED] = "N"
         UPDATETABLE![PROCESSED] = "N"
         UPDATETABLE.UPDATE
         Forms![TRANSFER_GOODS]![ITEM] = ""
         Forms![TRANSFER_GOODS]![QUANTITY] = ""
      Else
   
      End If
   ElseIf Update_Type = "Spreadsheet" Then

   Else
      ' fuck it bombed.

   End If

ElseIf TABLE_NAME = "Process_Tribes_Activity" Then
   ' FIND CURRENT HIGH RECORD AND THEN ADD ONE TO IT
   TRIBE = Forms![TURNS ACTIVITIES]![TRIBENUMBER]

   UPDATETABLE.AddNew
   UPDATETABLE![TRIBE] = TRIBE
   UPDATETABLE![ACTIVITY] = Forms![TURNS ACTIVITIES]![ACTIVITY01]
   UPDATETABLE![ITEM] = Forms![TURNS ACTIVITIES]![item01]
   UPDATETABLE![DISTINCTION] = Forms![TURNS ACTIVITIES]![Distinction01]
   UPDATETABLE![PEOPLE] = Forms![TURNS ACTIVITIES]![ACTIVES]
   UPDATETABLE![Slaves] = Forms![TURNS ACTIVITIES]![Slaves]
   UPDATETABLE![SPECIALISTS] = Forms![TURNS ACTIVITIES]![SPECIALISTS]
   UPDATETABLE![JOINT] = Forms![TURNS ACTIVITIES]![JOINT_PROJECT]
   UPDATETABLE![OWNING_TRIBE] = Forms![TURNS ACTIVITIES]![Eng_Tribe]
   UPDATETABLE![Number_of_Seeking_Groups] = Forms![TURNS ACTIVITIES]![Number_Seeking_Groups]
   UPDATETABLE![Whale_Size] = Forms![TURNS ACTIVITIES]![Whale_Size]
   UPDATETABLE![MINING_DIRECTION] = Forms![TURNS ACTIVITIES]![Mine_Direction]
   UPDATETABLE![PROCESSED] = "N"
   UPDATETABLE![Building] = Forms![TURNS ACTIVITIES]![Building]
   UPDATETABLE.UPDATE
   UPDATETABLE.Close
   
   Call UPDATE_TABLE("Process_Tribes_Item_Allocation", "FORM", "YES")
   
ElseIf TABLE_NAME = "Process_Tribes_Item_Allocation" Then
   If Update_Type = "FORM" Then
      Set MYFORM = Forms![TURNS ACTIVITIES]
      TRIBE = MYFORM![TRIBENUMBER]
      GOODS_TRIBE = MYFORM![GOODS_TRIBE]
      count = 1
      fullfield1 = CStr("ITEM0" & count)
      fullfield2 = CStr("ITEM_0" & count)
      fullfield3 = CStr("USE_AMT_0" & count)
      Do
          If Not IsNull(MYFORM(fullfield2)) Then
              UPDATETABLE.AddNew
              UPDATETABLE![TRIBE] = MYFORM![TRIBENUMBER]
              UPDATETABLE![ACTIVITY] = MYFORM![ACTIVITY01]
              UPDATETABLE![ITEM] = MYFORM![item01]
              UPDATETABLE![ITEM_USED] = MYFORM(fullfield2)
              UPDATETABLE![QUANTITY] = MYFORM(fullfield3)
              UPDATETABLE![PROCESSED] = "N"
              UPDATETABLE.UPDATE
              IMPLEMENT = MYFORM(fullfield2)
              QUANTITY = MYFORM(fullfield3)
              Call Update_Implement_Usage(CLAN, GOODS_TRIBE, IMPLEMENT, QUANTITY)
          End If
          count = count + 1
          If count < 10 Then
              fullfield1 = CStr("ITEM0" & count)
              fullfield2 = CStr("ITEM_0" & count)
              fullfield3 = CStr("USE_AMT_0" & count)
          Else
              fullfield1 = CStr("ITEM" & count)
              fullfield2 = CStr("ITEM_" & count)
              fullfield3 = CStr("USE_AMT_" & count)
          End If
           If count > 14 Then
             Exit Do
          End If
      Loop
      
   Else
       ' build to handle spreadsheets or flat files
       
   End If
   
   ' GET FORM INFO - CLAN, TRIBE, HEX,TERRAIN, AVAILABLE_ACTIVES
   TTRIBENUMBER = Forms![TURNS ACTIVITIES]![TRIBENUMBER]
   GOODS_TRIBE = Forms![TURNS ACTIVITIES]![GOODS_TRIBE]
   CURRENT_HEX = Forms![TURNS ACTIVITIES]![Current Hex]
   Current_Mineral = Forms![TURNS ACTIVITIES]![Current_Mineral]
   TRIBES_TERRAIN = Forms![TURNS ACTIVITIES]![CURRENT TERRAIN]
   TActivesAvailable = Forms![TURNS ACTIVITIES]![Available_Actives]
   TSlavesAvailable = Forms![TURNS ACTIVITIES]![Available_Slaves]
   ' CLOSE FORM
   EXIT_FORMS ("TURNS ACTIVITIES")
   ' OPEN FORM
   OPEN_FORMS ("TURNS ACTIVITIES")
   ' REPOPULATE FIELDS
   Forms![TURNS ACTIVITIES]![TRIBENUMBER] = TTRIBENUMBER
   Forms![TURNS ACTIVITIES]![GOODS_TRIBE] = GOODS_TRIBE
   Forms![TURNS ACTIVITIES]![Current Hex] = CURRENT_HEX
   Forms![TURNS ACTIVITIES]![CURRENT TERRAIN] = TRIBES_TERRAIN
   Forms![TURNS ACTIVITIES]![Current_Mineral] = Current_Mineral
   Forms![TURNS ACTIVITIES]![Available_Actives] = TActivesAvailable
   Forms![TURNS ACTIVITIES]![Available_Slaves] = TSlavesAvailable
   Go_To_Field ("ACTIVITY01")
   
ElseIf TABLE_NAME = "Process_Tribes_Transfers_Table" Then
  ' need to build


ElseIf TABLE_NAME = "Process_Tribe_Movement" Then
  ' need to build


ElseIf TABLE_NAME = "Pacification_Table" Then
      Set MYFORM = Forms![PACIFICATION]
      CLAN = MYFORM![CLAN NAME]
      TRIBE = MYFORM![TRIBE NAME]
      UPDATETABLE.Seek "=", CLAN, TRIBE
      UPDATETABLE.Edit
      UPDATETABLE![primary_hex] = MYFORM![primary_hex]
      count = 1
      fullfield1 = CStr("GL1_" & count)
      Do
           If IsNull(UPDATETABLE(fullfield1)) Then
               UPDATETABLE(fullfield1) = 0
           Else
               UPDATETABLE(fullfield1) = MYFORM(fullfield1)
           End If
           count = count + 1
           fullfield1 = CStr("GL1_" & count)
           If count > 6 Then
             Exit Do
          End If
      Loop
      count = 1
      fullfield1 = CStr("GL2_" & count)
      Do
           If IsNull(UPDATETABLE(fullfield1)) Then
               UPDATETABLE(fullfield1) = 0
           Else
               UPDATETABLE(fullfield1) = MYFORM(fullfield1)
           End If
           count = count + 1
           fullfield1 = CStr("GL2_" & count)
           If count > 12 Then
             Exit Do
          End If
      Loop
      count = 1
      fullfield1 = CStr("GL3_" & count)
      Do
           If IsNull(UPDATETABLE(fullfield1)) Then
               UPDATETABLE(fullfield1) = 0
           Else
               UPDATETABLE(fullfield1) = MYFORM(fullfield1)
           End If
           count = count + 1
           fullfield1 = CStr("GL3_" & count)
           If count > 18 Then
             Exit Do
          End If
      Loop
      count = 1
      fullfield1 = CStr("GL4_" & count)
      Do
           If IsNull(UPDATETABLE(fullfield1)) Then
               UPDATETABLE(fullfield1) = 0
           Else
               UPDATETABLE(fullfield1) = MYFORM(fullfield1)
           End If
           count = count + 1
           fullfield1 = CStr("GL4_" & count)
           If count > 24 Then
             Exit Do
          End If
      Loop
      UPDATETABLE.UPDATE
   ' CLOSE FORM
   EXIT_FORMS ("pacification")
   ' OPEN FORM
   OPEN_FORMS ("pacification")
Else
  Msg = "Table not catered for = " & TABLE_NAME
  MsgBox (Msg)
End If
      
       
ERR_UPDATE_TABLE_close:
   
   UPDATETABLE.Close
   Exit Function

ERR_UPDATE_TABLE:
If (Err = 91) Or (Err = 3420) Then
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_UPDATE_TABLE_close

End If


End Function


Public Function Update_Implement_Usage(CLAN, TRIBE, IMPLEMENT, QUANTITY)
On Error GoTo ERR_POPULATE
TRIBE_STATUS = "Update Implement Usage"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set ImplementUsage = TVDB.OpenRecordset("IMPLEMENT_USAGE")
ImplementUsage.index = "PRIMARYKEY"
ImplementUsage.MoveFirst
ImplementUsage.Seek "=", CLAN, TRIBE, IMPLEMENT
If Not ImplementUsage.NoMatch Then
   ImplementUsage.Edit
   ImplementUsage![Number_Used] = ImplementUsage![Number_Used] + QUANTITY
   ImplementUsage.UPDATE
End If
ImplementUsage.Close

ERR_POP_CLOSE:
   Exit Function

ERR_POPULATE:
If (Err = 3021) Or (Err = 3022) Then
   Resume Next

Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_POP_CLOSE
End If


End Function

Public Function GET_TRIBES_GOOD_QUANTITY(CLAN As String, TRIBE As String, ITEM As String)
On Error GoTo ERR_TABLES
TRIBE_STATUS = "Get Tribes Good Quantity"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", ITEM
       
If VALIDGOODS.NoMatch Then
   MsgBox (ITEM)
   MsgBox ("WAS NOT FOUND IN THE VALID_GOODS TABLE")
   Exit Function
End If

Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", CLAN, TRIBE, VALIDGOODS![TABLE], ITEM
If TRIBESGOODS.NoMatch Then
   GET_TRIBES_GOOD_QUANTITY = 0
Else
   GET_TRIBES_GOOD_QUANTITY = TRIBESGOODS![ITEM_NUMBER]
End If
       
ERR_close:
   Exit Function

ERR_TABLES:
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  Resume ERR_close

End Function

Public Function UPDATE_TRIBES_SPECIALISTS(CLAN As String, TRIBE As String, ITEM As String, MOVE_TYPE As String, MOVE_QUANTITY As Long)
On Error GoTo ERR_TABLES
TRIBE_STATUS = "Update Tribes Specialists"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
VALID_GOODS:
Set VALIDGOODS = TVDBGM.OpenRecordset("TRIBES_SPECIALISTS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", CLAN, TRIBE, ITEM
       
If Not VALIDGOODS.NoMatch Then
   If MOVE_TYPE = "ADD SPECIALISTS" Then
      If VALIDGOODS.NoMatch Then
         VALIDGOODS.AddNew
         VALIDGOODS![CLAN] = CLAN
         VALIDGOODS![TRIBE] = TRIBE
         VALIDGOODS![ITEM] = ITEM
         VALIDGOODS![SPECIALISTS] = MOVE_QUANTITY
         VALIDGOODS![SPECIALISTS_USED] = 0
         VALIDGOODS.UPDATE
      Else
         VALIDGOODS.Edit
         VALIDGOODS![SPECIALISTS] = VALIDGOODS![SPECIALISTS] + MOVE_QUANTITY
         VALIDGOODS.UPDATE
      End If
   ElseIf MOVE_TYPE = "SUBTRACT SPECIALISTS" Then

      VALIDGOODS.Edit
      VALIDGOODS![SPECIALISTS] = VALIDGOODS![SPECIALISTS] - MOVE_QUANTITY
      VALIDGOODS.UPDATE
   ElseIf MOVE_TYPE = "SPECIALISTS_USED" Then

      VALIDGOODS.Edit
      VALIDGOODS![SPECIALISTS_USED] = VALIDGOODS![SPECIALISTS_USED] + MOVE_QUANTITY
      VALIDGOODS.UPDATE
   End If
End If

ERR_close:
   Exit Function

ERR_TABLES:
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  Resume ERR_close

End Function

Public Function GET_TRIBES_SPECIALISTS_QUANTITY(CLAN As String, TRIBE As String, ITEM As String)
On Error GoTo ERR_TABLES
TRIBE_STATUS = "Get Tribes Specialists Quantity"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_SPECIALISTS")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst
TRIBESGOODS.Seek "=", CLAN, TRIBE, ITEM
If TRIBESGOODS.NoMatch Then
   GET_TRIBES_SPECIALISTS_QUANTITY = 0
Else
   GET_TRIBES_SPECIALISTS_QUANTITY = TRIBESGOODS![SPECIALISTS]
End If
       
ERR_close:
   Exit Function

ERR_TABLES:
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  Resume ERR_close

End Function

Function UPDATE_TRIBES_SKILLS_TABLE(TRIBE As String, ITEM As String, MOVE_QUANTITY As Long)
On Error GoTo ERR_TABLES
TRIBE_STATUS = "Update Tribes Skills Table"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set VALIDGOODS = TVDBGM.OpenRecordset("SKILLS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
       
VALIDGOODS.AddNew
VALIDGOODS![TRIBE] = TRIBE
VALIDGOODS![Skill] = ITEM
VALIDGOODS![SKILL LEVEL] = MOVE_QUANTITY
VALIDGOODS![SUCCESSFUL] = "N"
VALIDGOODS![ATTEMPTED] = "N"
VALIDGOODS.UPDATE
VALIDGOODS.Close
       
ERR_close:
   Exit Function

ERR_TABLES:
  Msg = "Error # " & Err & " " & Error$
  Msg = Msg & "The Tribe is most likely missing the following good to produce the item"
  Msg = Msg & " requested.  You can probably ignore this."
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  Resume ERR_close


End Function

Public Function UPDATE_GOODS_TRIBE()
On Error GoTo ERR_TABLES
TRIBE_STATUS = "Update Goods Tribe"

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set MYFORM = SCREEN.ActiveForm
   
Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.Seek "=", MYFORM![CLANNUMBER], MYFORM![TRIBENUMBER]
   
TRIBESINFO.Edit
TRIBESINFO![GOODS TRIBE] = MYFORM![New_Goods_Tribe]
TRIBESINFO.UPDATE

       
ERR_close:
   Exit Function

ERR_TABLES:
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  Resume ERR_close

End Function

