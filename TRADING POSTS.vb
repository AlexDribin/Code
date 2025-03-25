Attribute VB_Name = "TRADING POSTS"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'*===============================================================================*'
 
Global TRADE_TYPE As String
Global OUTPUTLINE As String
Global ITEM_NOT_FOUND As String
Global Section_Ident As String
Global SEQUENCE_NUMBER As Long
Global Available_Quantity As Long
Global GOODS_AVAILABLE As String
Global OUTPUT_LENGTH As Long
Global STOP_RUN As String
Global AMOUNT_SOLD As Long
Global ITEM_SOLD As String
Global AMOUNT_PURCHASED As Long
Global ITEM_PURCHASED As String
Function MOBILE_TRADING()
' MODIFY DATABASE

Dim TribesCheck As Recordset
Dim POST_ITEMS As Recordset

Dim FILE As String
Dim MAP As String
Dim MONTHS_OPEN As Long
Dim ECONOMICS As Long
Dim DIPLOMACY As Long
Dim SEASON As Long
Dim PRICE As Long
Dim BASE_PRICE As Long
Dim BUYERS As Long
Dim SELLERS As Long
Dim SILVER As Long
Dim LIMIT As Long
Dim ITEM As String
Dim QUANTITY As Long
Dim Cost As Long
Dim POST_FOUND As String
Dim PERCENTAGE As Long
Dim TURN_NUMBER As Long
Dim SILVER_ON_HAND As Long
Dim DICE1 As Long
Dim DICE2 As Long

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("tribes_general_info")
TRIBESINFO.index = "primarykey"
TRIBESINFO.MoveFirst

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst
TURN_NUMBER = Left(globalinfo![CURRENT TURN], 2)

Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXMAPCONST.index = "primarykey"
HEXMAPCONST.MoveFirst

Set TRADING_POST_GOODS = TVDBGM.OpenRecordset("TRADING_POST_GOODS")
TRADING_POST_GOODS.index = "TRIBE"
TRADING_POST_GOODS.MoveFirst

Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "primarykey"
SKILLSTABLE.MoveFirst

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst

Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst

Set TribesCheck = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TribesCheck.index = "HEX"
TribesCheck.MoveFirst

Set POST_ITEMS = TVDBGM.OpenRecordset("TEMP_TRADING_POST")
POST_ITEMS.index = "PRIMARYKEY"

roll1 = DROLL(6, 1, 100, 5, DICE_TRIBE, 0, 0)

SEQUENCE_NUMBER = 1

Do While Not TRIBESINFO.EOF
   TRADE_TYPE = "SOLD"
   Section_Ident = "Trading Post Sold"
   OUTPUTLINE = "Trading Post Sold :"

   Do While Not POST_ITEMS.EOF
      POST_ITEMS.Delete
      POST_ITEMS.MoveFirst
   Loop

   CLANNUMBER = TRIBESINFO![CLAN]
   TRIBENUMBER = TRIBESINFO![TRIBE]
   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      GOODS_CLAN = TRIBESINFO![GOODS_CLAN]
      GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      GOODS_CLAN = TRIBESINFO![CLAN]
      GOODS_TRIBE = TRIBESINFO![TRIBE]
   End If
   MAP = TRIBESINFO![Current Hex]
   ECONOMICS = 0
   DIPLOMACY = 0

   If CLANNUMBER = GOODS_CLAN Or IsNull(GOODS_CLAN) Then
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", CLANNUMBER, GOODS_TRIBE, "SILVER"
   Else
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", GOODS_CLAN, GOODS_TRIBE, "SILVER"
   End If


   If TRIBESGOODS.NoMatch Then
      SILVER_ON_HAND = 0
   Else
      SILVER_ON_HAND = TRIBESGOODS![NUMBER]
   End If

   'Determine if the hex/clan has a Trading Post
   HEXMAPCONST.index = "FORTHKEY"
   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", MAP, CLANNUMBER, "MONTHS TP OPEN"
   
   If HEXMAPCONST.NoMatch Then
      POST_FOUND = "N"
      STOP_RUN = "Y"
   ElseIf HEXMAPCONST![1] > 0 Then
      POST_FOUND = "Y"
      MONTHS_OPEN = HEXMAPCONST![1]
      STOP_RUN = "Y"
   End If

   'Find the Economics & Diplomacy skill levels  (How do we cope with multiple groups?)
   'Identify who is in the hex of the same clan and determine highest skill.
   If POST_FOUND = "Y" Then
      TribesCheck.index = "HEX"
      TribesCheck.MoveFirst
      TribesCheck.Seek "=", MAP
      STOP_RUN = "N"
    
      Do Until STOP_RUN = "Y"
         If TribesCheck.NoMatch Then
            STOP_RUN = "Y"
         ElseIf TribesCheck.EOF Then
            STOP_RUN = "Y"
            Exit Do
         ElseIf TribesCheck![CLAN] = CLANNUMBER Then
            SKILLSTABLE.MoveFirst
            SKILLSTABLE.Seek "=", TribesCheck![TRIBE], "DIPLOMACY"
            If Not SKILLSTABLE.NoMatch Then
               If DIPLOMACY > 0 Then
                  If SKILLSTABLE![SKILL LEVEL] > DIPLOMACY Then
                     DIPLOMACY = SKILLSTABLE![SKILL LEVEL]
                  End If
               Else
                  DIPLOMACY = SKILLSTABLE![SKILL LEVEL]
               End If
            End If

            SKILLSTABLE.MoveFirst
            SKILLSTABLE.Seek "=", TribesCheck![TRIBE], "ECONOMICS"
            If Not SKILLSTABLE.NoMatch Then
               If ECONOMICS > 0 Then
                  If SKILLSTABLE![SKILL LEVEL] > ECONOMICS Then
                     ECONOMICS = SKILLSTABLE![SKILL LEVEL]
                  End If
               Else
                  ECONOMICS = SKILLSTABLE![SKILL LEVEL]
               End If
            End If
            TribesCheck.MoveNext
         
         ElseIf TribesCheck![Current Hex] = MAP Then
            TribesCheck.MoveNext
         Else
            STOP_RUN = "Y"
            Exit Do
         End If
      Loop
  
   End If
   
   'Identify the items for sale by the tribe
   If POST_FOUND = "Y" Then
      TRADING_POST_GOODS.MoveFirst
      TRADING_POST_GOODS.Seek "=", "TRADE"
   
      Do
         If TRADING_POST_GOODS.NoMatch Then
            POST_FOUND = "N"
            Exit Do
         Else
            Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
            VALIDGOODS.index = "primarykey"
            VALIDGOODS.MoveFirst
            VALIDGOODS.Seek "=", TRADING_POST_GOODS![GOOD]
            
            If VALIDGOODS.NoMatch Then
               Msg = "TRADING POST GOODS HAS THE FOLLOWING GOOD - " & TRADING_POST_GOODS![GOOD]
               MSG1 = "VALID GOODS DOES NOT HAVE THE GOOD"
               Response = MsgBox(Msg & MSG1, True)
            Else
               If TURN_NUMBER >= 1 And TURN_NUMBER <= 3 Then
                  SEASON = VALIDGOODS![SPRING]
               ElseIf TURN_NUMBER >= 4 And TURN_NUMBER <= 6 Then
                  SEASON = VALIDGOODS![SUMMER]
               ElseIf TURN_NUMBER >= 7 And TURN_NUMBER <= 9 Then
                  SEASON = VALIDGOODS![AUTUMN]
               ElseIf TURN_NUMBER >= 10 And TURN_NUMBER <= 12 Then
                  SEASON = VALIDGOODS![WINTER]
               End If
               
               BASE_PRICE = VALIDGOODS![BASE SELL PRICE]
               
               POST_ITEMS.AddNew
               POST_ITEMS![ITEM] = TRADING_POST_GOODS![GOOD]
               POST_ITEMS![Name] = VALIDGOODS![SHORTNAME]
               POST_ITEMS![PRICE] = TRADING_POST_GOODS![SELL PRICE]
               POST_ITEMS![LIMIT] = TRADING_POST_GOODS![SELL LIMIT]
               If TRADING_POST_GOODS![SELL PRICE] = 0 Then
                  PERCENTAGE = 0
               ElseIf TRADING_POST_GOODS![SELL LIMIT] = 0 Then
                  PERCENTAGE = 0
               ElseIf BASE_PRICE = 0 Then
                  PERCENTAGE = CLng(SEASON * 1)
               Else
                  PERCENTAGE = CLng(SEASON * (BASE_PRICE / TRADING_POST_GOODS![SELL PRICE]))
               End If
               POST_ITEMS![PERCENTAGE] = PERCENTAGE
               POST_ITEMS.UPDATE
            
            End If
            TRADING_POST_GOODS.MoveNext
         
         End If
         If TRADING_POST_GOODS.EOF Then
            Exit Do
         End If
         If Not TRADING_POST_GOODS![TRIBE] = "TRADE" Then
            Exit Do
         End If
      Loop

   End If

   'Determine the Amount of silver available to buy goods from the Trading Post

   If POST_FOUND = "Y" Then
      If DIPLOMACY = 0 Then
         DIPLOMACY = 1
      End If
      If ECONOMICS = 0 Then
         ECONOMICS = 1
      End If
      
      DICE1 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)

      BUYERS = CLng((((DIPLOMACY * 10) * (MONTHS_OPEN / 12)) / 3) * 2)
      SILVER = CLng(BUYERS * DICE1 * ECONOMICS)
   
   End If
   
   'Determine what is sold
   If POST_FOUND = "Y" Then
      POST_ITEMS.MoveFirst

      Do While Not POST_ITEMS.EOF
         DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
         DICE2 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)
         POST_ITEMS.Edit
         POST_ITEMS![diceroll] = DICE1
         POST_ITEMS.UPDATE
         If DICE1 <= POST_ITEMS![PERCENTAGE] Then
            AMOUNT_SOLD = POST_ITEMS![LIMIT]
            ITEM_SOLD = POST_ITEMS![ITEM]
            
            If (AMOUNT_SOLD * POST_ITEMS![PRICE]) > SILVER Then
               AMOUNT_SOLD = CLng(SILVER / POST_ITEMS![PRICE])
               If (AMOUNT_SOLD * POST_ITEMS![PRICE]) > SILVER Then
                  AMOUNT_SOLD = AMOUNT_SOLD - 1
               End If
               SILVER = SILVER - (AMOUNT_SOLD * POST_ITEMS![PRICE])
            Else
               SILVER = SILVER - (AMOUNT_SOLD * POST_ITEMS![PRICE])
            End If
           
            ITEM_NOT_FOUND = "NO"

            If AMOUNT_SOLD > 0 Then
               Call UPDATE_GOODS(ITEM_SOLD, "SUBTRACT", AMOUNT_SOLD)
               If ITEM_NOT_FOUND = "NO" Then
                  Call UPDATE_GOODS("SILVER", "ADD", (AMOUNT_SOLD * POST_ITEMS![PRICE]))
                  OUTPUTLINE = OUTPUTLINE & " " & AMOUNT_SOLD & " " & POST_ITEMS![Name] & ","
               Else
                  SILVER = SILVER + (AMOUNT_SOLD * POST_ITEMS![PRICE])
               End If
            End If

            POST_ITEMS.Edit
            POST_ITEMS![SILVER] = SILVER
            POST_ITEMS![BUYERS] = BUYERS
            POST_ITEMS![AMOUNT_SOLD] = AMOUNT_SOLD
            POST_ITEMS![FOUND] = ITEM_NOT_FOUND
            POST_ITEMS.UPDATE
            
            AMOUNT_SOLD = 0
            ITEM_SOLD = ""
            POST_ITEMS.MoveNext
         Else
            POST_ITEMS.MoveNext
         End If
         If SILVER = 0 Then
            Exit Do
         End If
         If POST_ITEMS.EOF Then
            POST_ITEMS.MoveFirst
            Exit Do
         End If
      Loop

'      DoCmd OpenReport "TRADING_POST_SALES"
   End If
   
   If Len(OUTPUTLINE) > 19 Then
      Call OUTPUT_TRADING
   ElseIf POST_FOUND = "Y" Then
      OUTPUTLINE = Section_Ident
      Call OUTPUT_TRADING
   End If

   SEQUENCE_NUMBER = 1
   
   TRADE_TYPE = "PURCHASED"
   Section_Ident = "Trading Post Buy"
   OUTPUTLINE = "Trading Post Purchased :"
   
   'Identify the items wanted to purchase by the tribe
   
   Do While Not POST_ITEMS.EOF
      POST_ITEMS.Delete
      POST_ITEMS.MoveFirst
   Loop

   If POST_FOUND = "Y" Then
      TRADING_POST_GOODS.MoveFirst
      TRADING_POST_GOODS.Seek "=", "TRADE"
   
      Do While TRADING_POST_GOODS![TRIBE] = "TRADE"
         If TRADING_POST_GOODS.NoMatch Then
            Exit Do
         Else
            Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
            VALIDGOODS.index = "primarykey"
            VALIDGOODS.MoveFirst
            VALIDGOODS.Seek "=", TRADING_POST_GOODS![GOOD]
            
            If VALIDGOODS.NoMatch Then
               Msg = "TRADING POST GOODS HAS THE FOLLOWING GOOD - " & TRADING_POST_GOODS![GOOD]
               MSG1 = "VALID GOODS DOES NOT HAVE THE GOOD"
               Response = MsgBox(Msg & MSG1, True)
            
            Else
               If TURN_NUMBER >= 1 And TURN_NUMBER <= 3 Then
                  SEASON = VALIDGOODS![SPRING]
               ElseIf TURN_NUMBER >= 4 And TURN_NUMBER <= 6 Then
                  SEASON = VALIDGOODS![SUMMER]
               ElseIf TURN_NUMBER >= 7 And TURN_NUMBER <= 9 Then
                  SEASON = VALIDGOODS![AUTUMN]
               ElseIf TURN_NUMBER >= 10 And TURN_NUMBER <= 12 Then
                  SEASON = VALIDGOODS![WINTER]
               End If
               
               BASE_PRICE = VALIDGOODS![BASE BUY PRICE]

               POST_ITEMS.AddNew
               POST_ITEMS![ITEM] = TRADING_POST_GOODS![GOOD]
               POST_ITEMS![Name] = VALIDGOODS![SHORTNAME]
               POST_ITEMS![PRICE] = TRADING_POST_GOODS![BUY PRICE]
               POST_ITEMS![LIMIT] = TRADING_POST_GOODS![BUY LIMIT]
               If TRADING_POST_GOODS![BUY PRICE] = 0 Then
                  PERCENTAGE = 0
               ElseIf TRADING_POST_GOODS![BUY LIMIT] = 0 Then
                  PERCENTAGE = 0
               ElseIf BASE_PRICE = 0 Then
                  PERCENTAGE = CLng(SEASON * 1)
               Else
                  PERCENTAGE = CLng(SEASON * (TRADING_POST_GOODS![BUY PRICE] / BASE_PRICE))
               End If
               POST_ITEMS![PERCENTAGE] = PERCENTAGE
               POST_ITEMS.UPDATE
            End If
            TRADING_POST_GOODS.MoveNext
         End If
         If TRADING_POST_GOODS.EOF Then
            Exit Do
         End If
      Loop
   
   End If

   'Determine the Amount of silver available to buy goods from customers
   If POST_FOUND = "Y" Then
      If DIPLOMACY = 0 Then
         DIPLOMACY = 1
      End If
      If ECONOMICS = 0 Then
         ECONOMICS = 1
      End If
      
      DICE1 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)

      SELLERS = CLng((((DIPLOMACY * 10) * (MONTHS_OPEN / 12)) / 3) * 1)
      SILVER = CLng(SELLERS * DICE1 * ECONOMICS)
      If SILVER < SILVER_ON_HAND Then
         SILVER = SILVER_ON_HAND
      End If
      
   End If

   'Determine what is purchased

   If POST_FOUND = "Y" Then
      POST_ITEMS.MoveFirst

      Do While Not POST_ITEMS.EOF
         DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
         DICE2 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)
         POST_ITEMS.Edit
         POST_ITEMS![diceroll] = DICE1
         POST_ITEMS.UPDATE
         If DICE1 <= POST_ITEMS![PERCENTAGE] Then
            AMOUNT_PURCHASED = POST_ITEMS![LIMIT]
            ITEM_PURCHASED = POST_ITEMS![ITEM]
            
            If (AMOUNT_PURCHASED * POST_ITEMS![PRICE]) > SILVER Then
               AMOUNT_PURCHASED = CLng(SILVER / POST_ITEMS![PRICE])
               If (AMOUNT_PURCHASED * POST_ITEMS![PRICE]) > SILVER Then
                  AMOUNT_PURCHASED = AMOUNT_PURCHASED - 1
               End If
               SILVER = SILVER - (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
            Else
               SILVER = SILVER - (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
            End If
             
            ITEM_NOT_FOUND = "NO"

            If AMOUNT_PURCHASED > 0 Then
               Call UPDATE_GOODS("SILVER", "SUBTRACT", (AMOUNT_PURCHASED * POST_ITEMS![PRICE]))
               If ITEM_NOT_FOUND = "NO" Then
                  Call UPDATE_GOODS(ITEM_PURCHASED, "ADD", AMOUNT_PURCHASED)
                  OUTPUTLINE = OUTPUTLINE & " " & AMOUNT_PURCHASED & " " & POST_ITEMS![Name] & ","
               Else
                  SILVER = SILVER + (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
               End If
            End If

            POST_ITEMS.Edit
            POST_ITEMS![SILVER] = SILVER
            POST_ITEMS![BUYERS] = SELLERS
            POST_ITEMS![AMOUNT_SOLD] = AMOUNT_PURCHASED
            POST_ITEMS![FOUND] = ITEM_NOT_FOUND
            POST_ITEMS.UPDATE
            
            AMOUNT_PURCHASED = 0
            ITEM_PURCHASED = ""
            POST_ITEMS.MoveNext
         Else
            POST_ITEMS.MoveNext
         End If
         If SILVER = 0 Then
            Exit Do
         End If
         If POST_ITEMS.EOF Then
            Exit Do
         End If
      Loop

'      DoCmd OpenReport "TRADING_POST_PURCHASES"
   End If
   
   If Len(OUTPUTLINE) > 24 Then
      Call OUTPUT_TRADING
   ElseIf POST_FOUND = "Y" Then
      OUTPUTLINE = Section_Ident
      Call OUTPUT_TRADING
   End If
   
   TRIBESINFO.MoveNext

   If TRIBESINFO.EOF Then
      Exit Do
   End If

Loop
       
DoCmd.Hourglass False


End Function

Sub OUTPUT_TRADING()

Set ActivitiesTable = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
ActivitiesTable.index = "primarykey"
ActivitiesTable.Seek "=", CLANNUMBER, TRIBENUMBER, Section_Ident, SEQUENCE_NUMBER

Do

  If ActivitiesTable.NoMatch Then
     ActivitiesTable.AddNew
     ActivitiesTable![CLAN] = CLANNUMBER
     ActivitiesTable![TRIBE] = TRIBENUMBER
     ActivitiesTable("Section") = Section_Ident
     ActivitiesTable![LINE NUMBER] = SEQUENCE_NUMBER
     ActivitiesTable![line detail] = OUTPUTLINE
     ActivitiesTable.UPDATE
     Exit Do
  Else
     ActivitiesTable.Edit
     ActivitiesTable![line detail] = ActivitiesTable![line detail] & OUTPUTLINE
     ActivitiesTable.UPDATE
     Exit Do
  End If

Loop

ActivitiesTable.Close

End Sub

Function Trading_Post()
' THIS MODULE IS USED FOR CITY TRADING - IT IS ACCESSED FROM THE CITY TRADING SCREEN

Dim TURNSTRADINGPOSTACTIVITY As Recordset
Dim QUANTITY As Double
Dim PRICE As Double
Dim Cost As Double
Dim OUTPUTLINESOLD As String
Dim OUTPUTLINEBUY As String

DoCmd.Hourglass True

Set MYFORM = Forms![TRADING POST]

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("tribes_general_info")
TRIBESINFO.index = "primarykey"

Set TRIBESGOODS = TVDBGM.OpenRecordset("TRIBES_GOODS")
TRIBESGOODS.index = "primarykey"
TRIBESGOODS.MoveFirst

Set TURNSTRADINGPOSTACTIVITY = TVDBGM.OpenRecordset("Turns_TRADING_POST_ACTIVITY")
TURNSTRADINGPOSTACTIVITY.index = "PRIMARYKEY"
TURNSTRADINGPOSTACTIVITY.MoveFirst

TRIBENUMBER = "NEW"
TRADE_TYPE = "NEW"
SEQUENCE_NUMBER = 1

Do Until TURNSTRADINGPOSTACTIVITY.EOF
If TURNSTRADINGPOSTACTIVITY![PROCESSED] = "N" Then
   ' NEEDS TO RESET THIS EACH NEW TRIBE
   If Not TURNSTRADINGPOSTACTIVITY![TRIBE] = TRIBENUMBER _
   Or Not TRADE_TYPE = TURNSTRADINGPOSTACTIVITY![TRADE_TYPE] Then
      TRADE_TYPE = TURNSTRADINGPOSTACTIVITY![TRADE_TYPE]
      If TRADE_TYPE = "SELL" Then
         Section_Ident = "Trading Post Sold"
         OUTPUTLINESOLD = "City Purchased :"
      Else
         Section_Ident = "Trading Post Buy"
         OUTPUTLINEBUY = "City Sold :"
      End If
   End If
   
   CLANNUMBER = TURNSTRADINGPOSTACTIVITY![CLAN]
   TRIBENUMBER = TURNSTRADINGPOSTACTIVITY![TRIBE]
   ITEM = TURNSTRADINGPOSTACTIVITY![ITEM]
   TRADE_TYPE = TURNSTRADINGPOSTACTIVITY![TRADE_TYPE]
   QUANTITY = TURNSTRADINGPOSTACTIVITY![QUANTITY]
   PRICE = TURNSTRADINGPOSTACTIVITY![PRICE]
   Cost = QUANTITY * PRICE
         
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", CLANNUMBER, TRIBENUMBER
   TRIBESINFO.Edit
    
   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      TRIBENUMBER = TRIBESINFO![GOODS TRIBE]
      TRIBESINFO.MoveFirst
      TRIBESINFO.Seek "=", CLANNUMBER, TRIBENUMBER
      TRIBESINFO.Edit
   End If
     
   Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
   VALIDGOODS.index = "primarykey"
   VALIDGOODS.MoveFirst
   VALIDGOODS.Seek "=", ITEM
       
   If VALIDGOODS![TABLE] = "WAR" Or VALIDGOODS![TABLE] = "FINISHED" _
   Or VALIDGOODS![TABLE] = "SHIP" Or VALIDGOODS![TABLE] = "RAW" _
   Or VALIDGOODS![TABLE] = "MINERAL" Or VALIDGOODS![TABLE] = "ANIMAL" Then
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", CLANNUMBER, TRIBENUMBER, VALIDGOODS![TABLE], ITEM
            
      If TRIBESGOODS.NoMatch Then
         TRIBESGOODS.AddNew
         TRIBESGOODS![CLAN] = CLANNUMBER
         TRIBESGOODS![TRIBE] = TRIBENUMBER
         TRIBESGOODS![ITEM_TYPE] = VALIDGOODS![TABLE]
         TRIBESGOODS![ITEM] = ITEM
         TRIBESGOODS![ITEM_NUMBER] = 0
         TRIBESGOODS.UPDATE
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", CLANNUMBER, TRIBENUMBER, VALIDGOODS![TABLE], ITEM
      End If
   ElseIf VALIDGOODS![TABLE] = "GENERAL" Then
         ' NO ACTION REQUIRED
   ElseIf VALIDGOODS![TABLE] = "HUMANS" Then
         ' NO ACTION REQUIRED
   Else
      Msg = "ITEM NOT IN VALID GOODS TABLE " & ITEM
      MsgBox (Msg)
   End If
       
   If TRADE_TYPE = "SELL" Then
      If VALIDGOODS![TABLE] = "HUMANS" Then
         If ITEM = "SLAVE" Then
            If QUANTITY > TRIBESINFO![SLAVE] Then
               QUANTITY = TRIBESINFO![SLAVE]
               Cost = QUANTITY * PRICE
            End If
            TRIBESINFO.Edit
            TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] - QUANTITY
            TRIBESINFO.UPDATE
            TRIBESGOODS.MoveFirst
            TRIBESGOODS.Seek "=", CLANNUMBER, TRIBENUMBER, "MINERAL", "SILVER"
            If TRIBESGOODS.NoMatch Then
               TRIBESGOODS.AddNew
               TRIBESGOODS![CLAN] = CLANNUMBER
               TRIBESGOODS![TRIBE] = TRIBENUMBER
               TRIBESGOODS![ITEM_TYPE] = "MINERAL"
               TRIBESGOODS![ITEM] = "SILVER"
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + Cost
               TRIBESGOODS.UPDATE
            Else
               TRIBESGOODS.Edit
               TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + Cost
               TRIBESGOODS.UPDATE
            End If
         End If
      Else
         If QUANTITY > TRIBESGOODS![ITEM_NUMBER] Then
            QUANTITY = TRIBESGOODS![ITEM_NUMBER]
            Cost = QUANTITY * PRICE
         End If
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - QUANTITY
         TRIBESGOODS.UPDATE
         TRIBESGOODS.MoveFirst
         TRIBESGOODS.Seek "=", CLANNUMBER, TRIBENUMBER, "MINERAL", "SILVER"
         If TRIBESGOODS.NoMatch Then
            TRIBESGOODS.AddNew
            TRIBESGOODS![CLAN] = CLANNUMBER
            TRIBESGOODS![TRIBE] = TRIBENUMBER
            TRIBESGOODS![ITEM_TYPE] = "MINERAL"
            TRIBESGOODS![ITEM] = "SILVER"
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + Cost
            TRIBESGOODS.UPDATE
         ElseIf IsNull(TRIBESGOODS![ITEM_NUMBER]) Then
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = Cost
            TRIBESGOODS.UPDATE
         Else
            TRIBESGOODS.Edit
            TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + Cost
            TRIBESGOODS.UPDATE
         End If
      End If
   Else
      TRIBESGOODS.MoveFirst
      TRIBESGOODS.Seek "=", CLANNUMBER, TRIBENUMBER, "MINERAL", "SILVER"
      TRIBESGOODS.Edit
      If Cost > TRIBESGOODS![ITEM_NUMBER] Then
         QUANTITY = Fix(TRIBESGOODS![ITEM_NUMBER] / PRICE) ' removes the fraction
         Cost = QUANTITY * PRICE
      End If
      TRIBESGOODS.Edit
      TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - Cost
      TRIBESGOODS.UPDATE
      If TRIBESGOODS![ITEM_NUMBER] < 0 Then
          TRIBESGOODS.Edit
          TRIBESGOODS![ITEM_NUMBER] = 0
          TRIBESGOODS.UPDATE
      End If
      If VALIDGOODS![TABLE] = "HUMANS" Then
         If ITEM = "SLAVE" Then
            TRIBESINFO.Edit
            TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] + QUANTITY
            TRIBESINFO.UPDATE
         End If
      Else
         TRIBESGOODS.Seek "=", CLANNUMBER, TRIBENUMBER, VALIDGOODS![TABLE], ITEM
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + QUANTITY
         TRIBESGOODS.UPDATE
      End If
   End If

   TURNSTRADINGPOSTACTIVITY.Edit
   TURNSTRADINGPOSTACTIVITY![PROCESSED] = "Y"
   TURNSTRADINGPOSTACTIVITY.UPDATE
   
   If TRADE_TYPE = "SELL" Then
      Section_Ident = "Trading Post Sold"
      OUTPUTLINESOLD = OUTPUTLINESOLD & " " & QUANTITY & " " & ITEM & ","
   Else
      Section_Ident = "Trading Post Buy"
      OUTPUTLINEBUY = OUTPUTLINEBUY & " " & QUANTITY & " " & ITEM & ","
   End If
End If

   TURNSTRADINGPOSTACTIVITY.MoveNext
   If TURNSTRADINGPOSTACTIVITY.EOF Then
      If Len(OUTPUTLINESOLD) > 0 Then
         Section_Ident = "Trading Post Sold"
         OUTPUTLINE = OUTPUTLINESOLD
         Call OUTPUT_TRADING
         OUTPUTLINESOLD = ""
      End If
      If Len(OUTPUTLINEBUY) > 0 Then
         Section_Ident = "Trading Post Buy"
         OUTPUTLINE = OUTPUTLINEBUY
         Call OUTPUT_TRADING
         OUTPUTLINEBUY = ""
      End If
      Exit Do
   End If
   If Not TRIBENUMBER = TURNSTRADINGPOSTACTIVITY![TRIBE] Then
      If Len(OUTPUTLINESOLD) > 0 Then
         Section_Ident = "Trading Post Sold"
         OUTPUTLINE = OUTPUTLINESOLD
         Call OUTPUT_TRADING
         OUTPUTLINESOLD = ""
      End If
      If Len(OUTPUTLINEBUY) > 0 Then
         Section_Ident = "Trading Post Buy"
         OUTPUTLINE = OUTPUTLINEBUY
         Call OUTPUT_TRADING
         OUTPUTLINEBUY = ""
      End If
   End If

Loop

DoCmd.Hourglass False

EXIT_FORMS ("TRADING POST")
OPEN_FORMS ("TRADING POST")

End Function

Sub UPDATE_GOODS(ITEM, MOVE_TYPE, MOVE_QUANTITY)
On Error GoTo ERR_TABLES
Dim VALTABLE As String

VALID_GOODS:
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", ITEM
       
VALTABLE = VALIDGOODS![TABLE]

If VALIDGOODS.NoMatch Then
   MsgBox (ITEM)
   MsgBox (MOVE_QUANTITY)
End If
If VALTABLE = "ANIMAL" Or VALTABLE = "WAR" Or VALTABLE = "RAW" Or VALTABLE = "FINISHED" Or VALTABLE = "SHIP" Or VALTABLE = "MINERAL" Then
GOODS:
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", GOODS_CLAN, GOODS_TRIBE, VALIDGOODS![TABLE], ITEM
   If MOVE_TYPE = "ADD" Then
      If TRIBESGOODS.NoMatch Then
         TRIBESGOODS.AddNew
         TRIBESGOODS![CLAN] = CLANNUMBER
         TRIBESGOODS![TRIBE] = GOODS_TRIBE
         TRIBESGOODS![ITEM_TYPE] = VALIDGOODS![TABLE]
         TRIBESGOODS![ITEM] = ITEM
         TRIBESGOODS![ITEM_NUMBER] = MOVE_QUANTITY
         TRIBESGOODS.UPDATE
      ElseIf IsNull(TRIBESGOODS![ITEM_NUMBER]) Then
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = MOVE_QUANTITY
         TRIBESGOODS.UPDATE
      Else
         TRIBESGOODS.Edit
         TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] + MOVE_QUANTITY
         TRIBESGOODS.UPDATE
      End If
   ElseIf Not TRIBESGOODS.NoMatch Then
       TRIBESGOODS.Edit
       TRIBESGOODS![ITEM_NUMBER] = TRIBESGOODS![ITEM_NUMBER] - MOVE_QUANTITY
       TRIBESGOODS.UPDATE
       If TRIBESGOODS![ITEM_NUMBER] <= 0 Then
          TRIBESGOODS.Delete
       End If
    End If
  
ElseIf VALIDGOODS![TABLE] = "HUMANS" Then
GENERAL:
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", CLANNUMBER, GOODS_TRIBE
   If MOVE_TYPE = "ADD" Then
      If ITEM = "SLAVE" Then
         TRIBESINFO.Edit
         TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] + MOVE_QUANTITY
         TRIBESINFO.UPDATE
      End If
   Else
      TRIBESINFO.Edit
      TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] - MOVE_QUANTITY
      TRIBESINFO.UPDATE
   End If
     
End If
       
ERR_close:
   Exit Sub

ERR_TABLES:
If Err = 3021 Then
   ITEM_NOT_FOUND = "YES"
   Resume ERR_close

ElseIf (Err = 91) Or (Err = 3420) Then
If VALIDGOODS![TABLE] = "ANIMAL" Or "WAR" Or "RAW" Or "FINISHED" Or "SHIP" Or "MINERAL" Then
   Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.MoveFirst
   
   Resume GOODS
   
ElseIf VALIDGOODS![TABLE] = "HUMANS" Then
   Set TRIBESINFO = TVDBGM.OpenRecordset("tribes_general_info")
   TRIBESINFO.index = "PRIMARYKEY"
   TRIBESINFO.MoveFirst

    Resume GENERAL
     
Else
    Resume VALID_GOODS

End If
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  MsgBox (MOVE_TYPE)
  MsgBox ("Update Goods")
  Resume ERR_close

End If

End Sub

Sub VERIFY_TRADE_QUANTITY(ITEM, MOVE_QUANTITY)
On Error GoTo ERR_VERIFY
Dim VALTABLE As String

VALIDS:
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst
VALIDGOODS.Seek "=", ITEM
       
If VALIDGOODS.NoMatch Then
   MsgBox (ITEM)
   MsgBox (MOVE_QUANTITY)
End If

VALTABLE = VALIDGOODS![TABLE]

If VALTABLE = "ANIMAL" Or VALTABLE = "WAR" Or VALTABLE = "RAW" Or VALTABLE = "FINISHED" _
Or VALTABLE = "SHIP" Or VALTABLE = "MINERAL" Then
VAL_GOODS:
   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", CLANNUMBER, GOODS_TRIBE, VALIDGOODS![TABLE], ITEM
   If Not TRIBESGOODS.NoMatch Then
      If TRIBESGOODS![ITEM_NUMBER] < MOVE_QUANTITY Then
         If TRIBESGOODS![ITEM_NUMBER] > 0 Then
          Available_Quantity = TRIBESGOODS![ITEM_NUMBER]
            GOODS_AVAILABLE = "Y"
         Else
            GOODS_AVAILABLE = "N"
         End If
      ElseIf TRIBESGOODS![ITEM_NUMBER] > 0 Then
        GOODS_AVAILABLE = "Y"
      Else
         GOODS_AVAILABLE = "N"
      End If
   Else
      Available_Quantity = 0
      GOODS_AVAILABLE = "N"
   End If
End If
       

ERR_VAL_CLOSE:
   Exit Sub


ERR_VERIFY:
If (Err = 91) Or (Err = 3420) Then
 If VALIDGOODS![TABLE] = "ANIMAL" Or VALIDGOODS![TABLE] = "WAR" _
 Or VALIDGOODS![TABLE] = "RAW" Or VALIDGOODS![TABLE] = "FINISHED" _
 Or VALIDGOODS![TABLE] = "SHIP" Or VALIDGOODS![TABLE] = "MINERAL" Then
   Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
   TRIBESGOODS.index = "PRIMARYKEY"
   TRIBESGOODS.MoveFirst
   
   Resume VAL_GOODS
   
 Else
    
    Resume VALIDS

 End If
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox (ITEM)
  MsgBox (MOVE_QUANTITY)
  MsgBox ("Verify Trade Quantity")
  TQuantity(Index1) = 0
  TNUMOCCURS = 0
  Resume ERR_VAL_CLOSE
End If

End Sub

Function Village_Trading()
On Error GoTo ERR_VERIFY
' MODIFY DATABASE

Dim TribesCheck As Recordset
Dim POST_ITEMS As Recordset

Dim FILE As String
Dim MAP As String
Dim MONTHS_OPEN As Long
Dim ECONOMICS As Long
Dim DIPLOMACY As Long
Dim SEASON As Long
Dim PRICE As Long
Dim BASE_PRICE As Long
Dim BUYERS As Long
Dim SELLERS As Long
Dim SILVER As Long
Dim LIMIT As Long
Dim ITEM As String
Dim QUANTITY As Long
Dim Cost As Long
Dim POST_FOUND As String
Dim PERCENTAGE As Long
Dim TURN_NUMBER As Long
Dim SILVER_ON_HAND As Long
Dim DICE1 As Long
Dim DICE2 As Long
Dim ITEM_FOR_SALE_FOUND As String
Dim ITEM_FOR_PURCHASE_FOUND As String

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILE = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILE, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("tribes_general_info")
TRIBESINFO.index = "primarykey"
TRIBESINFO.MoveFirst

Set globalinfo = TVDBGM.OpenRecordset("Global")
globalinfo.index = "PRIMARYKEY"
globalinfo.MoveFirst
TURN_NUMBER = Left(globalinfo![CURRENT TURN], 2)

Set HEXMAPCONST = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXMAPCONST.index = "primarykey"
HEXMAPCONST.MoveFirst

Set TRADING_POST_GOODS = TVDBGM.OpenRecordset("TRADING_POST_GOODS")
TRADING_POST_GOODS.index = "TRIBE"
TRADING_POST_GOODS.MoveFirst

Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
SKILLSTABLE.index = "primarykey"
SKILLSTABLE.MoveFirst

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"
VALIDGOODS.MoveFirst

Set TRIBESGOODS = TVDBGM.OpenRecordset("Tribes_Goods")
TRIBESGOODS.index = "PRIMARYKEY"
TRIBESGOODS.MoveFirst
  
Set TribesCheck = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TribesCheck.index = "HEX"
TribesCheck.MoveFirst

Set POST_ITEMS = TVDBGM.OpenRecordset("TEMP_TRADING_POST")
POST_ITEMS.index = "PRIMARYKEY"

roll1 = DROLL(6, 1, 100, 5, DICE_TRIBE, 0, 0)

SEQUENCE_NUMBER = 1

Do While Not TRIBESINFO.EOF
   ITEM_FOR_SALE_FOUND = "NO"
   ITEM_FOR_PURCHASE_FOUND = "NO"
   TRADE_TYPE = "SOLD"
   Section_Ident = "Trading Post Sold"
   OUTPUTLINE = "^BTrading Post Sold :^B"

   Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TEMP_TRADING_POST;")
   qdfCurrent.Execute

   CLANNUMBER = TRIBESINFO![CLAN]
   TRIBENUMBER = TRIBESINFO![TRIBE]
   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      GOODS_TRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      GOODS_TRIBE = TRIBESINFO![TRIBE]
   End If
   MAP = TRIBESINFO![Current Hex]
   ECONOMICS = 0
   DIPLOMACY = 0

   TRIBESGOODS.MoveFirst
   TRIBESGOODS.Seek "=", CLANNUMBER, GOODS_TRIBE, "MINERAL", "SILVER"

   If TRIBESGOODS.NoMatch Then
      SILVER_ON_HAND = 0
   ElseIf IsNull(TRIBESGOODS![ITEM_NUMBER]) Then
      SILVER_ON_HAND = 0
   Else
      SILVER_ON_HAND = TRIBESGOODS![ITEM_NUMBER] - 100
   End If

   'Determine if the hex/clan has a Trading Post
   HEXMAPCONST.index = "FORTHKEY"
   HEXMAPCONST.MoveFirst
   HEXMAPCONST.Seek "=", MAP, CLANNUMBER, "MONTHS TP OPEN"
   
   If HEXMAPCONST.NoMatch Then
      POST_FOUND = "N"
      STOP_RUN = "Y"
   ElseIf HEXMAPCONST![1] > 0 Then
      POST_FOUND = "Y"
      MONTHS_OPEN = HEXMAPCONST![1]
      STOP_RUN = "Y"
   End If

   'Find the Economics & Diplomacy skill levels  (How do we cope with multiple groups?)
   'Identify who is in the hex of the same clan and determine highest skill.
   If POST_FOUND = "Y" Then
      TribesCheck.index = "HEX"
      TribesCheck.MoveFirst
      TribesCheck.Seek "=", MAP
      STOP_RUN = "N"
    
      Do Until STOP_RUN = "Y"
         If TribesCheck.NoMatch Then
            STOP_RUN = "Y"
         ElseIf TribesCheck.EOF Then
            STOP_RUN = "Y"
            Exit Do
         ElseIf TribesCheck![CLAN] = CLANNUMBER Then
            SKILLSTABLE.MoveFirst
            SKILLSTABLE.Seek "=", TribesCheck![TRIBE], "DIPLOMACY"
            If Not SKILLSTABLE.NoMatch Then
               If DIPLOMACY > 0 Then
                  If SKILLSTABLE![SKILL LEVEL] > DIPLOMACY Then
                     DIPLOMACY = SKILLSTABLE![SKILL LEVEL]
                  End If
               Else
                  DIPLOMACY = SKILLSTABLE![SKILL LEVEL]
               End If
            End If

            SKILLSTABLE.MoveFirst
            SKILLSTABLE.Seek "=", TribesCheck![TRIBE], "ECONOMICS"
            If Not SKILLSTABLE.NoMatch Then
               If ECONOMICS > 0 Then
                  If SKILLSTABLE![SKILL LEVEL] > ECONOMICS Then
                     ECONOMICS = SKILLSTABLE![SKILL LEVEL]
                  End If
               Else
                  ECONOMICS = SKILLSTABLE![SKILL LEVEL]
               End If
            End If
            TribesCheck.MoveNext
         
         ElseIf TribesCheck![Current Hex] = MAP Then
            TribesCheck.MoveNext
         Else
            STOP_RUN = "Y"
            Exit Do
         End If
      Loop
  
   End If
   
   'Identify the items for sale by the tribe
   If POST_FOUND = "Y" Then
      TRADING_POST_GOODS.index = "TRIBE"
      TRADING_POST_GOODS.MoveFirst
      TRADING_POST_GOODS.Seek "=", TRIBENUMBER
   
      Do
         If TRADING_POST_GOODS.NoMatch Then
            POST_FOUND = "N"
            Exit Do
         Else
            Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
            VALIDGOODS.index = "primarykey"
            VALIDGOODS.MoveFirst
            VALIDGOODS.Seek "=", TRADING_POST_GOODS![GOOD]
            
            If VALIDGOODS.NoMatch Then
               Msg = "TRADING POST GOODS HAS THE FOLLOWING GOOD - " & TRADING_POST_GOODS![GOOD]
               MSG1 = "VALID GOODS DOES NOT HAVE THE GOOD"
               Response = MsgBox(Msg & MSG1, True)
            Else
               If TURN_NUMBER >= 1 And TURN_NUMBER <= 3 Then
                  SEASON = VALIDGOODS![SPRING]
               ElseIf TURN_NUMBER >= 4 And TURN_NUMBER <= 6 Then
                  SEASON = VALIDGOODS![SUMMER]
               ElseIf TURN_NUMBER >= 7 And TURN_NUMBER <= 9 Then
                  SEASON = VALIDGOODS![AUTUMN]
               ElseIf TURN_NUMBER >= 10 And TURN_NUMBER <= 12 Then
                  SEASON = VALIDGOODS![WINTER]
               End If
               
               BASE_PRICE = VALIDGOODS![BASE SELL PRICE]
               
               If TRADING_POST_GOODS![SELL PRICE] > 0 Then
                   POST_ITEMS.AddNew
                   POST_ITEMS![ITEM] = TRADING_POST_GOODS![GOOD]
                   POST_ITEMS![Name] = VALIDGOODS![SHORTNAME]
                   POST_ITEMS![PRICE] = TRADING_POST_GOODS![SELL PRICE]
                   POST_ITEMS![LIMIT] = TRADING_POST_GOODS![SELL LIMIT]
                   POST_ITEMS![PERCENTAGE] = 0
                   POST_ITEMS.UPDATE
                   ITEM_FOR_SALE_FOUND = "YES"
               End If
            
            End If
            TRADING_POST_GOODS.MoveNext
         
         End If
         If TRADING_POST_GOODS.EOF Then
            Exit Do
         End If
         If Not TRADING_POST_GOODS![TRIBE] = TRIBENUMBER Then
            Exit Do
         End If
      Loop

   End If

   'Determine the Amount of silver available to buy goods from the Trading Post

   If POST_FOUND = "Y" Then
      If DIPLOMACY = 0 Then
         DIPLOMACY = 1
      End If
      If ECONOMICS = 0 Then
         ECONOMICS = 1
      End If
      
      DICE1 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)

      BUYERS = CLng((((DIPLOMACY * 10) * (MONTHS_OPEN / 12)) / 3) * 2)
      SILVER = CLng((BUYERS * DICE1) * ECONOMICS)
   End If
   
   'Determine what is sold
   If POST_FOUND = "Y" Then
       POST_ITEMS.MoveFirst

      Do While Not POST_ITEMS.EOF
         DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
         DICE2 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)
         POST_ITEMS.Edit
         POST_ITEMS![diceroll] = DICE1
         POST_ITEMS.UPDATE
         TRADING_POST_GOODS.index = "primarykey"
         TRADING_POST_GOODS.MoveFirst
         TRADING_POST_GOODS.Seek "=", "SALE", Left(MAP, 2), "GM SALE", POST_ITEMS![ITEM]
         PERCENTAGE = CLng(SEASON * (TRADING_POST_GOODS![SELL PRICE] / POST_ITEMS![PRICE]))
     
         If DICE1 <= PERCENTAGE Then
            ' LIMIT IS TO BE THE LIMIT FROM THE GM SALE
            ' need to check against price.  If Post_Price is greater thean Tribesgoods Price, reduce the limit by percentage price is greater than tribesgoods
            If TRADING_POST_GOODS![BUY LIMIT] < POST_ITEMS![LIMIT] Then
                AMOUNT_SOLD = TRADING_POST_GOODS![BUY LIMIT]
            Else
                 AMOUNT_SOLD = POST_ITEMS![LIMIT]
            End If
            If TRADING_POST_GOODS![BUY PRICE] < POST_ITEMS![PRICE] Then
                AMOUNT_SOLD = CLng(AMOUNT_SOLD * (TRADING_POST_GOODS![BUY PRICE] / POST_ITEMS![PRICE]))
            End If

            ITEM_SOLD = POST_ITEMS![ITEM]
            Available_Quantity = 0
            GOODS_AVAILABLE = "N"
            
            TRADING_POST_GOODS.index = "tribe"
            
            Call VERIFY_TRADE_QUANTITY(ITEM_SOLD, AMOUNT_SOLD)
            
            If GOODS_AVAILABLE = "Y" Then
               If Available_Quantity > 0 Then
                  AMOUNT_SOLD = Available_Quantity
               End If
            Else
               AMOUNT_SOLD = 0
            End If

            If (AMOUNT_SOLD * POST_ITEMS![PRICE]) > SILVER Then
               AMOUNT_SOLD = CLng(SILVER / POST_ITEMS![PRICE])
               If (AMOUNT_SOLD * POST_ITEMS![PRICE]) > SILVER Then
                  AMOUNT_SOLD = AMOUNT_SOLD - 1
               End If
               SILVER = SILVER - (AMOUNT_SOLD * POST_ITEMS![PRICE])
            Else
               SILVER = SILVER - (AMOUNT_SOLD * POST_ITEMS![PRICE])
            End If
           
            ITEM_NOT_FOUND = "NO"

            If AMOUNT_SOLD > 0 Then
               Call UPDATE_GOODS(ITEM_SOLD, "SUBTRACT", AMOUNT_SOLD)
               If ITEM_NOT_FOUND = "NO" Then
                  Call UPDATE_GOODS("SILVER", "ADD", (AMOUNT_SOLD * POST_ITEMS![PRICE]))
                  OUTPUTLINE = OUTPUTLINE & " " & AMOUNT_SOLD & " " & POST_ITEMS![Name]
                  OUTPUTLINE = OUTPUTLINE & " sold for " & (AMOUNT_SOLD * POST_ITEMS![PRICE]) & ","
               Else
                  SILVER = SILVER + (AMOUNT_SOLD * POST_ITEMS![PRICE])
               End If
            End If

            POST_ITEMS.Edit
            POST_ITEMS![SILVER] = SILVER
            POST_ITEMS![BUYERS] = BUYERS
            POST_ITEMS![AMOUNT_SOLD] = AMOUNT_SOLD
            POST_ITEMS![FOUND] = ITEM_NOT_FOUND
            POST_ITEMS.UPDATE
            
            AMOUNT_SOLD = 0
            ITEM_SOLD = ""
            POST_ITEMS.MoveNext
         Else
            POST_ITEMS.MoveNext
         End If
         If SILVER = 0 Then
            Exit Do
         End If
         If POST_ITEMS.EOF Then
            Exit Do
         End If
      Loop

      TRADING_POST_GOODS.index = "TRIBE"
      TRADING_POST_GOODS.MoveFirst

'      DoCmd OpenReport "TRADING_POST_SALES"
   End If
   
   If POST_FOUND = "Y" Then
      If Len(OUTPUTLINE) > 19 Then
         Call OUTPUT_TRADING
       ElseIf POST_FOUND = "Y" Then
         OUTPUTLINE = Section_Ident
         Call OUTPUT_TRADING
      End If
   End If
   
   SEQUENCE_NUMBER = 1
   
   TRADE_TYPE = "PURCHASED"
   Section_Ident = "Trading Post Buy"
   OUTPUTLINE = "^BTrading Post Purchased :^B"
   
   'Identify the items wanted to purchase by the tribe
   
   Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TEMP_TRADING_POST;")
   qdfCurrent.Execute

   If POST_FOUND = "Y" Then
      TRADING_POST_GOODS.MoveFirst
      TRADING_POST_GOODS.Seek "=", TRIBENUMBER
   
      Do While TRADING_POST_GOODS![TRIBE] = TRIBENUMBER
         If TRADING_POST_GOODS.NoMatch Then
            Exit Do
         Else
            Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
            VALIDGOODS.index = "primarykey"
            VALIDGOODS.MoveFirst
            VALIDGOODS.Seek "=", TRADING_POST_GOODS![GOOD]
            
            If VALIDGOODS.NoMatch Then
               Msg = "TRADING POST GOODS HAS THE FOLLOWING GOOD - " & TRADING_POST_GOODS![GOOD]
               MSG1 = "VALID GOODS DOES NOT HAVE THE GOOD"
               Response = MsgBox(Msg & MSG1, True)
            
            Else
               If TURN_NUMBER >= 1 And TURN_NUMBER <= 3 Then
                  SEASON = VALIDGOODS![SPRING]
               ElseIf TURN_NUMBER >= 4 And TURN_NUMBER <= 6 Then
                  SEASON = VALIDGOODS![SUMMER]
               ElseIf TURN_NUMBER >= 7 And TURN_NUMBER <= 9 Then
                  SEASON = VALIDGOODS![AUTUMN]
               ElseIf TURN_NUMBER >= 10 And TURN_NUMBER <= 12 Then
                  SEASON = VALIDGOODS![WINTER]
               End If
               
               BASE_PRICE = VALIDGOODS![BASE BUY PRICE]

               If TRADING_POST_GOODS![BUY PRICE] > 0 Then
                   POST_ITEMS.AddNew
                   POST_ITEMS![ITEM] = TRADING_POST_GOODS![GOOD]
                   POST_ITEMS![Name] = VALIDGOODS![SHORTNAME]
                   POST_ITEMS![PRICE] = TRADING_POST_GOODS![BUY PRICE]
                   POST_ITEMS![LIMIT] = TRADING_POST_GOODS![BUY LIMIT]
                   POST_ITEMS![PERCENTAGE] = 0
                   POST_ITEMS.UPDATE
                   ITEM_FOR_PURCHASE_FOUND = "YES"
               End If
            End If
            TRADING_POST_GOODS.MoveNext
         End If
         If TRADING_POST_GOODS.EOF Then
            Exit Do
         End If
      Loop
   
   End If

   'Determine the Amount of silver available to buy goods from customers
   If POST_FOUND = "Y" Then
      If DIPLOMACY = 0 Then
         DIPLOMACY = 1
      End If
      If ECONOMICS = 0 Then
         ECONOMICS = 1
      End If
      
      DICE1 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)

      SELLERS = CLng((((DIPLOMACY * 10) * (MONTHS_OPEN / 12)) / 3) * 1)
      SILVER = CLng((SELLERS * DICE1) * ECONOMICS)
      If SILVER < SILVER_ON_HAND Then
         SILVER = SILVER_ON_HAND
      End If
      
   End If

   'Determine what is purchased

   If POST_FOUND = "Y" Then
       If ITEM_FOR_PURCHASE_FOUND = "YES" Then
          POST_ITEMS.MoveFirst
      End If
      
      Do While Not POST_ITEMS.EOF
         DICE1 = DROLL(6, 1, 100, 0, DICE_TRIBE, 0, 0)
         DICE2 = DROLL(6, 1, 10, 0, DICE_TRIBE, 0, 0)
         POST_ITEMS.Edit
         POST_ITEMS![diceroll] = DICE1
         POST_ITEMS.UPDATE
         TRADING_POST_GOODS.index = "primarykey"
         TRADING_POST_GOODS.MoveFirst
         TRADING_POST_GOODS.Seek "=", "SALE", Left(MAP, 2), "GM SALE", POST_ITEMS![ITEM]
         PERCENTAGE = CLng(SEASON * (POST_ITEMS![PRICE] / TRADING_POST_GOODS![SELL PRICE]))
    
         If DICE1 <= PERCENTAGE Then
            ' LIMIT IS TO BE THE LIMIT FROM THE GM SALE
            ' need to check against price.  If Post_Price is greater thean Tribesgoods Price, reduce the limit by percentage price is greater than tribesgoods
            If TRADING_POST_GOODS![SELL LIMIT] < POST_ITEMS![LIMIT] Then
                AMOUNT_PURCHASED = TRADING_POST_GOODS![SELL LIMIT]
            Else
                 AMOUNT_PURCHASED = POST_ITEMS![LIMIT]
            End If
            If TRADING_POST_GOODS![SELL PRICE] < POST_ITEMS![PRICE] Then
                AMOUNT_SOLD = CLng(AMOUNT_SOLD * (POST_ITEMS![PRICE] / TRADING_POST_GOODS![SELL PRICE]))
            End If

            ITEM_PURCHASED = POST_ITEMS![ITEM]
            
            TRADING_POST_GOODS.index = "TRIBE"
            
            Available_Quantity = 0
            GOODS_AVAILABLE = "N"

            Call VERIFY_TRADE_QUANTITY(ITEM_PURCHASED, AMOUNT_PURCHASED)
            
            If GOODS_AVAILABLE = "Y" Then
               If Available_Quantity > 0 Then
                  AMOUNT_PURCHASED = Available_Quantity
               End If
            End If

            If (AMOUNT_PURCHASED * POST_ITEMS![PRICE]) > SILVER Then
               AMOUNT_PURCHASED = CLng(SILVER / POST_ITEMS![PRICE])
               If (AMOUNT_PURCHASED * POST_ITEMS![PRICE]) > SILVER Then
                  AMOUNT_PURCHASED = AMOUNT_PURCHASED - 1
               End If
               SILVER = SILVER - (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
            Else
               SILVER = SILVER - (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
            End If
             
            ITEM_NOT_FOUND = "NO"

            If AMOUNT_PURCHASED > 0 Then
               Call UPDATE_GOODS("SILVER", "SUBTRACT", (AMOUNT_PURCHASED * POST_ITEMS![PRICE]))
               If ITEM_NOT_FOUND = "NO" Then
                  Call UPDATE_GOODS(ITEM_PURCHASED, "ADD", AMOUNT_PURCHASED)
                  OUTPUTLINE = OUTPUTLINE & " " & AMOUNT_PURCHASED & " " & POST_ITEMS![Name]
                  OUTPUTLINE = OUTPUTLINE & " purchased for " & (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
                  OUTPUTLINE = OUTPUTLINE & ","
               Else
                  SILVER = SILVER + (AMOUNT_PURCHASED * POST_ITEMS![PRICE])
               End If
            End If

            POST_ITEMS.Edit
            POST_ITEMS![SILVER] = SILVER
            POST_ITEMS![BUYERS] = SELLERS
            POST_ITEMS![AMOUNT_SOLD] = AMOUNT_PURCHASED
            POST_ITEMS![FOUND] = ITEM_NOT_FOUND
            POST_ITEMS.UPDATE
            
            AMOUNT_PURCHASED = 0
            ITEM_PURCHASED = ""
            POST_ITEMS.MoveNext
         Else
            POST_ITEMS.MoveNext
         End If
         If SILVER = 0 Then
            Exit Do
         End If
         If POST_ITEMS.EOF Then
            Exit Do
         End If
      Loop
      POST_ITEMS.MoveFirst
      
'      DoCmd OpenReport "TRADING_POST_PURCHASES"
   End If
   
   If POST_FOUND = "Y" Then
      If Len(OUTPUTLINE) > 24 Then
         Call OUTPUT_TRADING
      ElseIf POST_FOUND = "Y" Then
         OUTPUTLINE = Section_Ident
         Call OUTPUT_TRADING
      End If
   End If
   
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", CLANNUMBER, TRIBENUMBER
   TRIBESINFO.MoveNext
   If TRIBESINFO.EOF Then
      Exit Do
   End If

Loop
       
DoCmd.Hourglass False

ERR_VAL_CLOSE:
   Exit Function


ERR_VERIFY:
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  MsgBox ("MAP = " & Left(MAP, 2))
  MsgBox (POST_ITEMS![ITEM])
  TQuantity(Index1) = 0
  TNUMOCCURS = 0
  Resume ERR_VAL_CLOSE
  'Resume Next



End Function


Public Function UPDATE_TRADING_POST_GOODS_TABLE()
' THIS MODULE IS CALLED BY THE GLOBAL_INFO PROCEDURE.
' ITS PURPOSE IS TO MODIFY THE PRICES AND LIMITS.

Dim TRADINGPOSTGOODS As Recordset

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRADINGPOSTGOODS = TVDBGM.OpenRecordset("TRADING_POST_GOODS")
TRADINGPOSTGOODS.MoveFirst

Do Until TRADINGPOSTGOODS.EOF
   TRADINGPOSTGOODS.Edit

'shit










   TRADINGPOSTGOODS.MoveNext
   If TRADINGPOSTGOODS.EOF Then
      Exit Do
   End If
   
Loop

DoCmd.Hourglass False

End Function
