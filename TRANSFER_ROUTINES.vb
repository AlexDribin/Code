Attribute VB_Name = "TRANSFER_ROUTINES"
Option Compare Database   'Use database order for string comparisons
Option Explicit
Global From_Clan As String
Global To_Clan As String
Global FROMCLAN As String
Global TOCLAN As String


'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 21/01/96  *  Allow for transfer all goods                                   **'
'** 25/02/96  *  Fix transfer people                                            **'
'*===============================================================================*'
 

Function ABSORB_GROUP()
' MODIFY DATABASE
Dim TribeGoods As Recordset
Dim TRIBESBOOKS As Recordset
Dim FROMCLAN As String
Dim FROMTRIBE As String
Dim TOCLAN As String
Dim TOTRIBE As String
Dim ITEM As String
Dim QUANTITY As Long
Dim count As Long
Dim TRIBE_ITEM_TYPE As String
Dim TRIBE_ITEM As String
Dim PI_TOKENS As Long
Dim WARRIORS As Long
Dim ACTIVES As Long
Dim INACTIVES As Long
Dim Slaves As Long
Dim TURN01 As Long
Dim TURN02 As Long
Dim TURN03 As Long
Dim TURN04 As Long
Dim TURN05 As Long
Dim TURN06 As Long
Dim TURN07 As Long
Dim TURN08 As Long
Dim TURN09 As Long
Dim TURN10 As Long
Dim TURN11 As Long
Dim TURN12 As Long
Dim Continue_Absorbtion As String
Dim Absorb_Skills As String
Dim Absorb_Research As String
Dim Skill As String
Dim Skill_Level As Long
Dim TOPIC As String
 Dim BOOK As String
Dim Number_Of_Books As Long

Set MYFORM = Forms![TRANSFER_GOODS]

DoCmd.Hourglass True

' are you sure you want to absorb.

Continue_Absorbtion = InputBox("Do you really want to absorb group?", "ABSORBING", "N")

If Continue_Absorbtion = "N" Then
   DoCmd.Hourglass False
   Exit Function
End If

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TribeGoods = TVDBGM.OpenRecordset("tribes_Goods")
TribeGoods.index = "CLANTRIBE"
TribeGoods.MoveFirst
TribeGoods.Seek "=", MYFORM![FROM CLAN], MYFORM![FROM TRIBE]

Do Until TribeGoods.NoMatch
   TribeGoods.Edit
   TRIBE_ITEM_TYPE = TribeGoods![ITEM_TYPE]
   TRIBE_ITEM = TribeGoods![ITEM]
   QUANTITY = TribeGoods![ITEM_NUMBER]
   TribeGoods.Delete
   TribeGoods.index = "PRIMARYKEY"
   TribeGoods.MoveFirst
   TribeGoods.Seek "=", MYFORM![TO CLAN], MYFORM![TO TRIBE], TRIBE_ITEM_TYPE, TRIBE_ITEM
   If TribeGoods.NoMatch Then
      TribeGoods.AddNew
      TribeGoods![CLAN] = MYFORM![TO CLAN]
      TribeGoods![TRIBE] = MYFORM![TO TRIBE]
      TribeGoods![ITEM_TYPE] = TRIBE_ITEM_TYPE
      TribeGoods![ITEM] = TRIBE_ITEM
      TribeGoods![ITEM_NUMBER] = QUANTITY
      TribeGoods.UPDATE
   Else
      TribeGoods.Edit
      TribeGoods![ITEM_NUMBER] = TribeGoods![ITEM_NUMBER] + QUANTITY
      TribeGoods.UPDATE
   End If
   TribeGoods.index = "CLANTRIBE"
   TribeGoods.MoveFirst
   TribeGoods.Seek "=", MYFORM![FROM CLAN], MYFORM![FROM TRIBE]

Loop

TribeGoods.Close

Set TRIBESINFO = TVDBGM.OpenRecordset("tribes_general_info")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", MYFORM![FROM CLAN], MYFORM![FROM TRIBE]
TRIBESINFO.Edit
PI_TOKENS = TRIBESINFO![PI TOKENS]
WARRIORS = TRIBESINFO![WARRIORS]
ACTIVES = TRIBESINFO![ACTIVES]
INACTIVES = TRIBESINFO![INACTIVES]
Slaves = TRIBESINFO![SLAVE]
TRIBESINFO![WARRIORS] = 0
TRIBESINFO![ACTIVES] = 0
TRIBESINFO![INACTIVES] = 0
TRIBESINFO![SLAVE] = 0
TRIBESINFO![ABSORBED] = "Y"
TRIBESINFO.UPDATE
TRIBESINFO.MoveFirst
TRIBESINFO.Seek "=", MYFORM![TO CLAN], MYFORM![TO TRIBE]
TRIBESINFO.Edit
TRIBESINFO![PI TOKENS] = TRIBESINFO![PI TOKENS] + PI_TOKENS
TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] + WARRIORS
TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] + ACTIVES
TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] + INACTIVES
TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] + Slaves
TRIBESINFO.UPDATE
TRIBESINFO.Close

' MOVE POPULATION INCREASE
Set PopTable = TVDBGM.OpenRecordset("POPULATION_INCREASE")
PopTable.index = "PRIMARYKEY"
PopTable.MoveFirst
PopTable.Seek "=", MYFORM![FROM CLAN], MYFORM![FROM TRIBE]
If Not PopTable.NoMatch Then
PopTable.Edit
TURN01 = PopTable![TURN01]
TURN02 = PopTable![TURN02]
TURN03 = PopTable![TURN03]
TURN04 = PopTable![TURN04]
TURN05 = PopTable![TURN05]
TURN06 = PopTable![TURN06]
TURN07 = PopTable![TURN07]
TURN08 = PopTable![TURN08]
TURN09 = PopTable![TURN09]
TURN10 = PopTable![TURN10]
TURN11 = PopTable![TURN11]
TURN12 = PopTable![TURN12]
PopTable.Delete
PopTable.MoveFirst
PopTable.Seek "=", MYFORM![TO CLAN], MYFORM![TO TRIBE]
If PopTable.NoMatch Then
   PopTable.AddNew
   PopTable![CLAN] = MYFORM![TO CLAN]
   PopTable![TRIBE] = MYFORM![TO TRIBE]
   PopTable![TURN01] = TURN01
   PopTable![TURN02] = TURN02
   PopTable![TURN03] = TURN03
   PopTable![TURN04] = TURN04
   PopTable![TURN05] = TURN05
   PopTable![TURN06] = TURN06
   PopTable![TURN07] = TURN07
   PopTable![TURN08] = TURN08
   PopTable![TURN09] = TURN09
   PopTable![TURN10] = TURN10
   PopTable![TURN11] = TURN11
   PopTable![TURN12] = TURN12
   PopTable.UPDATE
Else
   PopTable.Edit
   PopTable![TURN01] = PopTable![TURN01] + TURN01
   PopTable![TURN02] = PopTable![TURN02] + TURN02
   PopTable![TURN03] = PopTable![TURN03] + TURN03
   PopTable![TURN04] = PopTable![TURN04] + TURN04
   PopTable![TURN05] = PopTable![TURN05] + TURN05
   PopTable![TURN06] = PopTable![TURN06] + TURN06
   PopTable![TURN07] = PopTable![TURN07] + TURN07
   PopTable![TURN08] = PopTable![TURN08] + TURN08
   PopTable![TURN09] = PopTable![TURN09] + TURN09
   PopTable![TURN10] = PopTable![TURN10] + TURN10
   PopTable![TURN11] = PopTable![TURN11] + TURN11
   PopTable![TURN12] = PopTable![TURN12] + TURN12
   PopTable.UPDATE
End If

PopTable.Close
End If

' TRANSFER BOOKS
Set TRIBESBOOKS = TVDBGM.OpenRecordset("TRIBES_BOOKS")
TRIBESBOOKS.index = "PRIMARYKEY"
TRIBESBOOKS.MoveFirst

Do
     If TRIBESBOOKS!TRIBE = MYFORM![FROM TRIBE] Then
        BOOK = TRIBESBOOKS![BOOK]
        Number_Of_Books = TRIBESBOOKS![NUMBER]
        TRIBESBOOKS.Delete
        TRIBESBOOKS.MoveFirst
        TRIBESBOOKS.Seek "=", MYFORM![TO CLAN], MYFORM![TO TRIBE], BOOK
        If TRIBESBOOKS.NoMatch Then
            TRIBESBOOKS.AddNew
            TRIBESBOOKS!CLAN = MYFORM![TO CLAN]
            TRIBESBOOKS!TRIBE = MYFORM![TO TRIBE]
            TRIBESBOOKS!BOOK = BOOK
            TRIBESBOOKS![NUMBER] = Number_Of_Books
            TRIBESBOOKS.UPDATE
        Else
            TRIBESBOOKS.Edit
            TRIBESBOOKS![NUMBER] = TRIBESBOOKS![NUMBER] + Number_Of_Books
            TRIBESBOOKS.UPDATE
        End If
        TRIBESBOOKS.MoveFirst
     End If
     TRIBESBOOKS.MoveNext
     If TRIBESBOOKS.NoMatch Then
         Exit Do
     End If
     If TRIBESBOOKS.EOF Then
         Exit Do
     End If
Loop

Absorb_Skills = "NO"
Absorb_Research = "NO"

Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTAB.index = "PRIMARYKEY"
COMPRESTAB.MoveFirst
COMPRESTAB.Seek "=", MYFORM![TO TRIBE], "Ability to absorb tribe without skill loss"

If COMPRESTAB.NoMatch Then
    COMPRESTAB.Seek "=", MYFORM![FROM TRIBE], "Ability to absorb tribe without skill loss"
    If COMPRESTAB.NoMatch Then
        Absorb_Skills = "NO"
    Else
        Absorb_Skills = "YES"
    End If
Else
    Absorb_Skills = "YES"
End If

COMPRESTAB.Seek "=", MYFORM![TO TRIBE], "Ability to absorb tribe without research loss"

If COMPRESTAB.NoMatch Then
    COMPRESTAB.Seek "=", MYFORM![FROM TRIBE], "Ability to absorb tribe without research loss"
    If COMPRESTAB.NoMatch Then
        Absorb_Research = "NO"
    Else
        Absorb_Research = "YES"
    End If
Else
    Absorb_Research = "YES"
End If

' IF GOT Ability to absorb tribe without skill loss
' READ SKILL TABLE
   Set SKILLSTABLE = TVDBGM.OpenRecordset("SKILLS")
   SKILLSTABLE.index = "TRIBE"
   SKILLSTABLE.MoveFirst
   SKILLSTABLE.Seek "=", MYFORM![FROM TRIBE]

   If Not SKILLSTABLE.NoMatch Then
   If Absorb_Skills = "YES" Then
       Do
            Skill = SKILLSTABLE![Skill]
            Skill_Level = SKILLSTABLE![SKILL LEVEL]
            SKILLSTABLE.Delete
            SKILLSTABLE.index = "PRIMARYKEY"
            SKILLSTABLE.MoveFirst
            SKILLSTABLE.Seek "=", MYFORM![TO TRIBE], Skill
            If SKILLSTABLE.NoMatch Then
               SKILLSTABLE.AddNew
               SKILLSTABLE!TRIBE = MYFORM![TO TRIBE]
               SKILLSTABLE!Skill = Skill
               SKILLSTABLE![SKILL LEVEL] = Skill_Level
               SKILLSTABLE![SUCCESSFUL] = "N"
               SKILLSTABLE![ATTEMPTED] = "N"
               SKILLSTABLE.UPDATE
            ElseIf SKILLSTABLE![SKILL LEVEL] < Skill_Level Then
               SKILLSTABLE.Edit
               SKILLSTABLE![SKILL LEVEL] = Skill_Level
               SKILLSTABLE.UPDATE
            End If
           SKILLSTABLE.index = "TRIBE"
           SKILLSTABLE.MoveFirst
           SKILLSTABLE.Seek "=", MYFORM![FROM TRIBE]
           If SKILLSTABLE.NoMatch Then
              Exit Do
           End If
       Loop
   End If
   End If
   
' IF GOT TOPIC Ability to absorb tribe without research loss
' READ COMPLETED RESEARCH TABLE
   Set COMPRESTAB = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
   COMPRESTAB.index = "TRIBE"
   COMPRESTAB.MoveFirst
   COMPRESTAB.Seek "=", MYFORM![FROM TRIBE]

   If Not COMPRESTAB.NoMatch Then
   If Absorb_Research = "YES" Then
       Do
            TOPIC = COMPRESTAB![TOPIC]
            COMPRESTAB.Delete
            COMPRESTAB.index = "PRIMARYKEY"
            COMPRESTAB.MoveFirst
            COMPRESTAB.Seek "=", MYFORM![TO TRIBE], TOPIC
            If COMPRESTAB.NoMatch Then
               COMPRESTAB.AddNew
               COMPRESTAB!TRIBE = MYFORM![TO TRIBE]
               COMPRESTAB!TOPIC = TOPIC
               COMPRESTAB!COMPLETED_THIS_TURN = "N"
               COMPRESTAB.UPDATE
            End If
           COMPRESTAB.index = "TRIBE"
           COMPRESTAB.MoveFirst
           COMPRESTAB.Seek "=", MYFORM![FROM TRIBE]
           If COMPRESTAB.NoMatch Then
              Exit Do
           End If
       Loop
   End If
   End If

DoCmd.Hourglass False
DoCmd.Close acForm, "TRANSFER_GOODS"
DoCmd.OpenForm "TRANSFER_GOODS"

End Function

Function fix_negatives()
Dim TRIBESTABLE As Recordset
Dim TribeGoods As Recordset

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TribeGoods = TVDBGM.OpenRecordset("tribes_GOODS")
TribeGoods.MoveFirst

Do Until TribeGoods.EOF
   TribeGoods.Edit
   If TribeGoods![NUMBER] < 0 Then
      TribeGoods![NUMBER] = 0
      TribeGoods.UPDATE
   End If
   TribeGoods.MoveNext

Loop

TribeGoods.Close

DoCmd.Hourglass False

End Function

Function TRANSFER_ALL_GOODS()
' MODIFY DATABASE
Dim TribeGoods As Recordset
Dim FROMCLAN As String
Dim FROMTRIBE As String
Dim TOCLAN As String
Dim TOTRIBE As String
Dim ITEM As String
Dim QUANTITY As Long
Dim count As Long
Dim TRIBE_ITEM_TYPE As String
Dim TRIBE_ITEM As String

Set MYFORM = Forms![TRANSFER_GOODS]

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TribeGoods = TVDBGM.OpenRecordset("tribes_GOODS")
TribeGoods.index = "CLANTRIBE"
TribeGoods.MoveFirst
TribeGoods.Seek "=", MYFORM![FROM CLAN], MYFORM![FROM TRIBE]

Do Until TribeGoods.NoMatch
   TribeGoods.Edit
   TRIBE_ITEM_TYPE = TribeGoods![ITEM_TYPE]
   TRIBE_ITEM = TribeGoods![ITEM]
   QUANTITY = TribeGoods![ITEM_NUMBER]
   TribeGoods.Delete
   TribeGoods.index = "PRIMARYKEY"
   TribeGoods.MoveFirst
   TribeGoods.Seek "=", MYFORM![TO CLAN], MYFORM![TO TRIBE], TRIBE_ITEM_TYPE, TRIBE_ITEM
   If TribeGoods.NoMatch Then
      TribeGoods.AddNew
      TribeGoods![CLAN] = MYFORM![TO CLAN]
      TribeGoods![TRIBE] = MYFORM![TO TRIBE]
      TribeGoods![ITEM_TYPE] = TRIBE_ITEM_TYPE
      TribeGoods![ITEM] = TRIBE_ITEM
      TribeGoods![ITEM_NUMBER] = QUANTITY
      TribeGoods.UPDATE
   Else
      TribeGoods.Edit
      TribeGoods![ITEM_NUMBER] = TribeGoods![ITEM_NUMBER] + QUANTITY
      TribeGoods.UPDATE
   End If
   TribeGoods.index = "CLANTRIBE"
   TribeGoods.MoveFirst
   TribeGoods.Seek "=", MYFORM![FROM CLAN], MYFORM![FROM TRIBE]

Loop

TribeGoods.Close

DoCmd.Hourglass False

DoCmd.Close acForm, "TRANSFER_GOODS"
DoCmd.OpenForm "TRANSFER_GOODS"

End Function

Function TRIBE_TRANSFERS()
On Error GoTo ERR_TRIBE_TRANSFERS

' MODIFY DATABASE
Dim TRIBESGOOD As Recordset, OUTTAB As Recordset
Dim TRIBESTRANSFERS As Recordset
Dim INFILE As String
Dim FROMTRIBE As String
Dim TOTRIBE As String
Dim ITEM As String
Dim OUTPUTLINE As String
Dim QUANTITY As Long
Dim count As Long
Dim INGOODSTRIBE As String
Dim TOGOODSTRIBE As String
Dim LINENUMBER As Long
Dim POSITION As Long

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
' OPEN TRANSFER TABLE
Set TRIBESTRANSFERS = TVDBGM.OpenRecordset("PROCESS_Tribes_TRANSFERS")
TRIBESTRANSFERS.index = "PRIMARYKEY"
TRIBESTRANSFERS.MoveFirst

' LOOP THROUGH IT DOING EACH TRANSFER
Do
If TRIBESTRANSFERS![PROCESSED] = "Y" Then
   GoTo NEXT_LOOP
Else
   FROMTRIBE = TRIBESTRANSFERS![From_Tribe]
   FROM_CLANNUMBER = "0" & Mid(FROMTRIBE, 2, 3)
   TCLANNUMBER = "0" & Mid(FROMTRIBE, 2, 3)
   TOTRIBE = TRIBESTRANSFERS![To_Tribe]
   To_Clan = "0" & Mid(TOTRIBE, 2, 3)
   ITEM = TRIBESTRANSFERS![ITEM]
   QUANTITY = TRIBESTRANSFERS![QUANTITY]
  
   Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
   TRIBESINFO.index = "PRIMARYKEY"
   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", TCLANNUMBER, FROMTRIBE

   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      INGOODSTRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      INGOODSTRIBE = FROMTRIBE
   End If

   TRIBESINFO.MoveFirst
   TRIBESINFO.Seek "=", To_Clan, TOTRIBE

   If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
      TOGOODSTRIBE = TRIBESINFO![GOODS TRIBE]
   Else
      TOGOODSTRIBE = TOTRIBE
   End If
   
   LINENUMBER = 1

   'SETUP LINE DETAIL
   Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
   OUTTAB.index = "primarykey"
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", LINENUMBER

   If OUTTAB.NoMatch Then
      OUTTAB.AddNew
      OUTTAB![CLAN] = TCLANNUMBER
      OUTTAB![TRIBE] = FROMTRIBE
      OUTTAB![Section] = "TRANSFERS OUT"
      OUTTAB![LINE NUMBER] = LINENUMBER
      OUTTAB![line detail] = "Transfer goods to " & TOTRIBE & ": "
      OUTTAB.UPDATE
      OUTTAB.MoveFirst
      OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
      OUTTAB.Edit
   Else
      OUTTAB.MoveFirst
      OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", (LINENUMBER + 1)
      If OUTTAB.NoMatch Then
         OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
      Else
         OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", (LINENUMBER + 2)
         If OUTTAB.NoMatch Then
            LINENUMBER = LINENUMBER + 1
            OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
         Else
            LINENUMBER = LINENUMBER + 2
         End If
      End If
      OUTTAB.Edit
      ' IF LINE CONTAINS TOTRIBE THEN JUST COMMA ELSE
      POSITION = InStr(OUTTAB![line detail], TOTRIBE)
      If POSITION > 0 Then
         OUTTAB![line detail] = OUTTAB![line detail] & " "
      Else
         OUTTAB![line detail] = OUTTAB![line detail] & "To " & TOTRIBE & ": "
      End If
      OUTTAB.UPDATE
      OUTTAB.MoveFirst
      OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
      OUTTAB.Edit
   End If

   OUTPUTLINE = OUTTAB![line detail]
   OUTTAB.Close

   ITEM = TRIBESTRANSFERS![ITEM]
   QUANTITY = TRIBESTRANSFERS![QUANTITY]

   If Not IsNull(ITEM) Then
      If Not ITEM = "" Then
         Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
         VALIDGOODS.index = "primarykey"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", ITEM
       
         If Not VALIDGOODS![TABLE] = "GENERAL" And Not VALIDGOODS![TABLE] = "HUMANS" Then
            Set TRIBESGOOD = TVDBGM.OpenRecordset("TRIBES_GOODS")
            TRIBESGOOD.index = "primarykey"
            TRIBESGOOD.MoveFirst
            TRIBESGOOD.Seek "=", TCLANNUMBER, INGOODSTRIBE, VALIDGOODS![TABLE], ITEM
         End If
       
         If VALIDGOODS![TABLE] = "HUMANS" Then
            Set TRIBESGOOD = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
            TRIBESGOOD.index = "primarykey"
            TRIBESGOOD.MoveFirst
            TRIBESGOOD.Seek "=", TCLANNUMBER, FROMTRIBE
            TRIBESGOOD.Edit

            If ITEM = "SLAVE" Then
               If QUANTITY >= TRIBESGOOD![SLAVE] Then
                  QUANTITY = TRIBESGOOD![SLAVE]
               End If
               TRIBESGOOD![SLAVE] = TRIBESGOOD![SLAVE] - QUANTITY
            ElseIf ITEM = "WARRIORS" Then
               If QUANTITY >= TRIBESGOOD![WARRIORS] Then
                  QUANTITY = TRIBESGOOD![WARRIORS]
               End If
               TRIBESGOOD![WARRIORS] = TRIBESGOOD![WARRIORS] - QUANTITY
            ElseIf ITEM = "ACTIVES" Then
               If QUANTITY >= TRIBESGOOD![ACTIVES] Then
                  QUANTITY = TRIBESGOOD![ACTIVES]
               End If
               TRIBESGOOD![ACTIVES] = TRIBESGOOD![ACTIVES] - QUANTITY
            ElseIf ITEM = "INACTIVES" Then
               If QUANTITY >= TRIBESGOOD![INACTIVES] Then
                  QUANTITY = TRIBESGOOD![INACTIVES]
               End If
               TRIBESGOOD![INACTIVES] = TRIBESGOOD![INACTIVES] - QUANTITY
            End If
            TRIBESGOOD.UPDATE

         ElseIf TRIBESGOOD.NoMatch Then
            QUANTITY = 0
         ElseIf QUANTITY >= TRIBESGOOD![ITEM_NUMBER] Then
            QUANTITY = TRIBESGOOD![ITEM_NUMBER]
            TRIBESGOOD.Edit
            TRIBESGOOD![ITEM_NUMBER] = TRIBESGOOD![ITEM_NUMBER] - QUANTITY
            TRIBESGOOD.UPDATE
         Else
            TRIBESGOOD.Edit
            TRIBESGOOD![ITEM_NUMBER] = TRIBESGOOD![ITEM_NUMBER] - QUANTITY
            TRIBESGOOD.UPDATE
         End If

         TRIBESGOOD.Close
      End If
   End If

    OUTPUTLINE = OUTPUTLINE & QUANTITY & " " & ITEM & ", "
    
    Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
    OUTTAB.index = "primarykey"
    OUTTAB.MoveFirst
    OUTTAB.Seek "=", TCLANNUMBER, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
    If OUTTAB.NoMatch Then
       OUTTAB.AddNew
       OUTTAB![CLAN] = TCLANNUMBER
       OUTTAB![TRIBE] = FROMTRIBE
       OUTTAB![Section] = "TRANSFERS OUT"
       OUTTAB![LINE NUMBER] = LINENUMBER
       OUTTAB![line detail] = OUTPUTLINE
       OUTTAB.UPDATE
       OUTTAB.Close
    Else
       OUTTAB.Edit
       OUTTAB![line detail] = OUTPUTLINE
       OUTTAB.UPDATE
       OUTTAB.Close
    End If

    LINENUMBER = 1

    Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
    OUTTAB.MoveFirst
    OUTTAB.index = "primarykey"
    OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", LINENUMBER

    If OUTTAB.NoMatch Then
       OUTTAB.AddNew
       OUTTAB![CLAN] = To_Clan
       OUTTAB![TRIBE] = TOTRIBE
       OUTTAB![Section] = "TRANSFERS IN"
       OUTTAB![LINE NUMBER] = LINENUMBER
       OUTTAB![line detail] = "Receive goods from " & FROMTRIBE & ": "
       OUTTAB.UPDATE
       OUTTAB.MoveFirst
       OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", LINENUMBER
       OUTTAB.Edit
    Else
       OUTTAB.MoveFirst
       OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", (LINENUMBER + 1)
       If OUTTAB.NoMatch Then
          OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", LINENUMBER
       Else
          OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", (LINENUMBER + 2)
          If OUTTAB.NoMatch Then
             LINENUMBER = LINENUMBER + 1
             OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", LINENUMBER
          Else
             LINENUMBER = LINENUMBER + 2
          End If
       End If
       OUTTAB.Edit
       ' IF LINE CONTAINS TOTRIBE THEN JUST COMMA ELSE
       POSITION = InStr(OUTTAB![line detail], FROMTRIBE)
       If POSITION > 0 Then
          OUTTAB![line detail] = OUTTAB![line detail] & " "
       Else
          OUTTAB![line detail] = OUTTAB![line detail] & "from " & FROMTRIBE & ": "
       End If
       OUTTAB.UPDATE
       OUTTAB.MoveFirst
       OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", LINENUMBER
       OUTTAB.Edit
    End If

    OUTPUTLINE = OUTTAB![line detail]

    If Not IsNull(ITEM) Then
       If Not ITEM = "" Then
          Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
          VALIDGOODS.index = "primarykey"
          VALIDGOODS.MoveFirst
          VALIDGOODS.Seek "=", ITEM
       
          If Not VALIDGOODS![TABLE] = "GENERAL" And Not VALIDGOODS![TABLE] = "HUMANS" Then
             Set TRIBESGOOD = TVDBGM.OpenRecordset("TRIBES_GOODS")
             TRIBESGOOD.index = "primarykey"
             TRIBESGOOD.MoveFirst
             TRIBESGOOD.Seek "=", To_Clan, TOGOODSTRIBE, VALIDGOODS![TABLE], ITEM
                
             If TRIBESGOOD.NoMatch Then
                TRIBESGOOD.AddNew
                TRIBESGOOD![CLAN] = To_Clan
                TRIBESGOOD![TRIBE] = TOGOODSTRIBE
                TRIBESGOOD![ITEM_TYPE] = VALIDGOODS![TABLE]
                TRIBESGOOD![ITEM] = ITEM
                TRIBESGOOD![ITEM_NUMBER] = QUANTITY
                TRIBESGOOD.UPDATE
                TRIBESGOOD.Close
             Else
                TRIBESGOOD.Edit
                TRIBESGOOD![ITEM_NUMBER] = TRIBESGOOD![ITEM_NUMBER] + QUANTITY
                TRIBESGOOD.UPDATE
                TRIBESGOOD.Close
             End If
          ElseIf VALIDGOODS![TABLE] = "HUMANS" Then
             Set TRIBESGOOD = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
             TRIBESGOOD.index = "primarykey"
             TRIBESGOOD.MoveFirst
             TRIBESGOOD.Seek "=", To_Clan, TOTRIBE
             TRIBESGOOD.Edit
 
             If ITEM = "SLAVE" Then
                TRIBESGOOD![SLAVE] = TRIBESGOOD![SLAVE] + QUANTITY
             ElseIf ITEM = "WARRIORS" Then
                TRIBESGOOD![WARRIORS] = TRIBESGOOD![WARRIORS] + QUANTITY
             ElseIf ITEM = "ACTIVES" Then
                TRIBESGOOD![ACTIVES] = TRIBESGOOD![ACTIVES] + QUANTITY
             ElseIf ITEM = "INACTIVES" Then
                TRIBESGOOD![INACTIVES] = TRIBESGOOD![INACTIVES] + QUANTITY
             End If
             TRIBESGOOD.UPDATE
         End If
     End If
       
  OUTPUTLINE = OUTPUTLINE & QUANTITY & " " & ITEM & ", "
          
  Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
  OUTTAB.index = "primarykey"
  OUTTAB.MoveFirst
  OUTTAB.Seek "=", To_Clan, TOTRIBE, "TRANSFERS IN", LINENUMBER
  If OUTTAB.NoMatch Then
     OUTTAB.AddNew
     OUTTAB![CLAN] = To_Clan
     OUTTAB![TRIBE] = TOTRIBE
     OUTTAB![Section] = "TRANSFERS IN"
     OUTTAB![LINE NUMBER] = LINENUMBER
     OUTTAB![line detail] = OUTPUTLINE
     OUTTAB.UPDATE
     OUTTAB.Close
  Else
     OUTTAB.Edit
     OUTTAB![line detail] = OUTPUTLINE
     OUTTAB.UPDATE
     OUTTAB.Close
  End If

  TRIBESTRANSFERS.Edit
  TRIBESTRANSFERS![PROCESSED] = "Y"
  TRIBESTRANSFERS.UPDATE

  ' RECALC CAPACITY & WEIGHT FOR EACH GROUP
  Call Determine_Capacities("group", TCLANNUMBER, FROMTRIBE)
  Call Determine_Capacities("group", To_Clan, TOTRIBE)
  Call Determine_Weights(TCLANNUMBER, FROMTRIBE)
  Call Determine_Weights(To_Clan, TOTRIBE)
End If
End If

NEXT_LOOP:

  TRIBESTRANSFERS.MoveNext
  If TRIBESTRANSFERS.EOF Then
     Exit Do
  End If



Loop



ERR_TRIBE_TRANSFERS_CLOSE:
   DoCmd.Hourglass False

   DoCmd.Close acForm, "TRANSFER_GOODS"
   DoCmd.OpenForm "TRANSFER_GOODS"

   Exit Function


ERR_TRIBE_TRANSFERS:
If (Err = 3021) Then
   Resume Next

Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_TRIBE_TRANSFERS_CLOSE
End If



End Function



Public Function TRIBES_BOOKS_TRANSFERS()
On Error GoTo ERR_TRIBES_BOOKS_TRANSFERS

' MODIFY DATABASE
Dim TRIBESBOOKS As Recordset, OUTTAB As Recordset
Dim INFILE As String
Dim FROMCLAN As String
Dim FROMTRIBE As String
Dim TOCLAN As String
Dim TOTRIBE As String
Dim ITEM As String
Dim OUTPUTLINE As String
Dim count As Long
Dim LINENUMBER As Long

DoCmd.Hourglass True

Set MYFORM = Forms![TRANSFER_BOOKS]

FROMCLAN = MYFORM![FROM CLAN]
FROMTRIBE = MYFORM![FROM TRIBE]
TOCLAN = MYFORM![TO CLAN]
TOTRIBE = MYFORM![TO TRIBE]

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
LINENUMBER = 1

Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
OUTTAB.index = "primarykey"
OUTTAB.MoveFirst
OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER

If OUTTAB.NoMatch Then
   OUTTAB.AddNew
   OUTTAB![CLAN] = MYFORM![FROM CLAN]
   OUTTAB![TRIBE] = MYFORM![FROM TRIBE]
   OUTTAB![Section] = "TRANSFERS OUT"
   OUTTAB![LINE NUMBER] = LINENUMBER
   OUTTAB![line detail] = "Transfer goods to " & MYFORM![TO TRIBE] & ": "
   OUTTAB.UPDATE
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
   OUTTAB.Edit
Else
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", (LINENUMBER + 1)
   If OUTTAB.NoMatch Then
      OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
   Else
      OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", (LINENUMBER + 2)
      If OUTTAB.NoMatch Then
         LINENUMBER = LINENUMBER + 1
         OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
      Else
         LINENUMBER = LINENUMBER + 2
      End If
   End If
   OUTTAB.Edit
   OUTTAB![line detail] = OUTTAB![line detail] & ",To " & TOTRIBE & ": "
   OUTTAB.UPDATE
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
   OUTTAB.Edit
End If

OUTPUTLINE = OUTTAB![line detail]
OUTTAB.Close


If Not IsNull(MYFORM![item01]) Then
   ITEM = MYFORM![item01]
End If

   If Not IsNull(ITEM) Then
      If Not ITEM = "" Then
         Set TRIBESBOOKS = TVDBGM.OpenRecordset("TRIBES_BOOKS")
         TRIBESBOOKS.index = "primarykey"
         TRIBESBOOKS.MoveFirst
         TRIBESBOOKS.Seek "=", FROMCLAN, FROMTRIBE, ITEM
       
         If Not TRIBESBOOKS.NoMatch Then
            ' MOVE TO NEW TRIBE
            If TRIBESBOOKS![NUMBER] > 1 Then
               TRIBESBOOKS.Edit
               TRIBESBOOKS![NUMBER] = TRIBESBOOKS![NUMBER] - 1
               TRIBESBOOKS.UPDATE
            ElseIf TRIBESBOOKS![NUMBER] = 1 Then
               TRIBESBOOKS.Delete
            End If
            TRIBESBOOKS.MoveFirst
            TRIBESBOOKS.Seek "=", TOCLAN, TOTRIBE, ITEM
            If Not TRIBESBOOKS.NoMatch Then
               TRIBESBOOKS.Edit
               TRIBESBOOKS![NUMBER] = TRIBESBOOKS![NUMBER] + 1
               TRIBESBOOKS.UPDATE
            Else
               TRIBESBOOKS.AddNew
               TRIBESBOOKS![CLAN] = TOCLAN
               TRIBESBOOKS![TRIBE] = TOTRIBE
               TRIBESBOOKS![BOOK] = ITEM
               TRIBESBOOKS![NUMBER] = 1
               TRIBESBOOKS.UPDATE
            End If
        End If

        TRIBESBOOKS.Close
    End If
    End If

      OUTPUTLINE = OUTPUTLINE & " " & ITEM & ", "
  
Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
OUTTAB.index = "primarykey"
OUTTAB.MoveFirst
OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
If OUTTAB.NoMatch Then
   OUTTAB.AddNew
   OUTTAB![CLAN] = FROMCLAN
   OUTTAB![TRIBE] = FROMTRIBE
   OUTTAB![Section] = "TRANSFERS OUT"
   OUTTAB![LINE NUMBER] = LINENUMBER
   OUTTAB![line detail] = OUTPUTLINE
   OUTTAB.UPDATE
   OUTTAB.Close
Else
   OUTTAB.Edit
   OUTTAB![line detail] = OUTPUTLINE
   OUTTAB.UPDATE
   OUTTAB.Close
End If

LINENUMBER = 1

Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
OUTTAB.MoveFirst
OUTTAB.index = "primarykey"
OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER

If OUTTAB.NoMatch Then
   OUTTAB.AddNew
   OUTTAB![CLAN] = TOCLAN
   OUTTAB![TRIBE] = TOTRIBE
   OUTTAB![Section] = "TRANSFERS IN"
   OUTTAB![LINE NUMBER] = LINENUMBER
   OUTTAB![line detail] = "Receive goods from " & FROMTRIBE & ": "
   OUTTAB.UPDATE
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER
   OUTTAB.Edit
Else
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", (LINENUMBER + 1)
   If OUTTAB.NoMatch Then
      OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER
   Else
      OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", (LINENUMBER + 2)
      If OUTTAB.NoMatch Then
         LINENUMBER = LINENUMBER + 1
         OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER
      Else
         LINENUMBER = LINENUMBER + 2
      End If
   End If
   OUTTAB.Edit
   OUTTAB![line detail] = OUTTAB![line detail] & ",from " & FROMTRIBE & ": "
   OUTTAB.UPDATE
   OUTTAB.MoveFirst
   OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER
   OUTTAB.Edit
End If

OUTPUTLINE = OUTTAB![line detail]

OUTPUTLINE = OUTPUTLINE & " " & ITEM & ", "
          
Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
OUTTAB.index = "primarykey"
OUTTAB.MoveFirst
OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER
If OUTTAB.NoMatch Then
   OUTTAB.AddNew
   OUTTAB![CLAN] = TOCLAN
   OUTTAB![TRIBE] = TOTRIBE
   OUTTAB![Section] = "TRANSFERS IN"
   OUTTAB![LINE NUMBER] = LINENUMBER
   OUTTAB![line detail] = OUTPUTLINE
   OUTTAB.UPDATE
   OUTTAB.Close
Else
   OUTTAB.Edit
   OUTTAB![line detail] = OUTPUTLINE
   OUTTAB.UPDATE
   OUTTAB.Close
End If

ERR_TRIBES_BOOKS_TRANSFERS_CLOSE:
   DoCmd.Hourglass False

   DoCmd.Close acForm, "TRANSFER_BOOKS"
   DoCmd.OpenForm "TRANSFER_BOOKS"

   Exit Function


ERR_TRIBES_BOOKS_TRANSFERS:
If (Err = 3021) Then
   Resume Next

Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_TRIBES_BOOKS_TRANSFERS_CLOSE
End If


End Function
