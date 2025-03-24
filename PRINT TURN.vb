Attribute VB_Name = "PRINT TURN"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*                              VERSION 3.1.1                                    *'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'** 25/02/96  *  Allow for print from clan                                      **'
'** 09/06/96  *  Allow for skill attempts                                       **'
'** 17/07/96  *  Allow for pacification levels                                  **'
'** 25/01/17  *  Allow for Office 2010 and later                                **'
'** 08/06/18  *  Added transfer section at end of report                        **'
'** 05/03/25  *  Current and previous Locations are printed (AlexD)             **'
'*===============================================================================*'
 
' MODULE NAME = PRINT TURN
' Program Flow
' First Function called is A100-Print-Turn
'

Global Turncontrol As Control
Global OutPutTab As Recordset
Global TRIBEINFO As Recordset, Tribes_Goods, VALIDGOODS As Recordset
Global ClanList As Recordset
Global UnitList As Recordset
Global HEXCONSTTABLE As Recordset, SkillsTab As Recordset
Global Globaltable As Recordset, COMPRESTABLE As Recordset
Global RELIGIONTABLE As Recordset, HEXTABLE As Recordset
Global ValidSkills As Recordset, UNDERCONSTTABLE As Recordset
Global RESEARCHTABLE As Recordset, FARMTABLE As Recordset
Global HERDSWAPTABLE As Recordset, CONSTTABLE As Recordset
Global LINENUMBERTABLE As Recordset, VALIDBUILDINGS As Recordset
Global MassXfers As Recordset
Global Special_Routes As Recordset
Global Perm_Mess_Tab As Recordset
Global AccApp As Access.Application
Global wrdApp As Word.Application
Global wrdDoc As Word.Document
Global wrdSel As Word.Selection
Global ClanArray As Variant
Global ClanCount, UnitCount, i As Integer
Global GM As String
Global SEASON As String
Global NextTurn As String
Global OutLine As String
Global OutCount As Long
Global PrintOutLine(10, 300) As String
Global FIRST_GROUP(300) As String
Global SECOND_GROUP(300) As String
Global ROW As Long
Global COLUMN As Long
Global FIRST_COUNT As Long
Global SECOND_COUNT As Long
Global FIRST_TAB_COUNT As Long
Global SECOND_TAB_COUNT As Long
Global Section As String
Global TVDirect As String
Global fileName As String
Global TurnNum As String
Global TribeI As Long
Global LineI As Long
Global TotalPeople As Long
Global TRIBENUMBER As String
Global CLANNUMBER As String
Global Village As String
Global bGT As Boolean ' does unit have a goods tribe
Global CurrentHex As String
Global PreviousHex As String
Global NumGoods As Long
Global TRIBES_IN_HEX As String
Global NUM_CHARS As Long
Global POLITICAL_LEVEL As Long
Global CONSTRUCTIONLINE As String
Global CONST_LINECOUNT As Long
Global CONST_CLAN As String
Global CONST_TRIBE As String
Global FROM_CLANNUMBER As String
Global TO_CLANNUMBER As String
Global CURRENT_HEX_MAP As String
Global SECTION_NAME As String
Global DIRECTPATH As String
Global DOCUMENTPATH As String
Global OUTPUT_TYPE As String
Global CROP(20) As String
Global CROP_FOUND(20) As String
Global CROP_AMOUNT(20, 12) As Long
Global FARM_TURN(12) As String
Global VALID_FARMING_TURN As String
Global NONSKILLED As Long
Global COMPRESFOUND As String
Global times_through As Long
Global Program_Area As String
Global SURROUNDING_DATA As String
Global MAX_COUNT As Long
Global STOP_PROCESSING As String
Global String_Found1 As Integer
Global String_Found2 As Integer
Global String_Start As Integer
Global String_Length As Integer
Global Chars_Read As Integer
Global first_B As String
Global Transfers_found As String
Global Movement_found As String
Global Scouting_found As String
Global First_Const As String
Global PlayerSpreadsheet As String
Global MessageText As String
Global DR_Count As Long
Global Boatshed_Req As Long
Global Min_Crew_Req As Long
Global Max_Crew As Long
Global Max_Cargo As Long
Global Ship_Found As String
Global numConstructionPrintingTrashold As Long

Function A100_Print_Turn()
On Error GoTo ERR_A100_PRINT
TRIBE_STATUS = "A100 Print Turn"

DebugOP ("A100 Print Turn")

Call Tribe_Checking("Update_All", "", "", "")

DoCmd.Hourglass True
STOP_PROCESSING = "NO"

Set wrdApp = Nothing
Set wrdDoc = Nothing

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GM = GMTABLE![Name]
FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set Perm_Mess_Tab = TVDBGM.OpenRecordset("Permanent_Messages_Table")
Perm_Mess_Tab.index = "PRIMARYKEY"
Perm_Mess_Tab.MoveFirst

Set Globaltable = TVDBGM.OpenRecordset("Global")
Globaltable.index = "PRIMARYKEY"
Globaltable.MoveFirst

Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst

Set HEXMAPCITY = TVDBGM.OpenRecordset("HEX_MAP_CITY")
HEXMAPCITY.index = "PRIMARYKEY"
HEXMAPCITY.MoveFirst

Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
HEXMAPMINERALS.index = "PRIMARYKEY"
HEXMAPMINERALS.MoveFirst

Set HEXMAPPOLITICS = TVDBGM.OpenRecordset("HEX_MAP_POLITICS")
HEXMAPPOLITICS.index = "PRIMARYKEY"
If Not HEXMAPPOLITICS.EOF Then
   HEXMAPPOLITICS.MoveFirst
End If

Set HEXCONSTTABLE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXCONSTTABLE.index = "PRIMARYKEY"
If Not HEXCONSTTABLE.EOF Then
   HEXCONSTTABLE.MoveFirst
End If

Set HERDSWAPTABLE = TVDBGM.OpenRecordset("HERD_SWAPS")
HERDSWAPTABLE.index = "TRIBE"
HERDSWAPTABLE.MoveFirst

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "PRIMARYKEY"
VALIDGOODS.MoveFirst

Set VALIDSHIPS = TVDB.OpenRecordset("VALID_SHIPS")
VALIDSHIPS.index = "PRIMARYKEY"
VALIDSHIPS.MoveFirst

Set ClanList = TVDBGM.OpenRecordset("select distinct TI1.clan, TI2.spreadsheet from (Tribes_General_Info as TI1 " & _
    "left outer join (select tribe, spreadsheet from TRIBES_GENERAL_INFO) as TI2 on TI1.clan=TI2.tribe) " & _
    "where clan >='" & Forms![PRINT_FROM_CLAN]![FROM_CLANNUMBER] & "' AND clan <='" & _
    Forms![PRINT_FROM_CLAN]![TO_CLANNUMBER] & "' " & _
    "order by clan")
ClanList.MoveFirst
ClanCount = ClanList.RecordCount
'ClanArray = ClanList.GetRows
'ClanList.Close

Set UnitList = TVDBGM.OpenRecordset("select distinct Tribes_General_Info.Tribe as Tribe FROM TRIBES_GENERAL_INFO;")
UnitList.MoveFirst
UnitCount = UnitList.RecordCount


Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst

Set Tribes_Goods = TVDBGM.OpenRecordset("Tribes_Goods")
Tribes_Goods.index = "SECONDARYKEY"
Tribes_Goods.MoveFirst

Set OutPutTab = TVDBGM.OpenRecordset("Turns_Activities")
OutPutTab.index = "PRIMARYKEY"
OutPutTab.MoveFirst

Set MassXfers = TVDBGM.OpenRecordset("SELECT MASSTRANSFERS.REPORT_CLAN as CLAN, MASSTRANSFERS.FROM as FROMUNIT, " & _
    "MASSTRANSFERS.TO as TOUNIT, MASSTRANSFERS.ITEM as ITEM, MASSTRANSFERS.QUANTITY as QUANTITY, " & _
    "MASSTRANSFERS.ACTUAL_QTY as ACTUAL_QTY, MASSTRANSFERS.PROCESS_MSG as PROCESS_MSG, MASSTRANSFERS.REPORT_CODE AS RPT_CODE " & _
    "FROM MASSTRANSFERS;")

Set Special_Routes = TVDBGM.OpenRecordset("SELECT " & _
    "HMC1.OWNER AS OWNER, SR.ROUTE_NAME AS ROUTE_NAME, HMC1.TYPE AS ROUTE_TYPE, HMC1.SUBTYPE AS SUBTYPE, " & _
    "SR.FROM_HEX AS FROM_HEX, SR.TO_HEX AS TO_HEX " & _
    "FROM SPECIAL_TRANSFER_ROUTES AS SR INNER JOIN " & _
    "(SELECT * FROM HEX_MAP_CITY WHERE TYPE='Named site' AND SUBTYPE='Player') AS HMC1 ON SR.FROM_HEX=HMC1.MAP " & _
    "UNION SELECT " & _
    "HMC1.OWNER AS OWNER, SR.ROUTE_NAME AS ROUTE_NAME, HMC1.TYPE AS ROUTE_TYPE, HMC1.SUBTYPE AS SUBTYPE, " & _
    "SR.FROM_HEX AS FROM_HEX, SR.TO_HEX AS TO_HEX " & _
    "FROM SPECIAL_TRANSFER_ROUTES AS SR INNER JOIN " & _
    "(SELECT * FROM HEX_MAP_CITY WHERE TYPE='Named site' AND SUBTYPE='Player') AS HMC1 ON SR.TO_HEX=HMC1.MAP " & _
    "ORDER BY ROUTE_NAME;")

TVDirect = Mid(Globaltable![CURRENT TURN], 1, 2) & Right(Globaltable![CURRENT TURN], 3)
TurnNum = Globaltable![CURRENT TURN]

If Left(Globaltable![CURRENT TURN], 2) = 12 Then
   NextTurn = "01/" & Right(Globaltable![CURRENT TURN], 3) + 1
Else
   NextTurn = (Left(Globaltable![CURRENT TURN], 2) + 1) & Right(Globaltable![CURRENT TURN], 4)
End If

SEASON = GET_SEASON(Globaltable![CURRENT TURN])

If Left(Globaltable![CURRENT TURN], 2) = "1/" Then
   NextTurn = "01" & Right(Globaltable![CURRENT TURN], 4)
End If

' Open Word
Call Open_Word(GM)
wrdApp.Visible = False

CURRENT_DIRECTORY = Dir(DIRECTPATH, vbDirectory)
If IsNull(CURRENT_DIRECTORY) Or CURRENT_DIRECTORY = "" Then
  MkDir (DIRECTPATH)
End If

FROM_CLANNUMBER = Forms![PRINT_FROM_CLAN]![FROM_CLANNUMBER]
TO_CLANNUMBER = Forms![PRINT_FROM_CLAN]![TO_CLANNUMBER]

TRIBEINFO.MoveFirst
TRIBEINFO.Seek "=", FROM_CLANNUMBER, FROM_CLANNUMBER
TRIBENUMBER = TRIBEINFO![TRIBE]

CLANNUMBER = FROM_CLANNUMBER
If IsNull(TRIBEINFO![OUTPUT_TYPE]) Then
   OUTPUT_TYPE = "WORD"
Else
   OUTPUT_TYPE = TRIBEINFO![OUTPUT_TYPE]
End If

Do While TRIBEINFO![CLAN] >= FROM_CLANNUMBER And TRIBEINFO![CLAN] <= TO_CLANNUMBER
   DebugOP ("Processing Clan " & CLANNUMBER)
   Forms![PRINT_FROM_CLAN]![Status] = "Processing Clan " & CLANNUMBER
   
   
   Forms![PRINT_FROM_CLAN].Repaint

   Call A300_UPDATE_ACTIVITIES
   
   Call Print_Mass_Transfers
   
   Call Print_Settlements
   
   Call Print_Special_Routes
      
   Call Save_Tribes_turn
      
   If STOP_PROCESSING = "YES" Then
      Exit Function
   End If

   If Not TRIBEINFO.EOF Then
      If Not TRIBEINFO![CLAN] = CLANNUMBER Then
         CLANNUMBER = TRIBEINFO![CLAN]
         If IsNull(TRIBEINFO![OUTPUT_TYPE]) Then
            OUTPUT_TYPE = "WORD"
         Else
            OUTPUT_TYPE = TRIBEINFO![OUTPUT_TYPE]
         End If
      End If
   Else
      Exit Do
   End If
Loop


Call CLOSE_WORD

Do While ClanList![CLAN] >= FROM_CLANNUMBER And ClanList![CLAN] <= TO_CLANNUMBER
   DebugOP ("Processing Orders Sheet for Clan " & ClanList![CLAN])
   Forms![PRINT_FROM_CLAN]![Status] = "Processing Orders Sheet for Clan " & ClanList![CLAN]
   Forms![PRINT_FROM_CLAN].Repaint
   DebugOP "Processing Orders Sheet (A100) for Clan " & ClanList![CLAN]
    
   If IsNull(ClanList![Spreadsheet]) Then
        PlayerSpreadsheet = "NotExcel"
   Else
        PlayerSpreadsheet = ClanList![Spreadsheet]
   End If
   
   Call create_workbook(Mid(ClanList![CLAN], 2, 3), PlayerSpreadsheet)
   
   DoEvents
   
   ClanList.MoveNext

   If ClanList.EOF Then
        Forms![PRINT_FROM_CLAN]![Status] = "Processing is complete."
        Forms![PRINT_FROM_CLAN].Repaint
        Exit Do
   End If
Loop

RESEARCHTABLE.Close
TRIBEINFO.Close
Tribes_Goods.Close
MassXfers.Close
VALIDGOODS.Close
Perm_Mess_Tab.Close
Globaltable.Close
HEXTABLE.Close
HEXMAPCITY.Close
HEXMAPMINERALS.Close
HEXMAPPOLITICS.Close
HEXCONSTTABLE.Close
HERDSWAPTABLE.Close
ClanList.Close

TVDBGM.Close
TVDB.Close

DebugOP ("END-A100 Print Turn")
   
'For i = 0 To (ClanCount - 1)

'   Call create_workbook(Mid(ClanArray(i, 0), 2, 3))

'Next i

ERR_A100_PRINT_CLOSE:
   DoCmd.Hourglass False
   Exit Function


ERR_A100_PRINT:
If (Err = 3021) Then
   Resume Next
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_A100_PRINT_CLOSE
  
End If

End Function
Function A200_Generate_Spreadsheets()
On Error GoTo ERR_A200_PRINT
TRIBE_STATUS = "A200 Generate Spreadsheets"

Call Tribe_Checking("Update_All", "", "", "")

DoCmd.Hourglass True
STOP_PROCESSING = "NO"

Set wrdApp = Nothing
Set wrdDoc = Nothing

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

GM = GMTABLE![Name]
FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set Perm_Mess_Tab = TVDBGM.OpenRecordset("Permanent_Messages_Table")
Perm_Mess_Tab.index = "PRIMARYKEY"
Perm_Mess_Tab.MoveFirst

Set Globaltable = TVDBGM.OpenRecordset("Global")
Globaltable.index = "PRIMARYKEY"
Globaltable.MoveFirst

Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst

Set HEXMAPCITY = TVDBGM.OpenRecordset("HEX_MAP_CITY")
HEXMAPCITY.index = "PRIMARYKEY"
HEXMAPCITY.MoveFirst

Set HEXMAPMINERALS = TVDBGM.OpenRecordset("HEX_MAP_MINERALS")
HEXMAPMINERALS.index = "PRIMARYKEY"
HEXMAPMINERALS.MoveFirst

Set HEXMAPPOLITICS = TVDBGM.OpenRecordset("HEX_MAP_POLITICS")
HEXMAPPOLITICS.index = "PRIMARYKEY"
If Not HEXMAPPOLITICS.EOF Then
   HEXMAPPOLITICS.MoveFirst
End If

Set HEXCONSTTABLE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXCONSTTABLE.index = "PRIMARYKEY"
If Not HEXCONSTTABLE.EOF Then
   HEXCONSTTABLE.MoveFirst
End If

Set HERDSWAPTABLE = TVDBGM.OpenRecordset("HERD_SWAPS")
HERDSWAPTABLE.index = "TRIBE"
HERDSWAPTABLE.MoveFirst

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "PRIMARYKEY"
VALIDGOODS.MoveFirst

Set ClanList = TVDBGM.OpenRecordset("select distinct TI1.clan, TI2.spreadsheet from (Tribes_General_Info as TI1 " & _
    "left outer join (select tribe, spreadsheet from TRIBES_GENERAL_INFO) as TI2 on TI1.clan=TI2.tribe) " & _
    "where clan >='" & Forms![PRINT_FROM_CLAN]![FROM_CLANNUMBER] & "' AND clan <='" & _
    Forms![PRINT_FROM_CLAN]![TO_CLANNUMBER] & "' " & _
    "order by clan")
ClanList.MoveFirst
ClanCount = ClanList.RecordCount
'ClanArray = ClanList.GetRows
'ClanList.Close

Set UnitList = TVDBGM.OpenRecordset("select distinct Tribes_General_Info.Tribe as Tribe FROM TRIBES_GENERAL_INFO;")
UnitList.MoveFirst
UnitCount = UnitList.RecordCount

Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst

Set Tribes_Goods = TVDBGM.OpenRecordset("Tribes_Goods")
Tribes_Goods.index = "SECONDARYKEY"
Tribes_Goods.MoveFirst

Set OutPutTab = TVDBGM.OpenRecordset("Turns_Activities")
OutPutTab.index = "PRIMARYKEY"
OutPutTab.MoveFirst

Set MassXfers = TVDBGM.OpenRecordset("SELECT MASSTRANSFERS.REPORT_CLAN as CLAN, MASSTRANSFERS.FROM as FROMUNIT, " & _
    "MASSTRANSFERS.TO as TOUNIT, MASSTRANSFERS.ITEM as ITEM, MASSTRANSFERS.QUANTITY as QUANTITY, " & _
    "MASSTRANSFERS.ACTUAL_QTY as ACTUAL_QTY, MASSTRANSFERS.PROCESS_MSG as PROCESS_MSG, MASSTRANSFERS.REPORT_CODE AS RPT_CODE " & _
    "FROM MASSTRANSFERS;")

TVDirect = Mid(Globaltable![CURRENT TURN], 1, 2) & Right(Globaltable![CURRENT TURN], 3)
TurnNum = Globaltable![CURRENT TURN]

If Left(Globaltable![CURRENT TURN], 2) = 12 Then
   NextTurn = "01/" & Right(Globaltable![CURRENT TURN], 3) + 1
Else
   NextTurn = (Left(Globaltable![CURRENT TURN], 2) + 1) & Right(Globaltable![CURRENT TURN], 4)
End If

SEASON = GET_SEASON(Globaltable![CURRENT TURN])

If Left(Globaltable![CURRENT TURN], 2) = "1/" Then
   NextTurn = "01" & Right(Globaltable![CURRENT TURN], 4)
End If

CURRENT_DIRECTORY = Dir(DIRECTPATH, vbDirectory)
If IsNull(CURRENT_DIRECTORY) Or CURRENT_DIRECTORY = "" Then
  MkDir (DIRECTPATH)
End If

FROM_CLANNUMBER = Forms![PRINT_FROM_CLAN]![FROM_CLANNUMBER]
TO_CLANNUMBER = Forms![PRINT_FROM_CLAN]![TO_CLANNUMBER]

TRIBEINFO.MoveFirst
TRIBEINFO.Seek "=", FROM_CLANNUMBER, FROM_CLANNUMBER
TRIBENUMBER = TRIBEINFO![TRIBE]

CLANNUMBER = FROM_CLANNUMBER

Do While ClanList![CLAN] >= FROM_CLANNUMBER And ClanList![CLAN] <= TO_CLANNUMBER
   Forms![PRINT_FROM_CLAN]![Status] = "Processing Orders Sheet for Clan " & ClanList![CLAN]
   Forms![PRINT_FROM_CLAN].Repaint
   DebugOP "Processing Orders Sheet (A200) for Clan " & ClanList![CLAN]
    
   If IsNull(ClanList![Spreadsheet]) Then
        PlayerSpreadsheet = "NotExcel"
   Else
        PlayerSpreadsheet = ClanList![Spreadsheet]
   End If
   
   Call create_workbook(Mid(ClanList![CLAN], 2, 3), PlayerSpreadsheet)
   
   DoEvents
   
   ClanList.MoveNext

   If ClanList.EOF Then
        Forms![PRINT_FROM_CLAN]![Status] = "Processing is complete."
        Forms![PRINT_FROM_CLAN].Repaint
        Exit Do
   End If
Loop

TRIBEINFO.Close
Tribes_Goods.Close
MassXfers.Close
VALIDGOODS.Close
Perm_Mess_Tab.Close
Globaltable.Close
HEXTABLE.Close
HEXMAPCITY.Close
HEXMAPMINERALS.Close
HEXMAPPOLITICS.Close
HEXCONSTTABLE.Close
HERDSWAPTABLE.Close
ClanList.Close

TVDBGM.Close
TVDB.Close
   
ERR_A200_PRINT_CLOSE:
   DoCmd.Hourglass False
   Exit Function


ERR_A200_PRINT:
If (Err = 3021) Then
   Resume Next
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_A200_PRINT_CLOSE
  
End If

End Function
Sub A300_UPDATE_ACTIVITIES()
On Error GoTo ERR_A300_UPDATE
TRIBE_STATUS = "A300 Update Activities"

wrdApp.Visible = False

Program_Area = "Start of A300"

first_B = "No"
Call Delete_Existing_Turn
Call OPEN_NEW_FILE
If STOP_PROCESSING = "YES" Then
   Exit Sub
End If

Do While TRIBEINFO![CLAN] = CLANNUMBER
   Set wrdSel = wrdApp.Selection
    
   TRIBENUMBER = TRIBEINFO![TRIBE]
   If TRIBENUMBER = CLANNUMBER Then
      'continue
   ElseIf Left(TRIBEINFO![TRIBE], 1) = Left(CLANNUMBER, 1) Then
      'continue
'      wrdApp.Selection.TypeText vbCrLf & vbCrLf
      wrdApp.Selection.InsertBreak TYPE:=wdPageBreak
   Else
      ' new page
      wrdApp.Selection.InsertBreak TYPE:=wdPageBreak
   End If
   
    Village = TRIBEINFO![Village]
    If Village = "Tribe" Then
        wrdApp.Selection.Style = wdStyleHeading1
    Else
        wrdApp.Selection.Style = wdStyleHeading2
    End If
   
   
   TRIBENUMBER = TRIBEINFO![TRIBE]
   DebugOP "Writing to MS Word: " & TRIBENUMBER
  
SECTION_NAME = "GLOBAL"

    wrdApp.Selection.Font.Bold = True
    wrdApp.Selection.TypeText TRIBEINFO![Village] & " " & TRIBEINFO![TRIBE]
'    wrdApp.Selection.Style = wdStyleHeading1
    wrdApp.Selection.Font.Bold = False
    wrdApp.Selection.Font.Italic = True
    wrdApp.Selection.TypeText ", " & TRIBEINFO![TRIBE NAME]
    wrdApp.Selection.Font.Italic = False
    
    CurrentHex = TRIBEINFO![CURRENT HEX]
    PreviousHex = Nz(TRIBEINFO![Previous_Hex], "")
    
    'remove Grid Ref
'    CurrentHex = "## " & Mid(CurrentHex, 4, 4)
    
    If Len(PreviousHex) <= 0 Then
        PreviousHex = "N/A"
'    Else
'        PreviousHex = "## " & Mid(PreviousHex, 4, 4)
    End If
    
    
    wrdApp.Selection.TypeText _
            ", Current Hex = " & CurrentHex & _
            ", (Previous Hex = " & _
            PreviousHex & ")"
    
    wrdApp.Selection.TypeText vbCr
    wrdApp.Selection.Style = wdStyleNormal
    
    
   Call TABS_REQUIRED(SECTION_NAME)
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
   'Set Bold
   

    NextTurn = ChangeTurnYYYMM(GetCurrentTurn(), 1)

   wrdApp.Selection.TypeText "Current Turn " & _
                            GetCurrentTurn() & _
                            " (#" & GetCurrentTurnNo() & ")" & _
                            ", " & SEASON

   Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
   HEXTABLE.index = "PRIMARYKEY"
   HEXTABLE.MoveFirst
   HEXTABLE.Seek "=", TRIBEINFO![CURRENT HEX]
   
   If HEXTABLE![WEATHER_ZONE] = "GREEN" Then
      wrdApp.Selection.TypeText ", " & Globaltable![Zone1]
   ElseIf HEXTABLE![WEATHER_ZONE] = "RED" Then
      wrdApp.Selection.TypeText ", " & Globaltable![Zone2]
   ElseIf HEXTABLE![WEATHER_ZONE] = "ORANGE" Then
      wrdApp.Selection.TypeText ", " & Globaltable![Zone3]
   ElseIf HEXTABLE![WEATHER_ZONE] = "YELLOW" Then
      wrdApp.Selection.TypeText ", " & Globaltable![Zone4]
   ElseIf HEXTABLE![WEATHER_ZONE] = "BLUE" Then
      wrdApp.Selection.TypeText ", " & Globaltable![Zone5]
   ElseIf HEXTABLE![WEATHER_ZONE] = "BROWN" Then
      wrdApp.Selection.TypeText ", " & Globaltable![Zone6]
   End If
    
    wrdApp.Selection.TypeText vbTab
    
    If TRIBENUMBER = CLANNUMBER Then
        wrdApp.Selection.TypeText "Next Turn " & _
                                NextTurn & _
                                " (#" & GetCurrentTurnNo() + 1 & ")" & _
                                ", " & Globaltable![Date Due] & vbCr
    Else
        wrdApp.Selection.TypeText vbCr
   
    End If
    
   If TRIBEINFO![TRIBE] = CLANNUMBER Then

   End If

   If TRIBEINFO![TRIBE] = CLANNUMBER Then
        wrdApp.Selection.TypeText "Received: $"
        wrdApp.Selection.TypeText TRIBEINFO![AMT RECEIVED] & ", Cost: $ "
        wrdApp.Selection.TypeText TRIBEINFO![Cost]
        
        If TRIBEINFO![CREDIT] < 0 Then
            wrdApp.Selection.Font.Color = vbRed
            wrdApp.Selection.TypeText vbTab & "Credit: $ " & TRIBEINFO![CREDIT] & vbCr
            wrdApp.Selection.Font.Color = vbBlack
        Else
            wrdApp.Selection.TypeText vbTab & "Credit: $ " & TRIBEINFO![CREDIT] & vbCr
        End If
   End If
   
   
    If IsNull(TRIBEINFO![GOODS TRIBE]) Or TRIBENUMBER = TRIBEINFO![GOODS TRIBE] Then
        bGT = False
    Else
        bGT = True
    End If
    
    If bGT Then
        wrdApp.Selection.TypeText "Goods Tribe: " & _
                                    TRIBEINFO![GOODS TRIBE] & vbCrLf & vbCrLf
    Else
        wrdApp.Selection.TypeText "Goods Tribe: No GT" & vbCrLf & vbCrLf
    End If
   
   
   
   If TRIBEINFO![TRIBE] = CLANNUMBER Then
     'Print out commodities as a concatenated string from Clan_DesiredCommodities Table

        Dim sSQL As String
        Dim sCommodities As String
        sSQL = "SELECT  ""("" & [DC_Index] & "") "" & [DC_DesiredCommodity] AS ConcatOP " & vbCrLf & _
                "From Clan_DesiredCommodities " & vbCrLf & _
                "WHERE Clan_DesiredCommodities.DC_CLAN='" & _
                CLANNUMBER & _
                "' " & _
                "ORDER BY DC_Index ASC;"
        sCommodities = ConcatRelatedSQL(sSQL, ", ")
        If Len(Nz(sCommodities, "")) <= 0 Then
            sCommodities = "No commodities allocated"
        End If
                                    
        wrdApp.Selection.Font.Bold = True
        wrdApp.Selection.TypeText "Desired Commodities: "
        wrdApp.Selection.Font.Bold = False
        wrdApp.Selection.TypeText sCommodities
        wrdApp.Selection.TypeText vbCrLf & vbCrLf
        sSQL = ""
    End If
    
    '============Special Hexes====================
    Dim sOutputTitle As String
    Dim sOutputType As String
    Dim sOutputSubType As String
    Dim sOutputDescription As String
    Dim sCurrentHex As String
    
    sCurrentHex = TRIBEINFO![CURRENT HEX]
    
    If Not IsNull(ELookup("MAP", _
                        "HEX_MAP_CITY", _
                        "MAP = '" & sCurrentHex & "'")) Then
                        
        sOutputTitle = Nz(ELookup("CITY", _
                    "HEX_MAP_CITY", _
                    "MAP = '" & sCurrentHex & "'"), "")
                    
        sOutputType = Nz(ELookup("TYPE", _
                    "HEX_MAP_CITY", _
                    "MAP = '" & sCurrentHex & "'"), "")
                    
        sOutputSubType = Nz(ELookup("SUBTYPE", _
                    "HEX_MAP_CITY", _
                    "MAP = '" & sCurrentHex & "'"), "")
                    
        sOutputDescription = Nz(ELookup("OFFERTEXT", _
                    "HEX_MAP_CITY", _
                    "MAP = '" & sCurrentHex & "'"), "")
        
        
        wrdApp.Selection.Font.Color = HEXtoLong("ED7D31")
        wrdApp.Selection.Font.Bold = True
        wrdApp.Selection.TypeText "Special Hex"
        wrdApp.Selection.Font.Bold = False
        wrdApp.Selection.Font.Color = vbBlack
        wrdApp.Selection.TypeText vbCrLf
        
        wrdApp.Selection.TypeText sOutputTitle & " (" & _
                                    sOutputType & _
                                    "-" & _
                                    sOutputSubType & ")" & _
                                    vbCrLf
                                    
        wrdApp.Selection.Font.Italic = True
        wrdApp.Selection.TypeText sOutputDescription
        wrdApp.Selection.Font.Italic = False
        wrdApp.Selection.TypeText vbCrLf & vbCrLf
        
        
        
    End If
    
    '============/special hexes================
    
   
SECTION_NAME = "PERM MESSAGE"
   Call TABS_REQUIRED(SECTION_NAME)
   If TRIBEINFO![TRIBE] = CLANNUMBER Then
      ' print the permanent message if any
      Perm_Mess_Tab.MoveFirst
      Perm_Mess_Tab.Seek "=", "all", "all"
      
      If Perm_Mess_Tab.NoMatch Then
         'No Permanent Message
      Else
         wrdApp.Selection.TypeText Perm_Mess_Tab![Message] & vbCr
      End If
         
      Perm_Mess_Tab.MoveFirst
      Perm_Mess_Tab.Seek "=", CLANNUMBER, TRIBENUMBER
     
      If Perm_Mess_Tab.NoMatch Then
         'No Permanent Message
      Else
         wrdApp.Selection.TypeText Perm_Mess_Tab![Message] & vbCr
      End If
   Else
      ' check for group specific messages
      Perm_Mess_Tab.MoveFirst
      Perm_Mess_Tab.Seek "=", CLANNUMBER, TRIBENUMBER
     
      If Perm_Mess_Tab.NoMatch Then
         'No Permanent Message
      Else
         wrdApp.Selection.TypeText Perm_Mess_Tab![Message] & vbCr
      End If
   End If

wrdApp.Selection.TypeText vbCrLf

SECTION_NAME = "COMMENTS"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "ENCOUNTERS"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "RESPONSE"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "TRADING POST SOLD"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "TRADING POST BUY"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "ACTIVITIES"
   Call FIND_AND_PRINT_ACTIVITIES

   ' Sleep for 1 second
   'Call sleep(1000)
 
   'blank line after activities
   wrdApp.Selection.TypeText vbCrLf
   wrdApp.Selection.TypeText vbCrLf

SECTION_NAME = "TRANSFERS OUT"
   Transfers_found = "NO"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "TRANSFERS IN"
   Call FIND_AND_PRINT_ACTIVITIES
   
   If Transfers_found = "YES" Then
      'blank line after transfers, if there are transfers
      wrdApp.Selection.TypeText vbCrLf
   End If
   
SECTION_NAME = "HERD SWAPS"
Program_Area = "A300 - Herd Swaps"

   ' DO HERD SWAPS
   HERDSWAPTABLE.index = "TRIBE ONLY"
   HERDSWAPTABLE.MoveFirst
   HERDSWAPTABLE.Seek "=", TRIBENUMBER
   If HERDSWAPTABLE.NoMatch Then
      'NO HERDSWAP FOUND
   Else
      Do While HERDSWAPTABLE![TRIBE] = TRIBENUMBER
         wrdApp.Selection.TypeText HERDSWAPTABLE![ANIMAL] & " swapped with "
         wrdApp.Selection.TypeText HERDSWAPTABLE![tribe swapped with] & " - You have "
         wrdApp.Selection.TypeText HERDSWAPTABLE![turns to go] & " turns of benefit left"
        HERDSWAPTABLE.MoveNext
        If HERDSWAPTABLE.EOF Then
           Exit Do
        End If
      Loop
   End If
  
SECTION_NAME = "TRIBE MOVEMENT"
Program_Area = "A300 - Tribe Movement"
   Movement_found = "NO"
   Call FIND_AND_PRINT_ACTIVITIES

   'blank line after movement
   If Movement_found = "Yes" Then
      wrdApp.Selection.TypeText vbCrLf
   End If
   
SECTION_NAME = "SCOUT  1 MOVEMENT"
Program_Area = "A300 - Scout 1 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  2 MOVEMENT"
Program_Area = "A300 - Scout 2 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  3 MOVEMENT"
Program_Area = "A300 - Scout 3 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  4 MOVEMENT"
Program_Area = "A300 - Scout 4 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  5 MOVEMENT"
Program_Area = "A300 - Scout 5 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  6 MOVEMENT"
Program_Area = "A300 - Scout 6 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  7 MOVEMENT"
Program_Area = "A300 - Scout 7 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "SCOUT  8 MOVEMENT"
Program_Area = "A300 - Scout 8 Movement"
   Call FIND_AND_PRINT_ACTIVITIES

SECTION_NAME = "FARMING"
Program_Area = "A300 - Farming"
Call TABS_REQUIRED(SECTION_NAME)
   
' initialise values
   
'this value holds the weather for each turn
count = 0
Do Until count > 12
   FARM_TURN(count) = "EMPTY"
   count = count + 1
Loop

'this value holds the crops and if they are found in the farming table
count = 0
Do Until count > 20
   CROP(count) = "EMPTY"
   CROP_FOUND(count) = "NO"
   count = count + 1
Loop

'this value holds the amount for each of the turns
count = 0
COUNT1 = 0
Do Until count > 20
   Do Until COUNT1 > 12
      CROP_AMOUNT(count, COUNT1) = 0
      COUNT1 = COUNT1 + 1
   Loop
   count = count + 1
Loop
   
'set the weather for each turn
Set GAMES_WEATHER = TVDBGM.OpenRecordset("GAMES_WEATHER")
GAMES_WEATHER.index = "WEATHER_ZONE"
GAMES_WEATHER.MoveFirst
GAMES_WEATHER.Seek "=", HEXTABLE![WEATHER_ZONE]
  
count = 1
Do Until count > 12
   FARM_TURN(count) = GAMES_WEATHER![TURN]
   count = count + 1
   GAMES_WEATHER.MoveNext
   If Not GAMES_WEATHER![WEATHER_ZONE] = HEXTABLE![WEATHER_ZONE] Then
      Exit Do
   End If
   If GAMES_WEATHER.EOF Then
      Exit Do
   End If
Loop
GAMES_WEATHER.Close
  
   
' DO FARMING
Set FARMTABLE = TVDBGM.OpenRecordset("HEXMAP_FARMING")
FARMTABLE.index = "TRIBE"
If FARMTABLE.BOF Then
   ' do nothing
Else
   FARMTABLE.MoveFirst
End If
  
FARMTABLE.Seek "=", TRIBEINFO![CURRENT HEX], CLANNUMBER, TRIBENUMBER
VALID_FARMING_TURN = "N"

If FARMTABLE.NoMatch Then
   ' no crops to identify
Else
   Do Until Not FARMTABLE![HEXMAP] = TRIBEINFO![CURRENT HEX]
      If FARMTABLE![CLAN] = CLANNUMBER And FARMTABLE![TRIBE] = TRIBENUMBER Then
         If FARMTABLE![ITEM_NUMBER] >= 0 Then
            ' FIND CROP
            count = 1
            Do
              If FARMTABLE![ITEM] = "PLOWED" Then
                 GoTo FARMTABLE_LOOP
              End If
              If CROP(count) = "EMPTY" Then
                 CROP(count) = FARMTABLE![ITEM]
                 Exit Do
              End If
              If CROP(count) = FARMTABLE![ITEM] Then
                 Exit Do
              End If
              count = count + 1
              If count > 20 Then
                 GoTo FARMTABLE_LOOP
              End If
            Loop
            CROP_AMOUNT(count, Val(Left(FARMTABLE![TURN], 2))) = FARMTABLE![ITEM_NUMBER]
            CROP_FOUND(count) = "YES"
            VALID_FARMING_TURN = "Y"
         End If
      End If
FARMTABLE_LOOP:
         FARMTABLE.MoveNext
         If FARMTABLE.EOF Then
            Exit Do
         End If
     Loop
     
     wrdApp.Selection.TypeText vbCrLf & "Turn"
     count = 1
     Do Until count > 12
        If FARM_TURN(count) = "EMPTY" Or count = 12 Then
            wrdApp.Selection.TypeText vbCr
            Exit Do
        Else
             wrdApp.Selection.TypeText vbTab & FARM_TURN(count)
             count = count + 1
        End If
     Loop
        
     MAX_COUNT = count - 1
     ' LOOP THROUGH THE CROPS
     FIRST_COUNT = 1
     OutCount = 2
     Do
        If CROP_FOUND(FIRST_COUNT) = "YES" Then
           wrdApp.Selection.TypeText CROP(FIRST_COUNT)
           ' loop through the amount
           SECOND_COUNT = 1
           Do
               wrdApp.Selection.TypeText vbTab & CROP_AMOUNT(FIRST_COUNT, SECOND_COUNT)
               SECOND_COUNT = SECOND_COUNT + 1
               If SECOND_COUNT > MAX_COUNT Then
                  wrdApp.Selection.TypeText vbCr
                  Exit Do
               End If
               If SECOND_COUNT > 12 Then
                  wrdApp.Selection.TypeText vbCr
                  Exit Do
               End If
           Loop
        End If
        FIRST_COUNT = FIRST_COUNT + 1
        If FIRST_COUNT > 20 Then
           wrdApp.Selection.TypeText vbCr
           Exit Do
        End If
     Loop

  End If

  Set PermFarmingTable = TVDBGM.OpenRecordset("HEXMAP_PERMANENT_FARMING")
  PermFarmingTable.index = "TRIBE"
  If PermFarmingTable.BOF Then
      ' do nothing
  Else
      PermFarmingTable.MoveFirst
  End If
  
  PermFarmingTable.Seek "=", TRIBEINFO![CURRENT HEX], CLANNUMBER, TRIBENUMBER
  If Not PermFarmingTable.NoMatch Then
     wrdApp.Selection.TypeText "Permanent Crops : "
      Do Until Not PermFarmingTable![HEXMAP] = TRIBEINFO![CURRENT HEX]
             If PermFarmingTable![ITEM_NUMBER] > 0 Then
                 wrdApp.Selection.TypeText PermFarmingTable![ITEM] & "   " & vbTab
                 wrdApp.Selection.TypeText PermFarmingTable![ITEM_NUMBER] & "   " & vbTab
             End If
             PermFarmingTable.MoveNext
             If PermFarmingTable.EOF Then
                 Exit Do
             End If
      Loop
  End If
  
wrdApp.Selection.TypeText vbCr

TRIBES_IN_HEX = WHO_IS_IN_HEX(CLANNUMBER, TRIBENUMBER, TRIBEINFO![CURRENT HEX], "Y")

Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"

SECTION_NAME = "TRIBE STATUS"
Program_Area = "A300 - Status"

TRIBEINFO.MoveFirst
TRIBEINFO.Seek "=", CLANNUMBER, TRIBENUMBER

wrdApp.Selection.Font.Bold = True
wrdApp.Selection.TypeText TRIBEINFO![TRIBE] & " Status: "
wrdApp.Selection.Font.Bold = False

Set HEXTABLE = TVDBGM.OpenRecordset("HEX_MAP")
HEXTABLE.index = "PRIMARYKEY"
HEXTABLE.MoveFirst
HEXTABLE.Seek "=", TRIBEINFO![CURRENT HEX]
HEXMAPCITY.MoveFirst
HEXMAPCITY.Seek "=", TRIBEINFO![CURRENT HEX]
HEXMAPMINERALS.MoveFirst
HEXMAPMINERALS.Seek "=", TRIBEINFO![CURRENT HEX]
CURRENT_HEX_MAP = TRIBEINFO![CURRENT HEX]

If HEXTABLE.NoMatch Then
   wrdApp.Selection.TypeText " " & TRIBES_IN_HEX & vbCr
Else
   wrdApp.Selection.TypeText HEXTABLE![TERRAIN] & ","
   If Not HEXMAPCITY.NoMatch Then
      If Not IsNull(HEXMAPCITY![CITY]) Then
         wrdApp.Selection.TypeText " " & HEXMAPCITY![CITY] & ","
      End If
   End If
   
   If Not HEXMAPCITY.NoMatch Then
      If Not IsNull(HEXMAPCITY![CITY_2]) Then
         wrdApp.Selection.TypeText " " & HEXMAPCITY![CITY_2] & ","
      End If
   End If
   
   If Not HEXMAPMINERALS.NoMatch Then
      If Not IsNull(HEXMAPMINERALS![ORE_TYPE]) Then
         wrdApp.Selection.TypeText " " & HEXMAPMINERALS![ORE_TYPE] & ","
      End If
      Call Get_Research_Data(TCLANNUMBER, TRIBENUMBER, "Geologists")
      If RESEARCH_FOUND = "Y" Then
          If Not IsNull(HEXMAPMINERALS![SECOND_ORE]) Then
             wrdApp.Selection.TypeText " " & HEXMAPMINERALS![SECOND_ORE] & ","
          End If
          If Not IsNull(HEXMAPMINERALS![THIRD_ORE]) Then
             wrdApp.Selection.TypeText " " & HEXMAPMINERALS![THIRD_ORE] & ","
          End If
          If Not IsNull(HEXMAPMINERALS![FORTH_ORE]) Then
             wrdApp.Selection.TypeText " " & HEXMAPMINERALS![FORTH_ORE] & ","
          End If
      End If
   End If
   
   TERRAIN = ""
   GET_SURROUNDING_DATA (CURRENT_HEX_MAP)
   Call GET_MOUNTAINS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_OCEANS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_LAKES(TERRAIN, CURRENT_HEX_MAP)
   Call GET_PASSES(TERRAIN, CURRENT_HEX_MAP)
   Call GET_RIVERS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_FORDS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_BEACHS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_CLIFFS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_ROADS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_CANALS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_CANYONS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_STREAMS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_WATERFALLS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_QUARRIES(TERRAIN, CURRENT_HEX_MAP)
   Call GET_SPRINGS(TERRAIN, CURRENT_HEX_MAP)
   Call GET_SALMON_RUN(TERRAIN, CURRENT_HEX_MAP)
   Call GET_WHALING_AREA(TERRAIN, CURRENT_HEX_MAP)
   Call GET_FISH_AREA(TERRAIN, CURRENT_HEX_MAP)
  
   wrdApp.Selection.TypeText TERRAIN
   wrdApp.Selection.TypeText " " & TRIBES_IN_HEX & vbCr
End If

If TRIBEINFO![ABSORBED] = "Y" Then
   wrdApp.Selection.TypeText "Group has been Absorbed"
End If

    wrdApp.Selection.TypeText vbCrLf

'*****
'***** Start of category reporting
'*****

SECTION_NAME = "CONSTRUCTION"
Program_Area = "A300 - Construction"

Set VALIDBUILDINGS = TVDB.OpenRecordset("VALID_BUILDINGS")
VALIDBUILDINGS.index = "PRIMARYKEY"
VALIDBUILDINGS.MoveFirst

Set HEXCONSTTABLE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
HEXCONSTTABLE.index = "TRIBECONST"
HEXCONSTTABLE.MoveFirst

Do While HEXCONSTTABLE![TRIBE] <> TRIBENUMBER
   HEXCONSTTABLE.MoveNext
   If HEXCONSTTABLE.EOF Then
      Exit Do
   End If
Loop

First_Const = "Yes"

If Not HEXCONSTTABLE.EOF Then

   
Do While HEXCONSTTABLE![TRIBE] = TRIBENUMBER
   If First_Const = "Yes" Then
   
      wrdApp.Selection.Font.Color = vbBlue
      wrdApp.Selection.TypeText "Buildings:" & vbCr
      wrdApp.Selection.Font.Color = vbBlack
      First_Const = "No"
   End If
   VALIDBUILDINGS.MoveFirst
   VALIDBUILDINGS.Seek "=", HEXCONSTTABLE![CONSTRUCTION]
    
   If Not VALIDBUILDINGS.NoMatch Then
   'isContainerBuilding(CONSTRUCTION)
        If VALIDBUILDINGS![LIMITS] >= 10 Then
            numConstructionPrintingTrashold = -1
        Else
            numConstructionPrintingTrashold = 0
        End If
      If Not HEXCONSTTABLE![CONSTRUCTION] = "MONTHS TP OPEN" Then
         wrdApp.Selection.TypeText " " & VALIDBUILDINGS![SHORTNAME] & " " & HEXCONSTTABLE![1]
         If HEXCONSTTABLE![2] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![2]
         End If
         If HEXCONSTTABLE![3] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![3]
         End If
         If HEXCONSTTABLE![4] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![4]
         End If
         If HEXCONSTTABLE![5] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![5]
         End If
         If HEXCONSTTABLE![6] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![6]
         End If
         If HEXCONSTTABLE![7] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![7]
         End If
         If HEXCONSTTABLE![8] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![8]
         End If
         If HEXCONSTTABLE![9] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![9]
         End If
         If HEXCONSTTABLE![10] > numConstructionPrintingTrashold Then
            wrdApp.Selection.TypeText "-" & HEXCONSTTABLE![10]
         End If
         wrdApp.Selection.TypeText ", "
    End If
   End If
   
   HEXCONSTTABLE.MoveNext
   If HEXCONSTTABLE.EOF Then
      Exit Do
   End If
   If Not HEXCONSTTABLE![TRIBE] = TRIBENUMBER Then
      wrdApp.Selection.TypeText vbCrLf
      Exit Do
   End If
Loop
End If

'blank line after buildings
wrdApp.Selection.TypeText vbCrLf

OutLine = "EMPTY"

Program_Area = "A300 - People"
SECTION_NAME = "TOTAL PEOPLE"
wrdApp.Selection.Font.Color = vbBlue
wrdApp.Selection.TypeText "Humans" & vbCr
wrdApp.Selection.Font.Color = vbBlack

Call TABS_REQUIRED(SECTION_NAME)

Set TRIBEINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.MoveFirst
TRIBEINFO.Seek "=", CLANNUMBER, TRIBENUMBER

TotalPeople = TRIBEINFO![WARRIORS] + TRIBEINFO![ACTIVES] + TRIBEINFO![INACTIVES]

wrdApp.Selection.TypeText "People" & vbTab & TotalPeople & vbTab
wrdApp.Selection.TypeText vbCrLf
wrdApp.Selection.TypeText "Warriors " & vbTab & TRIBEINFO![WARRIORS] & vbTab
wrdApp.Selection.TypeText "Actives " & vbTab & TRIBEINFO![ACTIVES] & vbTab
wrdApp.Selection.TypeText "Inactives " & vbTab & TRIBEINFO![INACTIVES] & vbTab
If TRIBEINFO![SLAVE] > 0 Then
      wrdApp.Selection.TypeText "Slaves " & vbTab & TRIBEINFO![SLAVE] & vbTab
End If
If TRIBEINFO![HIRELINGS] > 0 Then
      wrdApp.Selection.TypeText "Hirelings " & vbTab & TRIBEINFO![HIRELINGS] & vbTab
End If
If TRIBEINFO![LOCALS] > 0 Then
      wrdApp.Selection.TypeText "Locals " & vbTab & TRIBEINFO![LOCALS] & vbTab
End If
If TRIBEINFO![Auxiliaries] > 0 Then
      wrdApp.Selection.TypeText "Auxiliaries " & vbTab & TRIBEINFO![Auxiliaries] & vbTab
End If
If TRIBEINFO![MERCENARIES] > 0 Then
      wrdApp.Selection.TypeText "Mercenaries " & vbTab & TRIBEINFO![MERCENARIES]
End If
wrdApp.Selection.TypeText vbCrLf & vbCrLf

' PRINT OUT SPECIALISTS
SECTION_NAME = "SPECIALISTS"
Call TABS_REQUIRED(SECTION_NAME)

Set TribesSpecialists = TVDBGM.OpenRecordset("Tribes_Specialists")
TribesSpecialists.index = "SECONDARYKEY"
TribesSpecialists.MoveFirst
TribesSpecialists.Seek "=", CLANNUMBER, TRIBENUMBER

If Not TribesSpecialists.NoMatch Then
    wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Specialists: "
   wrdApp.Selection.Font.Color = vbBlack
   wrdApp.Selection.TypeText vbCrLf
   
   Do Until TribesSpecialists.EOF
      If TribesSpecialists![TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText vbTab & TribesSpecialists![ITEM]
         If TribesSpecialists![ITEM] = "TRAINING" Then
            wrdApp.Selection.TypeText vbTab & TribesSpecialists![SPECIALISTS] & " ( Turn " & TribesSpecialists![NUMBER_OF_TURNS_TRAINING] & " of 3 ) " & vbCr
         Else
            wrdApp.Selection.TypeText vbTab & vbTab & TribesSpecialists![SPECIALISTS] & vbCr
         End If
      End If
     
      TribesSpecialists.MoveNext
   Loop
End If

'blank line after specialists
wrdApp.Selection.TypeText vbCrLf

OutLine = "EMPTY"

SECTION_NAME = "ANIMALS"
Program_Area = "A300 - Animals"

Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "PRIMARYKEY"
VALIDGOODS.MoveFirst

NumGoods = 0
LineI = 1
SECTION_NAME = "ANIMALS"
If IsNull(TRIBEINFO![GOODS TRIBE]) Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Animals" & vbCr
ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Animals" & vbCr
End If
Call TABS_REQUIRED(SECTION_NAME)
wrdApp.Selection.Font.Color = vbBlack

Set Tribes_Goods = TVDBGM.OpenRecordset("Tribes_Goods")
Tribes_Goods.index = "PRIMARYKEY"
Tribes_Goods.MoveFirst
Tribes_Goods.Seek ">=", CLANNUMBER, TRIBENUMBER, "ANIMAL", "A"

If Not Tribes_Goods.EOF Then
   If Not Tribes_Goods.NoMatch Then
      If Tribes_Goods![TRIBE] = TRIBENUMBER Then
         If Tribes_Goods![ITEM_TYPE] = "ANIMAL" Then
            'do nothing
         Else
            wrdApp.Selection.TypeText "None" & vbCr
         End If
      ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "None" & vbCr
      ElseIf IsNull(TRIBEINFO![GOODS TRIBE]) Then
         wrdApp.Selection.TypeText "None" & vbCr
      Else
         'do nothing
      End If
      
      Do While Tribes_Goods![TRIBE] = TRIBENUMBER And Tribes_Goods![ITEM_TYPE] = "ANIMAL"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDGOODS.NoMatch Then
            wrdApp.Selection.TypeText VALIDGOODS![SHORTNAME] & " " & vbTab
            wrdApp.Selection.TypeText CStr(Tribes_Goods![ITEM_NUMBER]) & vbTab
            NumGoods = NumGoods + 1
            Call CHECK_NUMGOODS
         End If
         Tribes_Goods.MoveNext
         If Tribes_Goods.EOF Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![TRIBE] = TRIBENUMBER Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![ITEM_TYPE] = "ANIMAL" Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
      Loop
   Else
      wrdApp.Selection.TypeText "None" & vbCr
   End If
End If


SECTION_NAME = "MINERALS"
Program_Area = "A300 - Minerals"
If IsNull(TRIBEINFO![GOODS TRIBE]) Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Minerals" & vbCr
ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Minerals" & vbCr
End If
Call TABS_REQUIRED(SECTION_NAME)
wrdApp.Selection.Font.Color = vbBlack

NumGoods = 0
LineI = 1

Tribes_Goods.MoveFirst
Tribes_Goods.Seek ">=", CLANNUMBER, TRIBENUMBER, "MINERAL", "A"
   
If Not Tribes_Goods.EOF Then
   If Not Tribes_Goods.NoMatch Then
      If Tribes_Goods![TRIBE] = TRIBENUMBER Then
         If Tribes_Goods![ITEM_TYPE] = "MINERAL" Then
            'do nothing
         Else
            wrdApp.Selection.TypeText "None" & vbCr
         End If
      ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "None" & vbCr
      ElseIf IsNull(TRIBEINFO![GOODS TRIBE]) Then
         wrdApp.Selection.TypeText "None" & vbCr
      Else
         'do nothing
      End If
      
      Do While Tribes_Goods![TRIBE] = TRIBENUMBER And Tribes_Goods![ITEM_TYPE] = "MINERAL"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDGOODS.NoMatch Then
            wrdApp.Selection.TypeText VALIDGOODS![SHORTNAME] & " " & vbTab
            wrdApp.Selection.TypeText CStr(Tribes_Goods![ITEM_NUMBER]) & vbTab
            NumGoods = NumGoods + 1
            Call CHECK_NUMGOODS
         End If
         Tribes_Goods.MoveNext
         If Tribes_Goods.EOF Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![TRIBE] = TRIBENUMBER Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![ITEM_TYPE] = "MINERAL" Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
      Loop
   Else
      wrdApp.Selection.TypeText "None" & vbCr
   End If
End If


Program_Area = "A300 - War Equipment"

NumGoods = 0
LineI = 1
SECTION_NAME = "WAR"
If IsNull(TRIBEINFO![GOODS TRIBE]) Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "War Equipment" & vbCr
ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "War Equipment" & vbCr
End If
Call TABS_REQUIRED(SECTION_NAME)
wrdApp.Selection.Font.Color = vbBlack

Tribes_Goods.MoveFirst
Tribes_Goods.Seek ">=", CLANNUMBER, TRIBENUMBER, "WAR", "A"
   
If Not Tribes_Goods.EOF Then
   If Not Tribes_Goods.NoMatch Then
      If Tribes_Goods![TRIBE] = TRIBENUMBER Then
         If Tribes_Goods![ITEM_TYPE] = "WAR" Then
            'do nothing
         Else
            wrdApp.Selection.TypeText "None" & vbCr
         End If
      ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "None" & vbCr
      ElseIf IsNull(TRIBEINFO![GOODS TRIBE]) Then
         wrdApp.Selection.TypeText "None" & vbCr
      Else
         'do nothing
      End If
      
      Do While Tribes_Goods![TRIBE] = TRIBENUMBER And Tribes_Goods![ITEM_TYPE] = "WAR"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDGOODS.NoMatch Then
            wrdApp.Selection.TypeText VALIDGOODS![SHORTNAME] & " " & vbTab
            wrdApp.Selection.TypeText CStr(Tribes_Goods![ITEM_NUMBER]) & vbTab
            NumGoods = NumGoods + 1
            Call CHECK_NUMGOODS
         End If
         Tribes_Goods.MoveNext
         If Tribes_Goods.EOF Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![TRIBE] = TRIBENUMBER Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![ITEM_TYPE] = "WAR" Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
      Loop
   Else
      wrdApp.Selection.TypeText "None" & vbCr
  End If
End If


Program_Area = "A300 - Finished Goods"
SECTION_NAME = "FINISHED"
If IsNull(TRIBEINFO![GOODS TRIBE]) Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Finished Goods" & vbCr
ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Finished Goods" & vbCr
End If
Call TABS_REQUIRED(SECTION_NAME)
wrdApp.Selection.Font.Color = vbBlack


NumGoods = 0
LineI = 1
Tribes_Goods.MoveFirst
Tribes_Goods.Seek ">=", CLANNUMBER, TRIBENUMBER, "FINISHED", "A"
   
If Not Tribes_Goods.EOF Then
   If Not Tribes_Goods.NoMatch Then
      If Tribes_Goods![TRIBE] = TRIBENUMBER Then
         If Tribes_Goods![ITEM_TYPE] = "FINISHED" Then
            'do nothing
         Else
            wrdApp.Selection.TypeText "None" & vbCr
         End If
      ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "None" & vbCr
      ElseIf IsNull(TRIBEINFO![GOODS TRIBE]) Then
         wrdApp.Selection.TypeText "None" & vbCr
      Else
         'do nothing
      End If
      
      Do While Tribes_Goods![TRIBE] = TRIBENUMBER And Tribes_Goods![ITEM_TYPE] = "FINISHED"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDGOODS.NoMatch Then
            wrdApp.Selection.TypeText VALIDGOODS![SHORTNAME] & " " & vbTab
            wrdApp.Selection.TypeText CStr(Tribes_Goods![ITEM_NUMBER]) & vbTab
            NumGoods = NumGoods + 1
            Call CHECK_NUMGOODS
         End If
         Tribes_Goods.MoveNext
         If Tribes_Goods.EOF Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![TRIBE] = TRIBENUMBER Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![ITEM_TYPE] = "FINISHED" Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
      Loop
   Else
      wrdApp.Selection.TypeText "None" & vbCr
   End If
End If


Program_Area = "A300 - Raw Materials"
SECTION_NAME = "RAW"
If IsNull(TRIBEINFO![GOODS TRIBE]) Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Raw Materials" & vbCr
ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Raw Materials" & vbCr
End If
Call TABS_REQUIRED(SECTION_NAME)
wrdApp.Selection.Font.Color = vbBlack

NumGoods = 0
LineI = 1

Tribes_Goods.MoveFirst
Tribes_Goods.Seek ">=", CLANNUMBER, TRIBENUMBER, "RAW", "A"
   
If Not Tribes_Goods.EOF Then
   If Not Tribes_Goods.NoMatch Then
      If Tribes_Goods![TRIBE] = TRIBENUMBER Then
         If Tribes_Goods![ITEM_TYPE] = "RAW" Then
            'do nothing
         Else
            wrdApp.Selection.TypeText "None" & vbCr
         End If
      ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "None" & vbCr
      ElseIf IsNull(TRIBEINFO![GOODS TRIBE]) Then
         wrdApp.Selection.TypeText "None" & vbCr
      Else
         'do nothing
      End If
      
      Do While Tribes_Goods![TRIBE] = TRIBENUMBER And Tribes_Goods![ITEM_TYPE] = "RAW"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDGOODS.NoMatch Then
            wrdApp.Selection.TypeText VALIDGOODS![SHORTNAME] & " " & vbTab
            wrdApp.Selection.TypeText CStr(Tribes_Goods![ITEM_NUMBER]) & vbTab
            NumGoods = NumGoods + 1
            Call CHECK_NUMGOODS
         End If
         Tribes_Goods.MoveNext
         If Tribes_Goods.EOF Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![TRIBE] = TRIBENUMBER Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![ITEM_TYPE] = "RAW" Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
      Loop
   Else
      wrdApp.Selection.TypeText "None" & vbCr
   End If
End If


Program_Area = "A300 - Ships"
SECTION_NAME = "SHIPS"
DR_Count = 0
Boatshed_Req = 0
Min_Crew_Req = 0
Max_Crew = 0
Max_Cargo = 0
Ship_Found = "NO"

If IsNull(TRIBEINFO![GOODS TRIBE]) Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Ships" & vbCr
ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Ships" & vbCr
End If
Call TABS_REQUIRED(SECTION_NAME)
wrdApp.Selection.Font.Color = vbBlack

LineI = 1

Tribes_Goods.MoveFirst
Tribes_Goods.Seek ">=", CLANNUMBER, TRIBENUMBER, "SHIP", "A"

If Not Tribes_Goods.EOF Then
   If Not Tribes_Goods.NoMatch Then
      If Tribes_Goods![TRIBE] = TRIBENUMBER Then
         If Tribes_Goods![ITEM_TYPE] = "SHIP" Then
            'do nothing
         Else
            wrdApp.Selection.TypeText "None" & vbCr
         End If
      ElseIf TRIBEINFO![GOODS TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "None" & vbCr
      ElseIf IsNull(TRIBEINFO![GOODS TRIBE]) Then
         wrdApp.Selection.TypeText "None" & vbCr
      Else
         'do nothing
      End If
      
      Do While Tribes_Goods![TRIBE] = TRIBENUMBER And Tribes_Goods![ITEM_TYPE] = "SHIP"
         Ship_Found = "YES"
         VALIDGOODS.MoveFirst
         VALIDGOODS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDGOODS.NoMatch Then
            wrdApp.Selection.TypeText VALIDGOODS![SHORTNAME] & " " & vbTab
            wrdApp.Selection.TypeText CStr(Tribes_Goods![ITEM_NUMBER]) & vbTab
            NumGoods = NumGoods + 1
            Call CHECK_NUMGOODS
         End If
         
         VALIDSHIPS.MoveFirst
         VALIDSHIPS.Seek "=", Tribes_Goods![ITEM]
         If Not VALIDSHIPS.NoMatch Then
            DR_Count = DR_Count + (VALIDSHIPS![DR_Sail] * Tribes_Goods![ITEM_NUMBER])
            DR_Count = DR_Count + (VALIDSHIPS![DR_Hull] * Tribes_Goods![ITEM_NUMBER])
            Min_Crew_Req = Min_Crew_Req + (VALIDSHIPS![Crew_Required] * Tribes_Goods![ITEM_NUMBER])
            Max_Crew = Max_Crew + (VALIDSHIPS![Max_Crew] * Tribes_Goods![ITEM_NUMBER])
            Max_Cargo = Max_Cargo + (VALIDSHIPS![Cargo_Space] * Tribes_Goods![ITEM_NUMBER])
         End If
         
         Tribes_Goods.MoveNext
         If Tribes_Goods.EOF Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![TRIBE] = TRIBENUMBER Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
         If Not Tribes_Goods![ITEM_TYPE] = "SHIP" Then
            If NumGoods > 0 Then
               wrdApp.Selection.TypeText vbCr
            End If
            Exit Do
         End If
      Loop
   Else
      wrdApp.Selection.TypeText "None" & vbCr
   End If
End If
      
If Ship_Found = "YES" Then
   Boatshed_Req = DR_Count / 10
    wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Shipping fleet requires: " & vbCr
   wrdApp.Selection.Font.Color = vbBlack
   wrdApp.Selection.TypeText vbTab & vbTab & "boatsheds required = " & Boatshed_Req & vbCr
   wrdApp.Selection.TypeText vbTab & vbTab & "min crew required = " & Min_Crew_Req & vbCr
   wrdApp.Selection.TypeText vbTab & vbTab & "max people available (including crew) = " & Max_Crew & vbCr
   wrdApp.Selection.TypeText vbTab & vbTab & "max cargo available (with max people) = " & Max_Cargo & vbCr
End If

wrdApp.Selection.TypeText vbCrLf


Program_Area = "A300 - Books"
SECTION_NAME = "BOOKS"
Call TABS_REQUIRED(SECTION_NAME)

Set TRIBESBOOKS = TVDBGM.OpenRecordset("Tribes_Books")
TRIBESBOOKS.index = "PRIMARYKEY"
If Not TRIBESBOOKS.EOF Then
   TRIBESBOOKS.MoveFirst
   Do Until TRIBESBOOKS![TRIBE] = TRIBENUMBER
      TRIBESBOOKS.MoveNext
      If TRIBESBOOKS.EOF Then
         Exit Do
      End If
   Loop
End If

If Not TRIBESBOOKS.EOF Then
   Do While TRIBESBOOKS![TRIBE] = TRIBENUMBER
      wrdApp.Selection.TypeText "Book : " & vbTab & TRIBESBOOKS![BOOK] & " " & vbTab
      wrdApp.Selection.TypeText "Copies : " & vbTab & TRIBESBOOKS![NUMBER]
      wrdApp.Selection.TypeText vbCr
      TRIBESBOOKS.MoveNext
      If TRIBESBOOKS.EOF Then
         Exit Do
      End If
      If Not TRIBESBOOKS![TRIBE] = TRIBENUMBER Then
         Exit Do
      End If
   Loop
End If

Program_Area = "A300 - Under Construction"
SECTION_NAME = "UNDER CONSTRUCTION"
Call TABS_REQUIRED(SECTION_NAME)

Set UNDERCONSTTABLE = TVDBGM.OpenRecordset("UNDER_CONSTRUCTION")
UNDERCONSTTABLE.index = "TRIBE"
UNDERCONSTTABLE.MoveFirst
UNDERCONSTTABLE.Seek "=", TRIBENUMBER

If Not UNDERCONSTTABLE.NoMatch Then
   Do Until UNDERCONSTTABLE.EOF
     If UNDERCONSTTABLE![TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText "Under Construction" & vbTab & "(" & UNDERCONSTTABLE![CONSTRUCTION] & " "
         If UNDERCONSTTABLE![LOGS] > 0 Then
            wrdApp.Selection.TypeText UNDERCONSTTABLE![LOGS] & " Logs"
         End If
         If UNDERCONSTTABLE![STONES] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![STONES] & " Stones"
         End If
         If UNDERCONSTTABLE![COAL] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![COAL] & " Coal"
         End If
         If UNDERCONSTTABLE![BRASS] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![BRASS] & " Brass"
         End If
         If UNDERCONSTTABLE![BRONZE] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![BRONZE] & " Bronze"
         End If
         If UNDERCONSTTABLE![COPPER] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![COPPER] & " Copper"
         End If
         If UNDERCONSTTABLE![IRON] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![IRON] & " Iron"
         End If
         If UNDERCONSTTABLE![LEAD] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![LEAD] & " Lead"
         End If
         If UNDERCONSTTABLE![CLOTH] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![CLOTH] & " Cloth"
         End If
         If UNDERCONSTTABLE![LEATHER] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![LEATHER] & " Leather"
         End If
         If UNDERCONSTTABLE![ROPES] > 0 Then
            wrdApp.Selection.TypeText ", " & UNDERCONSTTABLE![ROPES] & " Ropes"
         End If
         wrdApp.Selection.TypeText ")" & vbCr
      Else
         wrdApp.Selection.TypeText vbCrLf
         Exit Do
      End If
     
      UNDERCONSTTABLE.MoveNext
   Loop
End If

Program_Area = "A300 - Skilled"
SECTION_NAME = "SKILLS"
NONSKILLED = 0

'Check for Element
String_Length = Len(TRIBENUMBER)
If String_Length > 4 Then
   NONSKILLED = 0
End If

If NONSKILLED = 0 Then
    
   wrdApp.Selection.Font.Color = vbBlue
   wrdApp.Selection.TypeText "Skills: " & vbCr
   Call TABS_REQUIRED(SECTION_NAME)
   wrdApp.Selection.Font.Color = vbBlack

   
   Call Get_Skills

   If TRIBEINFO![MORALE] > 0 Then
      wrdApp.Selection.Font.Bold = True
      wrdApp.Selection.TypeText "Morale : "
      wrdApp.Selection.Font.Bold = False
      wrdApp.Selection.TypeText vbTab & TRIBEINFO![MORALE] & vbCrLf & vbCrLf
      OutLine = "EMPTY"
   End If

End If

SECTION_NAME = "CAPACITY"
Call TABS_REQUIRED(SECTION_NAME)

   ' Weight is generated first to ensure that all weights are output, not just those belonging
   ' to tribes
   ' this will actually be printed after morale
   wrdApp.Selection.Font.Bold = True
   wrdApp.Selection.TypeText "Weight: "
   wrdApp.Selection.Font.Bold = False
'   wrdApp.Selection.TypeText vbTab & vbTab & vbTab & _
'                            Format(Nz(TRIBEINFO![WEIGHT], 0), "###,##0") & vbCr
                            
    wrdApp.Selection.TypeText Format(Nz(TRIBEINFO![WEIGHT], 0), "###,##0") & vbCr
                            
    'Carrying Capacity removed till code done
   wrdApp.Selection.Font.Bold = True
   wrdApp.Selection.TypeText "Walking CC: "
   wrdApp.Selection.Font.Bold = False
   wrdApp.Selection.TypeText vbTab & vbTab & vbTab & _
                            Format(Nz(TRIBEINFO![Walking_Capacity], 0), "###,##0") & vbCr
   wrdApp.Selection.Font.Bold = True
   ' if mounted capacity is zero then say so
   wrdApp.Selection.TypeText "Mounted CC: "
   wrdApp.Selection.Font.Bold = False
   wrdApp.Selection.TypeText vbTab & vbTab & vbTab & _
                            Format(Nz(TRIBEINFO![CAPACITY], 0), "###,##0") & vbCr

'   ' add in goods tribe

'    wrdApp.Selection.Font.Bold = True
'    wrdApp.Selection.TypeText "GT Walking CC: "
'    wrdApp.Selection.Font.Bold = False
'    wrdApp.Selection.TypeText vbTab & vbTab & vbTab & _
'                            Format(Nz(TRIBEINFO![GT_WALKING_CAPACITY], 0), "###,##0") & vbCr
'    wrdApp.Selection.Font.Bold = True
'    wrdApp.Selection.TypeText "GT Mounted CC: "
'    wrdApp.Selection.Font.Bold = False
'    wrdApp.Selection.TypeText vbTab & vbTab & vbTab & _
'                            Format(Nz(TRIBEINFO![GT_MOUNTED_CAPACITY], 0), "###,##0") & vbCr



SECTION_NAME = "POLITICS"
Call TABS_REQUIRED(SECTION_NAME)
Call Perform_Pacification_Printing

SkillsTab.Close

If Not IsNull(TRIBEINFO![TERRAIN PROFS]) Then
   wrdApp.Selection.Font.Bold = True
   wrdApp.Selection.TypeText "Terrain Profs : "
   wrdApp.Selection.Font.Bold = False
   wrdApp.Selection.TypeText vbTab & TRIBEINFO![TERRAIN PROFS] & vbCrLf
End If

If Not IsNull(TRIBEINFO![RELIGION]) Then
   If Not TRIBEINFO![RELIGION] = "" Then
      Set RELIGIONTABLE = TVDB.OpenRecordset("RELIGION")
      RELIGIONTABLE.index = "PRIMARYKEY"
      RELIGIONTABLE.MoveFirst
      RELIGIONTABLE.Seek "=", TRIBEINFO![RELIGION]
      If Not RELIGIONTABLE.NoMatch Then
         wrdApp.Selection.Font.Bold = True
         wrdApp.Selection.TypeText "Religion : "
         wrdApp.Selection.Font.Bold = False
         wrdApp.Selection.TypeText vbTab & TRIBEINFO![RELIGION] & ", "
         wrdApp.Selection.TypeText RELIGIONTABLE![MEMBERS] & vbCrLf
      Else
         MSG1 = "Chief, Religion " & TRIBEINFO![RELIGION]
         MSG2 = "has no members, woops"
         MsgBox (MSG1 & MSG2)
         wrdApp.Selection.TypeText vbCrLf
      End If
   End If
End If

Program_Area = "A300 - Research"
SECTION_NAME = "RESEARCH"
wrdApp.Selection.TypeText vbCrLf
Call TABS_REQUIRED(SECTION_NAME)

LineI = 1
Set RESEARCHTABLE = TVDBGM.OpenRecordset("TRIBE_RESEARCH")
RESEARCHTABLE.index = "SECONDARYKEY"
If Not RESEARCHTABLE.EOF Then
   RESEARCHTABLE.MoveFirst
End If

RESEARCHTABLE.Seek "=", TRIBENUMBER

count = 1

If Not RESEARCHTABLE.NoMatch Then
   wrdApp.Selection.Font.Bold = True
   wrdApp.Selection.TypeText "Research : "
   wrdApp.Selection.Font.Bold = False
   wrdApp.Selection.TypeText vbTab
   Do Until Not RESEARCHTABLE![TRIBE] = TRIBENUMBER
      If count = 1 Or count = 2 Then
         If RESEARCHTABLE![RESEARCH ATTEMPTED] = "Y" Then
            If RESEARCHTABLE![RESEARCH ATTAINED] = "Y" Then
               wrdApp.Selection.Font.Bold = True
               wrdApp.Selection.Font.Color = vbGreen
               wrdApp.Selection.TypeText RESEARCHTABLE![TOPIC] & "  DL"
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL ATTAINED] & " of "
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL REQUIRED]
               If RESEARCHTABLE![Cost] > 0 Then
                  'wrdApp.Selection.TypeText " (Cost " & RESEARCHTABLE![COST] & " silver)"
               End If
               wrdApp.Selection.Font.Bold = False
               wrdApp.Selection.Font.Color = vbBlack
            Else
               wrdApp.Selection.Font.Bold = True
               wrdApp.Selection.Font.Color = vbRed
               wrdApp.Selection.TypeText RESEARCHTABLE![TOPIC] & "  DL"
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL ATTAINED] & " of "
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL REQUIRED]
               If RESEARCHTABLE![Cost] > 0 Then
                  'wrdApp.Selection.TypeText " (Cost " & RESEARCHTABLE![COST] & " silver)"
               End If
               wrdApp.Selection.Font.Bold = False
               wrdApp.Selection.Font.Color = vbBlack
            End If
         Else
            wrdApp.Selection.TypeText RESEARCHTABLE![TOPIC] & "  DL"
            wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL ATTAINED] & " of "
            wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL REQUIRED]
            If RESEARCHTABLE![Cost] > 0 Then
               'wrdApp.Selection.TypeText " (Cost " & RESEARCHTABLE![COST] & " silver)"
            End If
            wrdApp.Selection.Font.Bold = False
            wrdApp.Selection.Font.Color = vbBlack
         End If
      Else
         If RESEARCHTABLE![RESEARCH ATTEMPTED] = "Y" Then
            If RESEARCHTABLE![RESEARCH ATTAINED] = "Y" Then
               wrdApp.Selection.Font.Bold = True
               wrdApp.Selection.Font.Color = vbGreen
               wrdApp.Selection.TypeText RESEARCHTABLE![TOPIC] & "  DL"
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL ATTAINED] & " of "
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL REQUIRED]
               wrdApp.Selection.Font.Color = vbBlack
               If RESEARCHTABLE![Cost] > 0 Then
                  'wrdApp.Selection.TypeText " (Cost " & RESEARCHTABLE![COST] & " silver)"
               End If
               wrdApp.Selection.Font.Bold = False
            Else
               wrdApp.Selection.Font.Bold = True
               wrdApp.Selection.Font.Color = vbRed
               wrdApp.Selection.TypeText RESEARCHTABLE![TOPIC] & "  DL"
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL ATTAINED] & " of "
               wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL REQUIRED]
               If RESEARCHTABLE![Cost] > 0 Then
                  'wrdApp.Selection.TypeText " (Cost " & RESEARCHTABLE![COST] & " silver)"
               End If
               wrdApp.Selection.Font.Bold = False
               wrdApp.Selection.Font.Color = vbBlack
            End If
         Else
            wrdApp.Selection.TypeText RESEARCHTABLE![TOPIC] & "  DL"
            wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL ATTAINED] & " of "
            wrdApp.Selection.TypeText RESEARCHTABLE![DL LEVEL REQUIRED]
            If RESEARCHTABLE![Cost] > 0 Then
               'wrdApp.Selection.TypeText " (Cost " & RESEARCHTABLE![COST] & " silver)"
            End If
            wrdApp.Selection.Font.Bold = False
            wrdApp.Selection.Font.Color = vbBlack
         End If
      End If
      If count = 3 Then
         wrdApp.Selection.TypeText vbCr
         count = 1
      ElseIf count = 2 Then
         wrdApp.Selection.TypeText vbTab
         count = 3
      Else
         wrdApp.Selection.TypeText vbTab
         count = 2
      End If
      RESEARCHTABLE.MoveNext
      If RESEARCHTABLE.EOF Then
         wrdApp.Selection.TypeText vbCr
         Exit Do
      End If
      If Not RESEARCHTABLE![TRIBE] = TRIBENUMBER Then
         wrdApp.Selection.TypeText vbCr
         Exit Do
      End If
   Loop
End If

LineI = 1

SECTION_NAME = "RESEARCH ATTEMPTED"
   Call FIND_AND_PRINT_ACTIVITIES
   
wrdApp.Selection.TypeText vbCrLf

SECTION_NAME = "COMPLETED_RESEARCH"
Call TABS_REQUIRED(SECTION_NAME)

Set COMPRESTABLE = TVDBGM.OpenRecordset("COMPLETED_RESEARCH")
COMPRESTABLE.index = "TRIBE"
COMPRESTABLE.MoveFirst

COMPRESTABLE.Seek "=", TRIBENUMBER

times_through = 0
count = 1
If Not COMPRESTABLE.NoMatch Then
wrdApp.Selection.Font.Bold = True
wrdApp.Selection.TypeText "Comp. Research : "
wrdApp.Selection.Font.Bold = False
wrdApp.Selection.TypeText vbCrLf

   Do Until COMPRESTABLE.EOF
      If COMPRESTABLE![TRIBE] = TRIBENUMBER Then
         If count < 4 Then
           If COMPRESTABLE![COMPLETED_THIS_TURN] = "Y" Then
              wrdApp.Selection.Font.Bold = True
              wrdApp.Selection.TypeText COMPRESTABLE![TOPIC]
              wrdApp.Selection.Font.Bold = False
              wrdApp.Selection.TypeText ", "
           Else
              wrdApp.Selection.TypeText COMPRESTABLE![TOPIC] & vbTab
           End If
           count = count + 1
           times_through = times_through + 1
         Else
            If COMPRESTABLE![COMPLETED_THIS_TURN] = "Y" Then
               wrdApp.Selection.Font.Bold = True
               wrdApp.Selection.TypeText COMPRESTABLE![TOPIC]
               wrdApp.Selection.Font.Bold = False
               wrdApp.Selection.TypeText ", "
            Else
               wrdApp.Selection.TypeText COMPRESTABLE![TOPIC] & vbTab
            End If
            count = 1
            times_through = times_through + 1
         End If
         If times_through >= 4 Then
            wrdApp.Selection.TypeText vbCrLf
            times_through = 0
         End If
      Else
         Exit Do
      End If
      COMPRESTABLE.MoveNext
   Loop
End If
   
wrdApp.Selection.TypeText vbCrLf
wrdApp.Selection.TypeText vbCrLf

   
   If Not IsNull(TRIBEINFO![TRUCES]) Then
      wrdApp.Selection.Font.Bold = True
      wrdApp.Selection.TypeText "Truces : "
      wrdApp.Selection.Font.Bold = False
      wrdApp.Selection.TypeText TRIBEINFO![TRUCES] & vbCr
   Else
      wrdApp.Selection.TypeText vbCr
   End If
   
   
   If Not IsNull(TRIBEINFO![War]) Then
      wrdApp.Selection.Font.Bold = True
      wrdApp.Selection.TypeText "War : "
      wrdApp.Selection.Font.Bold = False
      wrdApp.Selection.TypeText TRIBEINFO![War] & vbCrLf & vbCrLf
      OutLine = "EMPTY"
   End If

   TRIBEINFO.MoveNext
   If TRIBEINFO.EOF Then
      Exit Do
   End If
 
Loop

ERR_A300_UPDATE_CLOSE:
   'DoCmd.Close A_FORM, "UPDATE_SINGLE_CLAN"
   'DoCmd.OpenForm "UPDATE_SINGLE_CLAN"
   Exit Sub


ERR_A300_UPDATE:
If (Err = 3021) Then
   Resume Next
   
Else
  MSG1 = "Error # " & Err & " " & Error$ & " "
  MSG2 = "Area = " & Program_Area & " "
  MSG3 = "Tribe = " & TRIBENUMBER & " "
  MsgBox (MSG1 & MSG2 & MSG3)
  Resume ERR_A300_UPDATE_CLOSE
  
End If


End Sub
Sub CHECK_NUMGOODS()
   If NumGoods = 6 Then
      NumGoods = 0
      wrdApp.Selection.TypeText vbCr
   End If
End Sub

Sub CLOSE_WORD()
' close word and back to access
wrdApp.Quit

Set wrdApp = Nothing

'AppActivate "Microsoft Access - [TRIBEVIBES : Form]"
Forms![PRINT_FROM_CLAN].SetFocus

End Sub

Sub Delete_Existing_Turn()
On Error GoTo SEND_ERROR

   fileName = CurDir$ & TVDirect & "\" & CLANNUMBER & ".docx"
   
   Kill fileName
   

FINISH_SUB:
   Exit Sub


SEND_ERROR:
If Not (Err = 53) Then
   MSG1 = "ERROR = " & Err
   MSG2 = "FILE NOT DELETED"
   Response = MsgBox(MSG1 & MSG2, True)
End If

Resume FINISH_SUB

End Sub



Sub Get_Skills()
Set ValidSkills = TVDB.OpenRecordset("Valid_Skills")
ValidSkills.index = "PRIMARYKEY"
ValidSkills.MoveFirst

Set SkillsTab = TVDBGM.OpenRecordset("Skills")
SkillsTab.index = "PRIMARYKEY"
SkillsTab.MoveFirst
SkillsTab.Seek "=", TRIBENUMBER, "AAAAA"

If SkillsTab.NoMatch Then
   SkillsTab.AddNew
   SkillsTab![TRIBE] = TRIBENUMBER
   SkillsTab![Skill] = "AAAAA"
   SkillsTab.UPDATE
   SkillsTab.Seek "=", TRIBENUMBER, "AAAAA"
End If

SkillsTab.MoveNext

With wrdApp.Selection

NUM_CHARS = NUM_CHARS + 6
POLITICAL_LEVEL = 0
SECTION_NAME = "SKILLS"

Do Until Not (SkillsTab![TRIBE] = TRIBENUMBER)

   ValidSkills.Seek "=", SkillsTab![Skill]
   If SkillsTab![Skill] = "POLITICS" Then
      POLITICAL_LEVEL = SkillsTab![SKILL LEVEL]
   End If

   If NUM_CHARS = 0 Then
      If SkillsTab![ATTEMPTED] = "Y" Then
         If SkillsTab![SUCCESSFUL] = "Y" Then
            wrdApp.Selection.Font.Bold = True
            wrdApp.Selection.Font.Italic = True
            wrdApp.Selection.Font.Underline = True
            wrdApp.Selection.Font.Color = vbGreen
            wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
            wrdApp.Selection.Font.Bold = False
            wrdApp.Selection.Font.Italic = False
            wrdApp.Selection.Font.Underline = False
            wrdApp.Selection.Font.Color = vbBlack
         Else
            wrdApp.Selection.Font.Underline = True
            wrdApp.Selection.Font.Color = vbRed
            wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " "
            wrdApp.Selection.TypeText SkillsTab![SKILL LEVEL] & ", "
            wrdApp.Selection.Font.Underline = False
            wrdApp.Selection.Font.Color = vbBlack
         End If
      ElseIf SkillsTab![SUCCESSFUL] = "Y" Then
         wrdApp.Selection.Font.Bold = True
         wrdApp.Selection.Font.Italic = True
         wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
         wrdApp.Selection.Font.Bold = False
         wrdApp.Selection.Font.Italic = False
      Else
         wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
      End If
   ElseIf SkillsTab![ATTEMPTED] = "Y" Then
       If SkillsTab![SUCCESSFUL] = "Y" Then
          wrdApp.Selection.Font.Bold = True
          wrdApp.Selection.Font.Italic = True
          wrdApp.Selection.Font.Underline = True
          wrdApp.Selection.Font.Color = vbGreen
          wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
          wrdApp.Selection.Font.Bold = False
          wrdApp.Selection.Font.Italic = False
          wrdApp.Selection.Font.Underline = False
          wrdApp.Selection.Font.Color = vbBlack
       Else
          wrdApp.Selection.Font.Underline = True
          wrdApp.Selection.Font.Color = vbRed
          wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
          wrdApp.Selection.Font.Underline = False
          wrdApp.Selection.Font.Color = vbBlack
       End If
   Else
       If SkillsTab![SUCCESSFUL] = "Y" Then
          wrdApp.Selection.Font.Bold = True
          wrdApp.Selection.Font.Italic = True
          wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
          wrdApp.Selection.Font.Bold = False
          wrdApp.Selection.Font.Italic = False
       Else
          wrdApp.Selection.TypeText ValidSkills![SHORTNAME] & " " & SkillsTab![SKILL LEVEL] & ", "
       End If
   End If
   NUM_CHARS = NUM_CHARS + Len(ValidSkills![SHORTNAME]) + 2

  SkillsTab.MoveNext

Loop
          
wrdApp.Selection.TypeText vbNewLine & vbNewLine

SkillsTab.Seek "=", TRIBENUMBER, "AAAAA"
SkillsTab.Delete
SkillsTab.Close

End With

End Sub



Sub OPEN_NEW_FILE()
On Error GoTo ERR_OPEN_NEW_FILE
'for word 2010 plus
    wrdApp.Visible = False
    Set wrdDoc = wrdApp.Documents.Add

    wrdApp.ActiveDocument.SaveAs DIRECTPATH & "\" & CLANNUMBER
    With wrdDoc
        .PageSetup.LeftMargin = wrdApp.CentimetersToPoints(1)
        .PageSetup.RightMargin = wrdApp.CentimetersToPoints(1)
        .PageSetup.TopMargin = wrdApp.CentimetersToPoints(1)
        .PageSetup.BottomMargin = wrdApp.CentimetersToPoints(1.5)
        .PageSetup.PaperSize = wdPaperA4
        .Range.ParagraphFormat.SpaceAfter = 0
        With .Styles(wdStyleHeading1).Font
             .Name = "Calibri"
             .Size = 12
             .Bold = True
             .Color = wdColorBlack
        End With
        
        With .Styles(wdStyleHeading2).Font
             .Name = "Calibri"
             .Size = 12
             .Bold = True
             .Color = wdColorBlack
        End With

        With .Styles(wdStyleNormal).Font
             .Name = "Calibri"
             .Size = 10
             .Bold = False
             .Color = wdColorBlack
        End With
        .Styles(wdStyleHeading1).ParagraphFormat.SpaceBefore = 0
        .Styles(wdStyleHeading1).ParagraphFormat.SpaceAfter = 0
        .Styles(wdStyleHeading2).ParagraphFormat.SpaceBefore = 0
        .Styles(wdStyleHeading2).ParagraphFormat.SpaceAfter = 0
        .Styles(wdStyleNormal).ParagraphFormat.SpaceBefore = 0
        .Styles(wdStyleNormal).ParagraphFormat.SpaceAfter = 0
    End With
    
    With wrdApp
        .Selection.Font.Name = "Calibri"
        .Selection.Font.Size = 10
        ' clear all tabstops
        .Selection.Paragraphs(1).TabStops.ClearAll
        ' add in tabstops
        .Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(1), Alignment:=wdAlignTabLeft
        .Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(2.5), Alignment:=wdAlignTabLeft
        .Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(3.5), Alignment:=wdAlignTabLeft
        .Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(4.5), Alignment:=wdAlignTabLeft
    End With
    With wrdDoc
        .Saved = True
    End With
    'wrdDoc.Activate
    
ERR_OPEN_NEW_FILE_CLOSE:
   Exit Sub


ERR_OPEN_NEW_FILE:
If (Err = 5356) Then
  Msg = "File Still Open for Clan " & CLANNUMBER & " Program Stopping"
  MsgBox (Msg)
  STOP_PROCESSING = "YES"
  Exit Sub
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_OPEN_NEW_FILE_CLOSE
  
End If
End Sub

Sub FIND_AND_PRINT_ACTIVITIES()
Dim Political_LF As String

Political_LF = "NO"

   Call TABS_REQUIRED(SECTION_NAME)
   OutPutTab.MoveFirst
   OutPutTab.Seek "=", CLANNUMBER, TRIBENUMBER, SECTION_NAME, 1

   If Not OutPutTab.NoMatch Then
      Do While OutPutTab![Section] = SECTION_NAME
         If Left(OutPutTab![line detail], 5) = "EMPTY" Then
            'do nothing
         ElseIf SECTION_NAME = "Activities" Then
             ' find Bold, set, extract, unset, next, rinse repeat
             If Mid(OutPutTab![line detail], 1, 18) = "^BFinal Activities" Then
                  'line feed at the start of politics - do one
                  wrdApp.Selection.TypeText vbCrLf
                  wrdApp.Selection.TypeText vbCrLf
             End If
             If Mid(OutPutTab![line detail], 1, 18) = "^BPolitical Tithes" Then
                If Political_LF = "NO" Then
                  'line feed at the start of politics - do one
                  wrdApp.Selection.TypeText vbCrLf
                  wrdApp.Selection.TypeText vbCrLf
                  Political_LF = "YES"
                End If
             End If
            
             NumChars = Len(OutPutTab![line detail])
             String_Found1 = 1
             Chars_Read = 0
             Do While NumChars > 0
                OutLine = Mid(OutPutTab![line detail], Chars_Read + 1, NumChars)
                If Mid(OutLine, 2, 1) = "^" Then
                   wrdApp.Selection.TypeText Left(OutLine, 1)
                   Chars_Read = Chars_Read + 1
                   NumChars = NumChars - 1
                ElseIf Mid(OutLine, 2, 1) = "+" Then
                   wrdApp.Selection.TypeText Left(OutLine, 1)
                   Chars_Read = Chars_Read + 1
                   NumChars = NumChars - 1
                ElseIf Left(OutLine, 2) = "^B" Then
                   If first_B = "No" Then
                      wrdApp.Selection.Font.Bold = True
                      first_B = "Yes"
                   Else
                      wrdApp.Selection.Font.Bold = False
                      first_B = "No"
                   End If
                   Chars_Read = Chars_Read + 2
                   NumChars = NumChars - 2
                ElseIf Left(OutLine, 2) = "+9" Then
                   wrdApp.Selection.TypeText "("
                   Chars_Read = Chars_Read + 2
                   NumChars = NumChars - 2
                ElseIf Left(OutLine, 2) = "+0" Then
                   wrdApp.Selection.TypeText ")"
                   Chars_Read = Chars_Read + 2
                   NumChars = NumChars - 2
                Else
                   wrdApp.Selection.TypeText Left(OutLine, 2)
                   Chars_Read = Chars_Read + 2
                   NumChars = NumChars - 2
                End If
                If NumChars < 0 Then
                   wrdApp.Selection.TypeText vbCr
                   Exit Do
                End If
                OutLine = Mid(OutPutTab![line detail], Chars_Read + 1, NumChars)
           Loop
         ElseIf SECTION_NAME = "Tribe Movement" Then
            Movement_found = "Yes"
            wrdApp.Selection.TypeText vbCrLf
            If Left(OutPutTab![line detail], 13) = "Tribe Follows" Then
               wrdApp.Selection.TypeText OutPutTab![line detail] & vbCr
            Else
               ' find Bold, set, extract, unset, next, rinse repeat
               NumChars = Len(OutPutTab![line detail])
               String_Found1 = 1
               Chars_Read = 0
               Do While NumChars > 0
                  OutLine = Mid(OutPutTab![line detail], Chars_Read + 1, NumChars)
                  If Mid(OutLine, 2, 1) = "^" Then
                     wrdApp.Selection.TypeText Left(OutLine, 1)
                     Chars_Read = Chars_Read + 1
                     NumChars = NumChars - 1
                  ElseIf Mid(OutLine, 2, 1) = "+" Then
                     wrdApp.Selection.TypeText Left(OutLine, 1)
                     Chars_Read = Chars_Read + 1
                     NumChars = NumChars - 1
                  ElseIf Left(OutLine, 2) = "^B" Then
                     If first_B = "No" Then
                        wrdApp.Selection.Font.Bold = True
                        first_B = "Yes"
                     Else
                        wrdApp.Selection.Font.Bold = False
                        first_B = "No"
                     End If
                     Chars_Read = Chars_Read + 2
                     NumChars = NumChars - 2
                  ElseIf Left(OutLine, 2) = "+9" Then
                     wrdApp.Selection.TypeText "("
                     Chars_Read = Chars_Read + 2
                     NumChars = NumChars - 2
                  ElseIf Left(OutLine, 2) = "+0" Then
                     wrdApp.Selection.TypeText ")"
                     Chars_Read = Chars_Read + 2
                     NumChars = NumChars - 2
                  Else
                     wrdApp.Selection.TypeText Left(OutLine, 2)
                     Chars_Read = Chars_Read + 2
                     NumChars = NumChars - 2
                  End If
                  If NumChars < 0 Then
                     wrdApp.Selection.TypeText vbCr
                     Exit Do
                  End If
                  OutLine = Mid(OutPutTab![line detail], Chars_Read + 1, NumChars)
             Loop
            End If
         Else
            If Left(OutPutTab![line detail], 8) = "Transfer" Then
               Transfers_found = "YES"
            ElseIf Left(OutPutTab![line detail], 7) = "Receive" Then
               Transfers_found = "YES"
            End If
            wrdApp.Selection.TypeText OutPutTab![line detail] & vbCrLf
         End If
         OutPutTab.MoveNext
         If OutPutTab.EOF Then
            wrdApp.Selection.TypeText vbCrLf
            Exit Do
         End If
         If Not OutPutTab!TRIBE = TRIBENUMBER Then
            wrdApp.Selection.TypeText vbCrLf
            Exit Do
         End If
     Loop
   End If
End Sub

Sub Save_Tribes_turn()
   Dim myRange As Range
   ' Check_Printing_Switchs
   
   Set Printing_Switch_TABLE = TVDB.OpenRecordset("Printing_Switchs")
   Printing_Switch_TABLE.index = "PRIMARYKEY"
   Printing_Switch_TABLE.Seek "=", CLANNUMBER
   
   Set TRADING_POST_GOODS = TVDBGM.OpenRecordset("TRADING_POST_GOODS")
   TRADING_POST_GOODS.index = "TRIBE"
   TRADING_POST_GOODS.MoveFirst
   If Not TRADING_POST_GOODS.EOF Then
      If Not Printing_Switch_TABLE.NoMatch Then
         TRADING_POST_GOODS.Seek "=", Printing_Switch_TABLE![CITY]
      End If
   End If
   If Not Printing_Switch_TABLE.NoMatch And Not TRADING_POST_GOODS.NoMatch Then
      'new page
      wrdApp.Selection.InsertBreak TYPE:=wdPageBreak
      ' clear all tabstops
      wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
      ' add in tabstops
      wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
      wrdApp.Selection.TypeText "City Trading List" & vbCr
      wrdApp.Selection.TypeText "Turn : " & Globaltable![CURRENT TURN] & vbCr & vbNewLine
      
      ' CREATE THE TABLE
      Set myRange = wrdApp.ActiveDocument.Content
      myRange.Collapse Direction:=wdCollapseEnd
 
      wrdApp.Selection.Tables.Add Range:=myRange, numrows:=40, numcolumns:=5

      ' FORMAT THE TABLE
      wrdApp.Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(4), RulerStyle:=wdAdjustNone
      wrdApp.Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
      wrdApp.Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
      wrdApp.Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
      wrdApp.Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
    
      ' GO TO THE FIRST LINE OF THE TABLE
      wrdApp.Selection.Tables(1).Cell(1, 1).Select

      ' LINE 01 OF THE TABLE
      wrdApp.Selection.TypeText "Good"
      wrdApp.Selection.Tables(1).Cell(1, 2).Select
      wrdApp.Selection.TypeText "Buy Price"
      wrdApp.Selection.Tables(1).Cell(1, 3).Select
      wrdApp.Selection.TypeText "Buy Limit"
      wrdApp.Selection.Tables(1).Cell(1, 4).Select
      wrdApp.Selection.TypeText "Sell Price"
      wrdApp.Selection.Tables(1).Cell(1, 4).Select
      wrdApp.Selection.TypeText "Sell Limit"

       ' LINE 03 OF THE TABLE
       ROW = 2
       Do While TRADING_POST_GOODS![TRIBE] = Printing_Switch_TABLE![CITY]
          ROW = ROW + 1
          wrdApp.Selection.Tables(1).Cell(ROW, 1).Select
          wrdApp.Selection.TypeText TRADING_POST_GOODS![GOOD]
          wrdApp.Selection.Tables(1).Cell(ROW, 2).Select
          wrdApp.Selection.TypeText TRADING_POST_GOODS![BUY PRICE]
          wrdApp.Selection.Tables(1).Cell(ROW, 3).Select
          wrdApp.Selection.TypeText TRADING_POST_GOODS![BUY LIMIT]
          wrdApp.Selection.Tables(1).Cell(ROW, 4).Select
          wrdApp.Selection.TypeText TRADING_POST_GOODS![SELL PRICE]
          wrdApp.Selection.Tables(1).Cell(ROW, 5).Select
          wrdApp.Selection.TypeText TRADING_POST_GOODS![SELL LIMIT]
         
          TRADING_POST_GOODS.MoveNext
          If TRADING_POST_GOODS.EOF Then
             Exit Do
          End If
       Loop
   End If
   
    '============== remove unneeded returns===========

    Dim iReplaceCount As Integer
    For i = 1 To 3
        With wrdApp.Selection.FIND
            .ClearFormatting
            .Text = "^p^p^p"
            .Replacement.ClearFormatting
            .Replacement.Text = "^p^p"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End With
    Next i
    '============== remove unneeded returns===========

 wrdApp.Documents.Save
 'wrdApp.Documents.Close
 wrdApp.ActiveDocument.Close
 Set wrdDoc = Nothing
   

End Sub

Function Open_Word(Games_Master)
    DebugOP ("Open_Word(Games_Master)")
   Set wrdApp = New Word.Application
      
   'wrdApp.Visible = True
   wrdApp.Visible = False
     
   DIRECTPATH = CurDir$ & "\documents\" & TVDirect
   DOCUMENTPATH = CurDir$ & "\documents\"
   
   CURRENT_DIRECTORY = Dir(DIRECTPATH, vbDirectory)
   If IsNull(CURRENT_DIRECTORY) Or CURRENT_DIRECTORY = "" Then
      MkDir (DIRECTPATH)
   End If

End Function

Public Function TABS_REQUIRED(SECTION_TAB)
If SECTION_TAB = "GLOBAL" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(14), Alignment:=wdAlignTabLeft
   SECTION_NAME = "GLOBAL"
ElseIf SECTION_TAB = "FARMING" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(2.3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(3.8), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(5.3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6.8), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(8.3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(9.8), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(11.3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12.8), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(14.3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(15.8), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(17.3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(18.8), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(19), Alignment:=wdAlignTabRight
   SECTION_NAME = "FARMING"
ElseIf SECTION_TAB = "ANIMALS" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(2.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(3), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(5.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(8.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(9), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(11.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(14.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(15), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(17.5), Alignment:=wdAlignTabRight
   SECTION_NAME = "ANIMALS"
ElseIf SECTION_TAB = "BOOKS" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(2), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(3), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(9), Alignment:=wdAlignTabRight
   SECTION_NAME = "BOOKS"
ElseIf SECTION_TAB = "CAPACITY" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(2.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(3), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(5.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(8.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(9), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(11.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(14.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(15), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(17.5), Alignment:=wdAlignTabRight
   SECTION_NAME = "CAPACITY"
ElseIf SECTION_TAB = "SPECIALISTS" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(4), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(13), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(14), Alignment:=wdAlignTabRight

   SECTION_NAME = "SPECIALISTS"
ElseIf SECTION_TAB = "POLITICS" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(7), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(13), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(17), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(19), Alignment:=wdAlignTabRight
   SECTION_NAME = "POLITICS"
ElseIf SECTION_TAB = "TOTAL PEOPLE" Then
   SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(2.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(3), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(5.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(8.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(9), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(11.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(14.5), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(15), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(17.5), Alignment:=wdAlignTabRight
   SECTION_NAME = "TOTAL PEOPLE"
ElseIf SECTION_TAB = "RESEARCH" Then
      SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(6), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabRight
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(18), Alignment:=wdAlignTabRight
   SECTION_NAME = "RESEARCH"
ElseIf SECTION_TAB = "COMPLETED_RESEARCH" Then
      SECTION_NAME = "TABS"
   ' clear all tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
   ' add in tabstops
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(4.5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(9), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(13.5), Alignment:=wdAlignTabLeft
   wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(16), Alignment:=wdAlignTabLeft
   SECTION_NAME = "COMPLETED_RESEARCH"
End If


End Function

Public Function Perform_Pacification_Printing()
   Program_Area = "A300 - Politics"

   Set SkillsTab = TVDBGM.OpenRecordset("Skills")
   SkillsTab.index = "PRIMARYKEY"
   SkillsTab.MoveFirst
   SkillsTab.Seek "=", TRIBENUMBER, "POLITICS"

   If Not SkillsTab.NoMatch Then
      If SkillsTab![SKILL LEVEL] >= 10 Then
         wrdApp.Selection.Font.Bold = True
         wrdApp.Selection.TypeText "State :"
         wrdApp.Selection.Font.Bold = False
         wrdApp.Selection.TypeText vbTab & "GL" & TRIBEINFO![GOVT LEVEL] & vbCr
   End If
  
   ' GET PACIFICTAION LEVELS FOR HEXES BEING CONTROLLED
   HEXMAPPOLITICS.MoveFirst
   HEXMAPPOLITICS.Seek "=", TRIBEINFO![CURRENT HEX]
   If SkillsTab![SKILL LEVEL] >= 10 Then
      If TRIBEINFO![GOVT LEVEL] = 0 Then
         If Not HEXMAPPOLITICS.NoMatch Then
            AvailableSerfs = CLng(HEXMAPPOLITICS![POPULATION] * ((HEXMAPPOLITICS![PACIFICATION_LEVEL] * 2) / 100))
            wrdApp.Selection.TypeText "Primary Hex - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbCrLf
         End If
      End If
      If TRIBEINFO![GOVT LEVEL] > 0 Then
         If HEXMAPPOLITICS.NoMatch Then
            wrdApp.Selection.TypeText "Primary Hex - PL : " & vbTab & "0" & vbCrLf
         Else
            AvailableSerfs = CLng(HEXMAPPOLITICS![POPULATION] * ((HEXMAPPOLITICS![PACIFICATION_LEVEL] * 2) / 100))
            wrdApp.Selection.TypeText "Primary Hex - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbCrLf
         End If
      End If
      If TRIBEINFO![GOVT LEVEL] >= 1 Then
         ' Need to identify all of the 'B' hexes.
         '   HEX TO N
         Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
         HEXMAPPOLITICS.MoveFirst
         HEXMAPPOLITICS.Seek "=", CURRENT_HEX
         If HEXMAPPOLITICS.NoMatch Then
            wrdApp.Selection.TypeText "Hex to N - PL : " & vbTab & "0" & vbTab
         ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
            wrdApp.Selection.TypeText "Hex to N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
         Else
            wrdApp.Selection.TypeText "Hex to N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
         End If
         '   HEX TO NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      End If
      
      If TRIBEINFO![GOVT LEVEL] >= 2 Then
      ' Need to identify all of the 'C' hexes.
      '   HEX TO N/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO N/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "NE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO NE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO NE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO SE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO S/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "SE", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO S/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO S/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "SW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO SW/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "SW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO SW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      '   HEX TO NW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW/NWE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
      '   HEX TO N/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "NW", "NONE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
      End If
      
      If TRIBEINFO![GOVT LEVEL] >= 3 Then
          '   HEX TO N/N/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "N", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO N/N/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO N/NE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "NE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/NE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/NE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/NE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO NE/NE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "NE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/NE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/NE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/NE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO NE/NE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "NE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/NE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/NE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/NE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO SE/SE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "SE", "NE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE/SE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE/SE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE/SE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO SE/SE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "SE", "SE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE/SE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE/SE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE/SE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO SE/SE/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "SE", "S", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE/SE/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE/SE/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE/SE/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO S/S/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "SE", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO S/S/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "S", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO S/S/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "SW", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO SW/SW/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "SW", "S", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO SW/SW/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "SW", "SW", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO SW/SW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "SW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO NW/NW/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NW", "SW", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW/NW/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW/NW/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW/NW/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO NW/NW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NW", "NW", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW/NW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW/NW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW/NW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO NW/NW/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NW", "N", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW/NW/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW/NW/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW/NW/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO N/N/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "NW", "NONE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
       End If
       
       If TRIBEINFO![GOVT LEVEL] >= 4 Then
          '   HEX TO N/N/N/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "N", "N", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/N/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/N/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/N/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO N/N/N/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "N", "NE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/N/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/N/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/N/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO N/NE/NE/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "NE", "NE", "N", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/NE/NE/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/NE/NE/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/NE/NE/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO N/NE/NE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "NE", "NE", "NE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/NE/NE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/NE/NE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/NE/NE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO NE/NE/NE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "NE", "NE", "NE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/NE/NE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/NE/NE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/NE/NE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO NE/NE/NE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "NE", "NE", "SE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/NE/NE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/NE/NE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/NE/NE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO NE/SE/SE/NE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "SE", "SE", "NE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/SE/SE/NE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/SE/SE/NE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/SE/SE/NE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO NE/SE/SE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NE", "SE", "SE", "SE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NE/SE/SE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NE/SE/SE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NE/SE/SE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO SE/SE/SE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "SE", "SE", "SE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE/SE/SE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE/SE/SE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE/SE/SE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO SE/SE/SE/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SE", "SE", "SE", "S", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SE/SE/SE/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SE/SE/SE/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SE/SE/SE/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO S/S/SE/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "SE", "SE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/SE/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/SE/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/SE/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO S/S/S/SE
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "S", "SE", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/S/SE - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/S/SE - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/S/SE belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO S/S/S/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "S", "S", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/S/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/S/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/S/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO S/S/S/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "S", "S", "SW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/S/S/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/S/S/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/S/S/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO S/SW/SW/S
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "SW", "SW", "S", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/SW/SW/S - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/SW/SW/S - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/SW/SW/S belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO S/SW/SW/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "S", "SW", "SW", "SW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to S/SW/SW/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to S/SW/SW/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to S/SW/SW/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO SW/SW/SW/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "SW", "SW", "SW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW/SW/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW/SW/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW/SW/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO SW/SW/SW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "SW", "SW", "NW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW/SW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW/SW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW/SW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO SW/NW/NW/SW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "NW", "NW", "SW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/NW/NW/SW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/NW/NW/SW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/NW/NW/SW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
         End If
          '   HEX TO SW/NW/NW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "SW", "NW", "NW", "NW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to SW/SW/NW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to SW/SW/NW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to SW/SW/NW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO NW/NW/NW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NW", "NW", "NW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW/NW/NW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW/NW/NW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW/NW/NW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO NW/NW/NW/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "NW", "NW", "NW", "N", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to NW/NW/NW/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to NW/NW/NW/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to NW/NW/NW/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
          '   HEX TO N/N/NW/NW
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "NW", "NW", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/NW/NW - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/NW/NW - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/NW/NW belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          '   HEX TO N/N/NW/N
          Call Get_HexMAP_and_Terrain_of_a_hex(CURRENT_HEX_MAP, "N", "N", "NW", "N", "NONE", "NONE", "NONE", "NONE")
          HEXMAPPOLITICS.MoveFirst
          HEXMAPPOLITICS.Seek "=", CURRENT_HEX
          If HEXMAPPOLITICS.NoMatch Then
             wrdApp.Selection.TypeText "Hex to N/N/NW/N - PL : " & vbTab & "0" & vbTab
          ElseIf HEXMAPPOLITICS![PL_CLAN] = CLANNUMBER Then
             wrdApp.Selection.TypeText "Hex to N/N/NW/N - PL : " & vbTab & HEXMAPPOLITICS![PACIFICATION_LEVEL] & vbTab
          Else
             wrdApp.Selection.TypeText "Hex to N/N/NW/N belongs to " & vbTab & HEXMAPPOLITICS![PL_CLAN] & vbTab
          End If
          
          wrdApp.Selection.TypeText vbCrLf
          
   End If
   End If
End If

End Function
Sub Print_Mass_Transfers()

Dim wrdTbl As Word.TABLE
Dim newRange As Word.Range
Dim TableRow As Integer

MassXfers.MoveFirst

If Not MassXfers.EOF Then

    TableRow = 1
    
    ' clear all tabstops
    wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
    ' add in tabstops
    wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
    wrdApp.Selection.TypeText "Transfers" & vbCr
    
    ' CREATE THE TABLE
    Set newRange = wrdApp.ActiveDocument.Range

    newRange.Collapse Direction:=wdCollapseEnd
    
    wrdApp.Selection.Tables.Add Range:=newRange, numrows:=1, numcolumns:=6
    wrdApp.Selection.Tables(1).Rows.WrapAroundText = True
    
    ' FORMAT THE TABLE
    wrdApp.Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(2), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(2), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(2), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(2), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(6).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(11), RulerStyle:=wdAdjustNone
    ' GO TO THE FIRST LINE OF THE TABLE
    wrdApp.Selection.Tables(1).Cell(1, 1).Select
    
    ' LINE 01 OF THE TABLE
    wrdApp.Selection.TypeText "From"
    wrdApp.Selection.Tables(1).Cell(1, 2).Select
    wrdApp.Selection.TypeText "To"
    wrdApp.Selection.Tables(1).Cell(1, 3).Select
    wrdApp.Selection.TypeText "Item"
    wrdApp.Selection.Tables(1).Cell(1, 4).Select
    wrdApp.Selection.TypeText "Requested"
    wrdApp.Selection.Tables(1).Cell(1, 5).Select
    wrdApp.Selection.TypeText "Actual"
    wrdApp.Selection.Tables(1).Cell(1, 6).Select
    wrdApp.Selection.TypeText "Message"

    Do Until MassXfers.EOF
        If MassXfers![CLAN] = CLANNUMBER Then
            TableRow = TableRow + 1
            wrdApp.Selection.Tables(1).Rows.Add
            wrdApp.Selection.Tables(1).Cell(TableRow, 1).Select
            wrdApp.Selection.TypeText MassXfers![FromUnit]
            wrdApp.Selection.Tables(1).Cell(TableRow, 2).Select
            wrdApp.Selection.TypeText MassXfers![ToUnit]
            wrdApp.Selection.Tables(1).Cell(TableRow, 3).Select
            wrdApp.Selection.TypeText MassXfers![ITEM]
            wrdApp.Selection.Tables(1).Cell(TableRow, 4).Select
            wrdApp.Selection.TypeText MassXfers![QUANTITY]
            wrdApp.Selection.Tables(1).Cell(TableRow, 5).Select
            If IsNull(MassXfers![ACTUAL_QTY]) Then
               wrdApp.Selection.TypeText " "
            Else
               wrdApp.Selection.TypeText MassXfers![ACTUAL_QTY]
            End If
            wrdApp.Selection.Tables(1).Cell(TableRow, 6).Select
            'wrdApp.Selection.TypeText MassXfers![PROCESS_MSG] & " (" & Trim(MassXfers![RPT_CODE]) & ")"
            wrdApp.Selection.TypeText MassXfers![PROCESS_MSG]
        End If
        MassXfers.MoveNext
    Loop
    wrdApp.Selection.Font.Reset
End If



End Sub
Sub Print_Settlements()

Dim wrdTbl As Word.TABLE
Dim newRange As Word.Range
Dim TableRow As Integer

HEXMAPCITY.MoveFirst

If Not HEXMAPCITY.EOF Then

    TableRow = 1
    
    wrdApp.Selection.EndOf Unit:=wdStory, Extend:=wdMove
    ' clear all tabstops
    wrdApp.Selection.Paragraphs(1).TabStops.ClearAll
    ' add in tabstops
    wrdApp.Selection.Paragraphs(1).TabStops.Add POSITION:=wrdApp.CentimetersToPoints(12), Alignment:=wdAlignTabLeft
    wrdApp.Selection.Font.Bold = True
    wrdApp.Selection.TypeText vbCr & vbCr & "Settlements" & vbCr
    wrdApp.Selection.Font.Bold = False
    
    ' CREATE THE TABLE
    Set newRange = wrdApp.ActiveDocument.Range

    newRange.Collapse Direction:=wdCollapseEnd
    
    wrdApp.Selection.Tables.Add Range:=newRange, numrows:=1, numcolumns:=5
    wrdApp.Selection.Tables(1).Rows.WrapAroundText = True
    
    ' FORMAT THE TABLE
    wrdApp.Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(2), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(4), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(3).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(4).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
    wrdApp.Selection.Tables(1).Columns(5).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
    
    ' GO TO THE FIRST LINE OF THE TABLE
    wrdApp.Selection.Tables(1).Cell(1, 1).Select
    
    ' LINE 01 OF THE TABLE
    wrdApp.Selection.TypeText "Hex Code"
    wrdApp.Selection.Tables(1).Cell(1, 2).Select
    wrdApp.Selection.TypeText "Name"
    wrdApp.Selection.Tables(1).Cell(1, 3).Select
    wrdApp.Selection.TypeText "Note"
    wrdApp.Selection.Tables(1).Cell(1, 4).Select
    wrdApp.Selection.TypeText "Type"
    wrdApp.Selection.Tables(1).Cell(1, 5).Select
    wrdApp.Selection.TypeText "Subtype"

    Do Until HEXMAPCITY.EOF
        If HEXMAPCITY![OWNER] = CLANNUMBER Then
            TableRow = TableRow + 1
            wrdApp.Selection.Tables(1).Rows.Add
            wrdApp.Selection.Tables(1).Cell(TableRow, 1).Select
            wrdApp.Selection.TypeText HEXMAPCITY![MAP] & " "
            wrdApp.Selection.Tables(1).Cell(TableRow, 2).Select
            wrdApp.Selection.TypeText HEXMAPCITY![CITY] & " "
            wrdApp.Selection.Tables(1).Cell(TableRow, 3).Select
            wrdApp.Selection.TypeText HEXMAPCITY![CITY_2] & " "
            wrdApp.Selection.Tables(1).Cell(TableRow, 4).Select
            wrdApp.Selection.TypeText HEXMAPCITY![TYPE] & " "
            wrdApp.Selection.Tables(1).Cell(TableRow, 5).Select
            wrdApp.Selection.TypeText HEXMAPCITY![SUBTYPE] & " "
        End If
        HEXMAPCITY.MoveNext
    Loop
    wrdApp.Selection.Font.Reset
End If

End Sub
Sub Print_Special_Routes()

Dim wrdTbl As Word.TABLE
Dim newRange As Word.Range
Dim TableRow As Integer

Special_Routes.MoveFirst

If Not Special_Routes.EOF Then

    TableRow = 1
    
    wrdApp.Selection.EndOf Unit:=wdStory, Extend:=wdMove
    Set newRange = wrdApp.ActiveDocument.Content
    newRange.Collapse Direction:=wdCollapseEnd
    newRange.InsertParagraph
    newRange.InsertParagraph
    newRange.InsertAfter "Special Routes" & vbCr

    Set newRange = wrdApp.ActiveDocument.Range
    newRange.Collapse Direction:=wdCollapseEnd
        
    Set wrdTbl = wrdApp.Selection.Tables.Add(Range:=newRange, numrows:=1, numcolumns:=5)
    With wrdTbl
        .Rows.WrapAroundText = False
    
        ' FORMAT THE TABLE
        .Columns(1).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(5), RulerStyle:=wdAdjustNone
        .Columns(2).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(4), RulerStyle:=wdAdjustNone
        .Columns(3).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
        .Columns(4).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
        .Columns(5).SetWidth ColumnWidth:=wrdApp.CentimetersToPoints(3), RulerStyle:=wdAdjustNone
        
        
        ' LINE 01 OF THE TABLE
        .Cell(TableRow, 1).Range.Text = "Route Name"
        .Cell(TableRow, 2).Range.Text = "Type"
        .Cell(TableRow, 3).Range.Text = "Subtype"
        .Cell(TableRow, 4).Range.Text = "From Hex"
        .Cell(TableRow, 5).Range.Text = "To Hex"
    
        Do Until Special_Routes.EOF
            If Special_Routes![OWNER] = CLANNUMBER Then
                TableRow = TableRow + 1
                .Rows.Add
                .Cell(TableRow, 1).Range.Text = Special_Routes![Route_Name] & " "
                .Cell(TableRow, 2).Range.Text = Special_Routes![Route_Type] & " "
                .Cell(TableRow, 3).Range.Text = Special_Routes![SUBTYPE] & " "
                .Cell(TableRow, 4).Range.Text = Special_Routes![From_Hex] & " "
                .Cell(TableRow, 5).Range.Text = Special_Routes![To_Hex] & " "
            End If
            Special_Routes.MoveNext
        Loop
    End With
    wrdApp.Selection.Font.Reset
End If

End Sub


