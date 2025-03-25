Attribute VB_Name = "ImportProcesses"
Option Compare Database   'Use database order for string comparisons
Option Explicit

'Function MASS_TRANSFERS()
Function MASS_TRANSFERS(XferTiming As String, CALLER As String)
On Error GoTo ERR_MASS_TRANSFERS
TRIBE_STATUS = "Mass Transfers"

' Set up variables to be used
Dim TVWKSPACE As Workspace
Dim TVDB, TVDBGM As DAO.Database
Dim GMTABLE, CurrentTurn As Recordset
Dim TRIBESGOOD As Recordset, OUTTAB As Recordset
Dim MassXfers, MEETINGHOUSE, hexmaptable, SpecialXfer, citytable As Recordset
Dim INFILE As String
Dim FROMCLAN As String
Dim FROMTRIBE As String
Dim FROMHEX As String
Dim TOCLAN As String
Dim TOTRIBE As String
Dim TOHEX As String
Dim ITEM As String
Dim OUTPUTLINE As String
Dim QUANTITY As Long
Dim count As Long
Dim INGOODSTRIBE As String
Dim TOGOODSTRIBE As String
Dim LINENUMBER As Long
Dim POSITION As Long
Dim ADJ_HEXES(6) As String
Dim RIVERS(6) As Boolean
Dim BOGUS As Variant
Dim VALIDXFER As Boolean
Dim ISSPECIALXFER As Boolean
Dim i As Integer
Dim j As Integer
Dim ISRIVER As String
Dim FROMCLANMH As Boolean
Dim TOCLANMH As Boolean
Dim ErrorMsg As String
Dim InfoMsg As String
Dim ReportCode As Long
Dim ReportClan As String
Dim CurrentMonth As String
Dim intAnswer As String
Dim qryName As String
'Dim XferTiming As String
'Dim CALLER As String


If XferTiming = "AM" Then
    intAnswer = MsgBox("You are about to process AFTER movement transfers. Are you sure that you want to do this?", vbOKCancel)
    If intAnswer = vbCancel Then
        Exit Function
    End If
End If

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

' FILEGM = CurDir$ & "\" & GMTABLE![FILE]
FILEGM = GMTABLE![DIRECTORY] & "\" & GMTABLE![FILE]
TVDB.Close

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

'The following two lines can be uncommented for debugging
'XferTiming = "BM"
'CALLER = "IMPORT_TRANSFERS"

' Open transfer table with query returning only unprocessed records
If XferTiming = "AM" Then
    Set MassXfers = TVDBGM.OpenRecordset("SELECT * FROM MassTransfers WHERE (PROCESSED = 'N' OR PROCESSED IS NULL) and TRANSFER_TIMING = 'AM'")
Else
    Set MassXfers = TVDBGM.OpenRecordset("SELECT * FROM MassTransfers WHERE (PROCESSED = 'N' OR PROCESSED IS NULL) and (not TRANSFER_TIMING = 'AM' or TRANSFER_TIMING is null)")
End If

' Open Tribe Checking table
Set TRIBES_CHECKING = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TRIBES_CHECKING.index = "PRIMARYKEY"

' Open Hex Map Contruction table
Set MEETINGHOUSE = TVDBGM.OpenRecordset("HEX_MAP_CONST")
MEETINGHOUSE.index = "PRIMARYKEY"

' Open Hex Map table
Set hexmaptable = TVDBGM.OpenRecordset("HEX_MAP")
hexmaptable.index = "PRIMARYKEY"

' Open Hex Map City table
Set citytable = TVDBGM.OpenRecordset("HEX_MAP_CITY")
citytable.index = "PRIMARYKEY"

' Open Special Transfer Routes table
Set SpecialXfer = TVDBGM.OpenRecordset("Special_Transfer_Routes")
SpecialXfer.index = "PRIMARYKEY"

' Open Turn Activities table
Set OUTTAB = TVDBGM.OpenRecordset("TURNS_ACTIVITIES")
OUTTAB.index = "primarykey"

' Open Valid Goods table
Set VALIDGOODS = TVDBGM.OpenRecordset("VALID_GOODS")
VALIDGOODS.index = "primarykey"

' Open Tribes Goods table
Set TRIBESGOOD = TVDBGM.OpenRecordset("TRIBES_GOODS")
TRIBESGOOD.index = "primarykey"

' Open Tribes General Information table
Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
TRIBESINFO.index = "PRIMARYKEY"

' Open GLOBAL table
Set CurrentTurn = TVDBGM.OpenRecordset("GLOBAL")
CurrentTurn.index = "PRIMARYKEY"

' Get Current Month
CurrentTurn.MoveFirst
CurrentMonth = Mid(CurrentTurn![CURRENT TURN], 1, 2)
CurrentTurn.Close

' Loop thru once deleting null records, setting null PROCESSED fields, and marking invalid quantities
MassXfers.MoveFirst
Do Until MassXfers.EOF
    If IsNull(Trim(MassXfers![From])) Or IsNull(Trim(MassXfers![To])) Or IsNull(Trim(MassXfers![ITEM])) Or IsNull(MassXfers![QUANTITY]) Then
       'MassXfers.Delete
       MassXfers.Edit
       MassXfers![PROCESSED] = "X"
       MassXfers![PROCESS_MSG] = "A required field is null."
       MassXfers![REPORT_CODE] = 1
       MassXfers.UPDATE
    ElseIf MassXfers![QUANTITY] <= 0 Then
       MassXfers.Edit
       MassXfers![PROCESSED] = "X"
       MassXfers![PROCESS_MSG] = "Invalid quantity, probably 0 or negative."
       MassXfers![REPORT_CODE] = 1
       MassXfers.UPDATE
    Else
        If IsNull(MassXfers![PROCESSED]) Then
            MassXfers.Edit
            MassXfers![PROCESSED] = "N"
            MassXfers.UPDATE
        End If
        If IsNull(MassXfers![TRANSFER_TIMING]) Then
            MassXfers.Edit
            MassXfers![TRANSFER_TIMING] = "BM"
            MassXfers.UPDATE
        End If
    End If
    MassXfers.MoveNext
Loop
MassXfers.MoveFirst

' LOOP THROUGH MassTransfers DOING EACH TRANSFER
Do Until MassXfers.EOF

    If MassXfers![PROCESSED] = "N" Then
        ' Set or reset messages and flags
        ErrorMsg = ""
        InfoMsg = ""
        ReportCode = 1
        ReportClan = ""
        VALIDXFER = False
        ' Set or reset data variables
        INGOODSTRIBE = ""
        TOGOODSTRIBE = ""
        TOHEX = ""
        FROMHEX = ""
        
        ' Instantiate variables from current record
        FROMCLAN = "0" & Mid(MassXfers![From], 2, 3)
        FROMTRIBE = Trim(MassXfers![From])
        TOCLAN = "0" & Mid(MassXfers![To], 2, 3)
        TOTRIBE = Trim(MassXfers![To])
        ITEM = Trim(MassXfers![ITEM])
        QUANTITY = MassXfers![QUANTITY]
        
        ' Check for goods tribe relationship
        TRIBESINFO.MoveFirst
        If FROMCLAN = "0263" Then
            INGOODSTRIBE = "0263"
            ReportClan = TOCLAN
        Else
            ReportClan = FROMCLAN
            TRIBESINFO.Seek "=", FROMCLAN, FROMTRIBE
            If TRIBESINFO.NoMatch Then
                ErrorMsg = "From tribe general info not found."
                GoTo Err_Invalid_Transfer
            End If
            If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
                INGOODSTRIBE = TRIBESINFO![GOODS TRIBE]
            Else
                INGOODSTRIBE = FROMTRIBE
            End If
        End If
    
        If TOCLAN = "0263" Then
            TOGOODSTRIBE = "0263"
        Else
            TRIBESINFO.MoveFirst
            TRIBESINFO.Seek "=", TOCLAN, TOTRIBE
            If TRIBESINFO.NoMatch Then
                ErrorMsg = "To tribe general info not found."
                GoTo Err_Invalid_Transfer
            End If
            If Not IsNull(TRIBESINFO![GOODS TRIBE]) Then
                TOGOODSTRIBE = TRIBESINFO![GOODS TRIBE]
            Else
                TOGOODSTRIBE = TOTRIBE
            End If
        End If
        
        ' Get locations
        If FROMCLAN <> "0263" Then
            TRIBES_CHECKING.MoveFirst
            TRIBES_CHECKING.Seek "=", FROMCLAN, FROMTRIBE
            If TRIBES_CHECKING.NoMatch Then
               ErrorMsg = "From unit location not found"
               GoTo Err_Invalid_Transfer
            End If
            FROMHEX = TRIBES_CHECKING![Current Hex]
        End If
        If TOCLAN <> "0263" Then
            TRIBES_CHECKING.MoveFirst
            TRIBES_CHECKING.Seek "=", TOCLAN, TOTRIBE
            If TRIBES_CHECKING.NoMatch Then
               ErrorMsg = "To unit location not found"
               GoTo Err_Invalid_Transfer
            End If
            TOHEX = TRIBES_CHECKING![Current Hex]
        End If
         
        ' Check for valid good
        VALIDGOODS.MoveFirst
        VALIDGOODS.Seek "=", ITEM
        If VALIDGOODS.NoMatch Then
            ErrorMsg = "Not a valid good."
            GoTo Err_Invalid_Transfer
        End If
          
        ' Determine whether meeting house is in one of the hexes
        FROMCLANMH = False
        TOCLANMH = False
        If FROMCLAN <> "0263" Then
            MEETINGHOUSE.MoveFirst
            MEETINGHOUSE.Seek "=", FROMHEX, FROMCLAN
            If Not MEETINGHOUSE.NoMatch Then
               Do While True
                  If MEETINGHOUSE![CONSTRUCTION] = "MEETING HOUSE" Then
                     FROMCLANMH = True
                     Exit Do
                  End If
                  MEETINGHOUSE.MoveNext
                  If MEETINGHOUSE![MAP] <> FROMHEX Or MEETINGHOUSE![CLAN] <> FROMCLAN Then
                     Exit Do
                  End If
               Loop
            End If
        End If
        If Not FROMCLANMH And TOCLAN <> "0263" Then
            MEETINGHOUSE.MoveFirst
            MEETINGHOUSE.Seek "=", TOHEX, TOCLAN
            If Not MEETINGHOUSE.NoMatch Then
                Do While True
                    If MEETINGHOUSE![CONSTRUCTION] = "MEETING HOUSE" Then
                        TOCLANMH = True
                        Exit Do
                    End If
                    MEETINGHOUSE.MoveNext
                    If MEETINGHOUSE![MAP] <> TOHEX Or MEETINGHOUSE![CLAN] <> TOCLAN Then
                        Exit Do
                    End If
                Loop
            End If
        End If

        If FROMCLAN <> "0263" Then
            
            ' Get river sides of from hex
            hexmaptable.MoveFirst
            hexmaptable.Seek "=", FROMHEX
            i = 1
            Do While i <= 6
                j = 1 + (i - 1) * 2
                ISRIVER = Mid(hexmaptable![Borders], j, 2)
                If ISRIVER = "RI" Then
                    RIVERS(i) = True
                Else
                    RIVERS(i) = False
                End If
                i = i + 1
            Loop
        End If
        
        ' Test whether the two hexes are part of a special transfer route
        ISSPECIALXFER = False
        
        If FROMCLAN <> "0263" And TOCLAN <> "0263" Then

            If Not (SpecialXfer.EOF And SpecialXfer.BOF) Then
                SpecialXfer.MoveFirst
                Do Until SpecialXfer.EOF = True
                    If FROMHEX = SpecialXfer![From_Hex] Then
                        If TOHEX = SpecialXfer![To_Hex] Then
                            ISSPECIALXFER = True
                            Exit Do
                        End If
                    ElseIf FROMHEX = SpecialXfer![To_Hex] Then
                        If TOHEX = SpecialXfer![From_Hex] Then
                            ISSPECIALXFER = True
                            Exit Do
                        End If
                    ElseIf SpecialXfer![TYPE] = "Trade Envoy" Then
                        If (FROMTRIBE = SpecialXfer![From_Hex] And TOTRIBE = SpecialXfer![To_Hex]) Or _
                                (TOTRIBE = SpecialXfer![From_Hex] And FROMTRIBE = SpecialXfer![To_Hex]) Then
                            ISSPECIALXFER = True
                            Exit Do
                        End If
                    End If
                    SpecialXfer.MoveNext

                Loop
                If ISSPECIALXFER Then
                    If Not Mid(SpecialXfer![Valid_Months], CInt(CurrentMonth), 1) = "Y" Then
                        ISSPECIALXFER = False
                        QUANTITY = 0
                        VALIDXFER = False
                        InfoMsg = "Not a valid month to use " & SpecialXfer![TYPE] & "."
                        GoTo LINE_DETAIL_SECTION
                    End If
                End If
            End If
        End If
              
        ' Check for type of transfer
        If ISSPECIALXFER Then
            VALIDXFER = True
            ReportCode = 3
            InfoMsg = "Special transfer - " & SpecialXfer![TYPE] & ". " & XferTiming
        ElseIf FROMCLAN = "0263" Then
            If FROMTRIBE = "2263" Then
                VALIDXFER = True
                ReportCode = 7
                InfoMsg = "From research bonuses."
            ElseIf FROMTRIBE = "3263" Then
                citytable.MoveFirst
                citytable.Seek "=", TOHEX
                If citytable.NoMatch Then
                    VALIDXFER = False
                    InfoMsg = "Unit is not in a trade city hex. " & TOHEX
                Else
                    VALIDXFER = True
                    ReportCode = 5
                    InfoMsg = "Trade city transfer from " & citytable![CITY] & "."
                End If
            ElseIf FROMTRIBE = "7263" Then
                If CurrentMonth = "04" Or CurrentMonth = "10" Then
                    VALIDXFER = True
                    ReportCode = 6
                    InfoMsg = "Goods transfer from fair"
                Else
                    VALIDXFER = False
                    InfoMsg = "Transfer from fair transfer in a non-fair month."
                End If
            Else
                VALIDXFER = True
                ReportCode = 7
                InfoMsg = "Transfer from clan 0263"
            End If
        ElseIf TOCLAN = "0263" Then
            If TOTRIBE = "3263" Then
                citytable.MoveFirst
                citytable.Seek "=", FROMHEX
                If citytable.NoMatch Then
                    VALIDXFER = False
                    InfoMsg = "Unit is not in a trade city hex. " & FROMHEX
                Else
                    VALIDXFER = True
                    ReportCode = 5
                    InfoMsg = "Trade city transfer to " & citytable![CITY] & "."
                End If
            ElseIf TOTRIBE = "7263" Then
                If CurrentMonth = "04" Or CurrentMonth = "10" Then
                    VALIDXFER = True
                    ReportCode = 6
                    InfoMsg = "Goods transfer to fair"
                Else
                    VALIDXFER = False
                    InfoMsg = "Transfer to fair in a non-fair month."
                End If
            ElseIf TOTRIBE = "1263" Then
                VALIDXFER = True
                ReportCode = 4
                InfoMsg = "Goods transfer to usage"
            Else
                VALIDXFER = True
                ReportCode = 7
                InfoMsg = "Transfer to clan 0263"
            End If
        ElseIf TOHEX = FROMHEX Then
           VALIDXFER = True
           ReportCode = 7
           InfoMsg = "Transfer in same hex"
        ElseIf FROMCLAN = TOCLAN And (FROMCLANMH Or TOCLANMH) Then
           ' Get surrounding hexes
           VALIDXFER = False
           ADJ_HEXES(1) = GET_MAP_NORTH(FROMHEX)
           ADJ_HEXES(2) = GET_MAP_NORTH_EAST(FROMHEX)
           ADJ_HEXES(3) = GET_MAP_SOUTH_EAST(FROMHEX)
           ADJ_HEXES(4) = GET_MAP_SOUTH(FROMHEX)
           ADJ_HEXES(5) = GET_MAP_SOUTH_WEST(FROMHEX)
           ADJ_HEXES(6) = GET_MAP_NORTH_WEST(FROMHEX)
           i = 1
           Do Until i > 6
              If TOHEX = ADJ_HEXES(i) Then
                 If Not RIVERS(i) Then
                    VALIDXFER = True
                    InfoMsg = "Adjacent hexes with MH. (" & i & ")"
                 Else
                    InfoMsg = "There is a river between these units. (" & i & ") From:" & FROMHEX & " To:" & TOHEX
                 End If
              End If
              i = i + 1
           Loop
        Else
            If FROMCLAN = TOCLAN Then
                InfoMsg = "Could not transfer " & XferTiming & ". From:" & FROMHEX & " To:" & TOHEX
            Else
                InfoMsg = "Could not transfer " & XferTiming & " between clans"
            End If
        End If
        
        If Not VALIDXFER Then
            QUANTITY = 0
            If Not Trim(InfoMsg) = Null Then
                ErrorMsg = "Transfer not categorized."
            End If
        End If
        
LINE_DETAIL_SECTION:
          
        LINENUMBER = 1
            
        'SETUP LINE DETAIL
        OUTTAB.MoveFirst
        OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
            
        If OUTTAB.NoMatch Then
            OUTTAB.AddNew
            OUTTAB![CLAN] = FROMCLAN
            OUTTAB![TRIBE] = FROMTRIBE
            OUTTAB![Section] = "TRANSFERS OUT"
            OUTTAB![LINE NUMBER] = LINENUMBER
            OUTTAB![line detail] = "Transfer goods to " & TOTRIBE & ": "
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
            ' IF LINE CONTAINS TOTRIBE THEN JUST COMMA ELSE
            POSITION = InStr(OUTTAB![line detail], TOTRIBE)
            If POSITION > 0 Then
               OUTTAB![line detail] = OUTTAB![line detail] & " "
            Else
               OUTTAB![line detail] = OUTTAB![line detail] & ",To " & TOTRIBE & ": "
            End If
            OUTTAB.UPDATE
            OUTTAB.MoveFirst
            OUTTAB.Seek "=", FROMCLAN, FROMTRIBE, "TRANSFERS OUT", LINENUMBER
            OUTTAB.Edit
        End If
            
        OUTPUTLINE = OUTTAB![line detail]
        
        If QUANTITY > 0 Then
            VALIDGOODS.MoveFirst
            VALIDGOODS.Seek "=", ITEM
            If VALIDGOODS.NoMatch Then
               QUANTITY = 0
               ReportCode = 2
               InfoMsg = "Attempt to transfer " & ITEM & ", but it is not a valid good."
            ElseIf Not VALIDGOODS![TABLE] = "GENERAL" And Not VALIDGOODS![TABLE] = "HUMANS" Then
                TRIBESGOOD.MoveFirst
                TRIBESGOOD.Seek "=", FROMCLAN, INGOODSTRIBE, VALIDGOODS![TABLE], ITEM
                If TRIBESGOOD.NoMatch Then
                    QUANTITY = 0
                    ReportCode = 2
                    InfoMsg = "Attempt to transfer " & ITEM & ", but none exists in inventory."
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
            ElseIf VALIDGOODS![TABLE] = "HUMANS" Then
                TRIBESINFO.MoveFirst
                TRIBESINFO.Seek "=", FROMCLAN, FROMTRIBE
                If TRIBESINFO.NoMatch Then
                    QUANTITY = 0
                    ReportCode = 2
                    InfoMsg = FROMTRIBE & " has no record in the tribe info table."
                Else
                    TRIBESINFO.Edit
                    If ITEM = "SLAVE" Then
                        If QUANTITY >= TRIBESINFO![SLAVE] Then
                            QUANTITY = TRIBESINFO![SLAVE]
                        End If
                        TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] - QUANTITY
                    ElseIf ITEM = "HIRELINGS" Then
                        If QUANTITY >= TRIBESINFO![HIRELINGS] Then
                            QUANTITY = TRIBESINFO![HIRELINGS]
                        End If
                        TRIBESINFO![HIRELINGS] = TRIBESINFO![HIRELINGS] - QUANTITY
                    ElseIf ITEM = "MERCENARIES" Then
                        If QUANTITY >= TRIBESINFO![MERCENARIES] Then
                            QUANTITY = TRIBESINFO![MERCENARIES]
                        End If
                        TRIBESINFO![MERCENARIES] = TRIBESINFO![MERCENARIES] - QUANTITY
                    Else
                        If (FROMCLAN = TOCLAN) Or FROMCLAN = "0263" Then
                            If ITEM = "WARRIORS" Then
                                If QUANTITY >= TRIBESINFO![WARRIORS] Then
                                    QUANTITY = TRIBESINFO![WARRIORS]
                                End If
                                TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] - QUANTITY
                            ElseIf ITEM = "ACTIVES" Then
                                If QUANTITY >= TRIBESINFO![ACTIVES] Then
                                    QUANTITY = TRIBESINFO![ACTIVES]
                                End If
                                TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] - QUANTITY
                            ElseIf ITEM = "INACTIVES" Then
                                If QUANTITY >= TRIBESINFO![INACTIVES] Then
                                    QUANTITY = TRIBESINFO![INACTIVES]
                                End If
                                TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] - QUANTITY
                            End If
                        Else
                            QUANTITY = 0
                            ReportCode = 2
                            InfoMsg = "W/A/I may not be transferred between clans."
                        End If
                    End If
                    TRIBESINFO.UPDATE
                End If
            End If
        End If
            
        OUTPUTLINE = OUTPUTLINE & QUANTITY & " " & ITEM & ", "
        
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
        Else
           OUTTAB.Edit
           OUTTAB![line detail] = OUTPUTLINE
           OUTTAB.UPDATE
        End If
    
        LINENUMBER = 1
    
        OUTTAB.MoveFirst
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
           ' IF LINE CONTAINS TOTRIBE THEN JUST COMMA ELSE
           POSITION = InStr(OUTTAB![line detail], FROMTRIBE)
           If POSITION > 0 Then
              OUTTAB![line detail] = OUTTAB![line detail] & " "
           Else
              OUTTAB![line detail] = OUTTAB![line detail] & ",from " & FROMTRIBE & ": "
           End If
           OUTTAB.UPDATE
           OUTTAB.MoveFirst
           OUTTAB.Seek "=", TOCLAN, TOTRIBE, "TRANSFERS IN", LINENUMBER
           OUTTAB.Edit
        End If
    
        OUTPUTLINE = OUTTAB![line detail]
    
        If QUANTITY > 0 Then
            VALIDGOODS.MoveFirst
            VALIDGOODS.Seek "=", ITEM
            If Not VALIDGOODS![TABLE] = "GENERAL" And Not VALIDGOODS![TABLE] = "HUMANS" Then
                TRIBESGOOD.MoveFirst
                TRIBESGOOD.Seek "=", TOCLAN, TOGOODSTRIBE, VALIDGOODS![TABLE], ITEM
                
                If TRIBESGOOD.NoMatch Then
                   TRIBESGOOD.AddNew
                   TRIBESGOOD![CLAN] = TOCLAN
                   TRIBESGOOD![TRIBE] = TOGOODSTRIBE
                   TRIBESGOOD![ITEM_TYPE] = VALIDGOODS![TABLE]
                   TRIBESGOOD![ITEM] = ITEM
                   TRIBESGOOD![ITEM_NUMBER] = QUANTITY
                   TRIBESGOOD.UPDATE
                Else
                   TRIBESGOOD.Edit
                   TRIBESGOOD![ITEM_NUMBER] = TRIBESGOOD![ITEM_NUMBER] + QUANTITY
                   TRIBESGOOD.UPDATE
                End If
            
            ElseIf VALIDGOODS![TABLE] = "HUMANS" Then
               TRIBESINFO.MoveFirst
               TRIBESINFO.Seek "=", TOCLAN, TOTRIBE
               
               TRIBESINFO.Edit
            
               If ITEM = "SLAVE" Then
                  TRIBESINFO![SLAVE] = TRIBESINFO![SLAVE] + QUANTITY
               ElseIf ITEM = "HIRELINGS" Then
                  TRIBESINFO![HIRELINGS] = TRIBESINFO![HIRELINGS] + QUANTITY
               ElseIf ITEM = "MERCENARIES" Then
                  TRIBESINFO![MERCENARIES] = TRIBESINFO![MERCENARIES] + QUANTITY
               ElseIf ITEM = "WARRIORS" Then
                  TRIBESINFO![WARRIORS] = TRIBESINFO![WARRIORS] + QUANTITY
               ElseIf ITEM = "ACTIVES" Then
                  TRIBESINFO![ACTIVES] = TRIBESINFO![ACTIVES] + QUANTITY
               ElseIf ITEM = "INACTIVES" Then
                  TRIBESINFO![INACTIVES] = TRIBESINFO![INACTIVES] + QUANTITY
               End If
               TRIBESINFO.UPDATE
            End If
        End If
        
           
        OUTPUTLINE = OUTPUTLINE & QUANTITY & " " & ITEM & ", "
                           
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
        Else
           OUTTAB.Edit
           OUTTAB![line detail] = OUTPUTLINE
           OUTTAB.UPDATE
        End If
      
Err_Invalid_Transfer:
        MassXfers.Edit
        If Not ErrorMsg = "" Then
            MassXfers![ACTUAL_QTY] = 0
            MassXfers![PROCESSED] = "X"
            MassXfers![REPORT_CODE] = 1
            MassXfers![PROCESS_MSG] = ErrorMsg
        Else
            If QUANTITY <> MassXfers![QUANTITY] Then
                MassXfers![ACTUAL_QTY] = QUANTITY
            End If
            MassXfers![PROCESSED] = "Y"
            MassXfers![REPORT_CODE] = ReportCode
            MassXfers![PROCESS_MSG] = InfoMsg
        End If
        MassXfers![REPORT_CLAN] = ReportClan
        MassXfers.UPDATE
    End If
    
    MassXfers.MoveNext
    
Loop

ERR_MASS_TRANSFERS_CLOSE:
    MassXfers.Close
    TRIBESINFO.Close
    TRIBES_CHECKING.Close
    MEETINGHOUSE.Close
    VALIDGOODS.Close
    hexmaptable.Close
    SpecialXfer.Close
    OUTTAB.Close
    TRIBESGOOD.Close
    TVDBGM.Close
    
    ' reset implements
    Call Reset_Implements_and_Goods_Usage_Tables
        
    Call POPULATE_CAPACITIES

    Call POPULATE_WEIGHTS
    
    DoCmd.Hourglass False
    
    MsgBox "Transfer processing complete.", vbInformation, "Done"

    DoCmd.Close acForm, CALLER
    DoCmd.OpenForm CALLER
    
    Exit Function

ERR_MASS_TRANSFERS:
    If (Err = 3021) Then
       Resume Next
    
    Else
      Msg = "Error # " & Err & " " & Error$
      MsgBox (Msg)
      Resume ERR_MASS_TRANSFERS_CLOSE
    End If

End Function
