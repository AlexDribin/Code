Attribute VB_Name = "Importing"
Option Compare Database
Option Explicit
Public Function Export_Hexmap_Data()
Dim Export_File As String
Dim Export_Table As String
Dim qdfCurrent As QueryDef
Dim Import_Trades As Recordset
Dim Trading_Post As Recordset
Dim QUERY_STRING As String

DoCmd.Hourglass True

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Export_File = CurDir$ & "\documents\Hex_Map_Data.csv"
Export_Table = "HEX_MAP"

' this will export hex_map into a 2010 spreadsheet specified above
' will include field names

'DoCmd.TransferSpreadsheet acExport, 9, Export_Table, Export_File, True
DoCmd.TransferText acExportDelim, , Export_Table, Export_File, True

Export_File = CurDir$ & "\documents\Hex_Map_City_Data.csv"
Export_Table = "HEX_MAP_CITY"

DoCmd.TransferText acExportDelim, , Export_Table, Export_File, True

DoCmd.Hourglass False

End Function


Public Function Import_Trading_Post_Spreadsheet(TPS_CLANNUMBER)
Dim Import_File As String
Dim Import_Table As String
Dim qdfCurrent As QueryDef
Dim Import_Trades As Recordset
Dim Trading_Post As Recordset
Dim QUERY_STRING As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

' delete existing clan trading post data
QUERY_STRING = "DELETE * FROM TRADING_POST_GOODS"
QUERY_STRING = QUERY_STRING & " WHERE (((TRADING_POST_GOODS.TRIBE)='"
QUERY_STRING = QUERY_STRING & TPS_CLANNUMBER & "'));"
Set qdfCurrent = TVDBGM.CreateQueryDef("", QUERY_STRING)
qdfCurrent.Execute

Import_File = CurDir$ & "\trading_posts\" & TPS_CLANNUMBER & "_Trading_Post.xls"
Import_Table = "Import_Trading_Post_Goods"

' this will import the spreadsheet specified above into the table specified.
' the spreadsheet must have headings included.

DoCmd.TransferSpreadsheet acImport, 8, Import_Table, Import_File, True

' Transfer the info into the Trading_Post_Goods table
Set Trading_Post = TVDBGM.OpenRecordset("Trading_Post_Goods")
Trading_Post.MoveFirst

Set Import_Trades = TVDBGM.OpenRecordset("Import_Trading_Post_Goods")
Import_Trades.MoveFirst

Do While Not Import_Trades.EOF
   Trading_Post.AddNew
   Trading_Post![TYPE_OF_TRADING_POST] = "TRIBE"
   Trading_Post![TRIBE] = Import_Trades![TRIBE]
   Trading_Post![GOOD] = Import_Trades![GOOD]
   Trading_Post![HEX_MAP_ID] = "BA"
   Trading_Post![BUY PRICE] = Import_Trades![BUY PRICE]
   Trading_Post![BUY LIMIT] = Import_Trades![BUY LIMIT]
   Trading_Post![BUY_RESET_WAIT] = 0
   Trading_Post![NORMAL_BUY_LIMIT] = Import_Trades![BUY LIMIT]
   Trading_Post![TURNS_SINCE_LAST_BUY] = 0
   Trading_Post![BUY_THIS_TURN] = "N"
   Trading_Post![BUY_TOTAL] = 0
   Trading_Post![SELL PRICE] = Import_Trades![SELL PRICE]
   Trading_Post![SELL LIMIT] = Import_Trades![SELL LIMIT]
   Trading_Post![SELL_RESET_WAIT] = 0
   Trading_Post![NORMAL_SELL_LIMIT] = Import_Trades![SELL LIMIT]
   Trading_Post![TURNS_SINCE_LAST_SELL] = 0
   Trading_Post![SELL_THIS_TURN] = "N"
   Trading_Post![SELL_TOTAL] = 0
   Trading_Post.UPDATE

   Import_Trades.Delete
   Import_Trades.MoveFirst
   If Import_Trades.EOF Then
      Exit Do
   End If
Loop

Trading_Post.Close
Import_Trades.Close



End Function


Public Function Importing_Clan_Spreadsheets()
On Error GoTo ERR_TABLES
TRIBE_STATUS = "IMporting Clan Spreadsheets"

Dim Import_File As String
Dim Import_Table As String
Dim qdfCurrent As QueryDef
Dim QUERY_STRING As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

' loop through the clans with tp's
' look at the hex_map_const table

Set TRIBECHECK = TVDBGM.OpenRecordset("HEX_MAP_CONST")
TRIBECHECK.index = "PRIMARYKEY"
TRIBECHECK.MoveFirst

Do
   If TRIBECHECK![CONSTRUCTION] = "TRADING POST" Then
      CLANNUMBER = TRIBECHECK![CLAN]
      Import_File = CurDir$ & "\trading_posts\" & CLANNUMBER & "_Trading_Post.xls"
      
      'IS FILE THERE
      Open Import_File For Input As #1
      Close #1
    
     ' if find file then do the rest
   
     ' delete existing clan trading post data
      QUERY_STRING = "DELETE * FROM TRADING_POST_GOODS"
      QUERY_STRING = QUERY_STRING & " WHERE (((TRADING_POST_GOODS.TRIBE)='"
      QUERY_STRING = QUERY_STRING & CLANNUMBER & "'));"
      Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
      qdfCurrent.Execute
    
      Import_File = CurDir$ & "\trading_posts\" & CLANNUMBER & "_Trading_Post.xls"
      Import_Table = "Import_Trading_Post_Goods"
    
      ' this will import the spreadsheet specified above into the table specified.
      ' the spreadsheet must have headings included.
   
      DoCmd.TransferSpreadsheet acImport, 8, Import_Table, Import_File, True
      Kill Import_File
   
   
   
   End If
  
ERR_File_Missing:
  TRIBECHECK.MoveNext
  If TRIBECHECK.EOF Then
     Exit Do
  End If
Loop

' Transfer the info into the Trading_Post_Goods table
Set TRADING_POST_GOODS = TVDBGM.OpenRecordset("Trading_Post_Goods")
TRADING_POST_GOODS.MoveFirst

Set Import_Trades = TVDB.OpenRecordset("Import_Trading_Post_Goods")

If Not Import_Trades.EOF Then
   Import_Trades.MoveFirst

Do While Not Import_Trades.EOF
   TRADING_POST_GOODS.AddNew
   TRADING_POST_GOODS![TYPE_OF_TRADING_POST] = "TRIBE"
   TRADING_POST_GOODS![TRIBE] = Import_Trades![TRIBE]
   TRADING_POST_GOODS![GOOD] = Import_Trades![GOOD]
   TRADING_POST_GOODS![HEX_MAP_ID] = "BA"
   TRADING_POST_GOODS![BUY PRICE] = Import_Trades![BUY PRICE]
   TRADING_POST_GOODS![BUY LIMIT] = Import_Trades![BUY LIMIT]
   TRADING_POST_GOODS![BUY_RESET_WAIT] = 0
   TRADING_POST_GOODS![NORMAL_BUY_LIMIT] = Import_Trades![BUY LIMIT]
   TRADING_POST_GOODS![TURNS_SINCE_LAST_BUY] = 0
   TRADING_POST_GOODS![BUY_THIS_TURN] = "N"
   TRADING_POST_GOODS![BUY_TOTAL] = 0
   TRADING_POST_GOODS![SELL PRICE] = Import_Trades![SELL PRICE]
   TRADING_POST_GOODS![SELL LIMIT] = Import_Trades![SELL LIMIT]
   TRADING_POST_GOODS![SELL_RESET_WAIT] = 0
   TRADING_POST_GOODS![NORMAL_SELL_LIMIT] = Import_Trades![SELL LIMIT]
   TRADING_POST_GOODS![TURNS_SINCE_LAST_SELL] = 0
   TRADING_POST_GOODS![SELL_THIS_TURN] = "N"
   TRADING_POST_GOODS![SELL_TOTAL] = 0
   TRADING_POST_GOODS.UPDATE

   Import_Trades.Delete
   Import_Trades.MoveFirst
   If Import_Trades.EOF Then
      Exit Do
   End If
Loop
End If
TRADING_POST_GOODS.Close
Import_Trades.Close
  
ERR_close:
   Exit Function

ERR_TABLES:
If (Err = 53) Then
   ' This error occurs when a file does not exist.
   
   Resume ERR_File_Missing
   
ElseIf (Err = 52) Then
   ' This error occurs when a filename is invalid.
   
   Resume ERR_File_Missing
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_close

End If

End Function

Public Function Importing_City_Spreadsheets()
TRIBE_STATUS = "Importing City Spreadsheets"

On Error GoTo ERR_TABLES
Dim Import_File As String
Dim Import_Table As String
Dim qdfCurrent As QueryDef
Dim QUERY_STRING As String

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

' loop through the clans with tp's
' look at the hex_map_const table

Set TRIBECHECK = TVDBGM.OpenRecordset("HEX_MAP_CITY")
TRIBECHECK.index = "PRIMARYKEY"
TRIBECHECK.MoveFirst

Do
   CLANNUMBER = TRIBECHECK![CITY]
   Import_File = CurDir$ & "\trading_posts\" & CLANNUMBER & "_Trading_Post.xls"
      
   'IS FILE THERE
   Open Import_File For Input As #1
   Close #1
    
   ' if find file then do the rest
  
   ' delete existing clan trading post data
   QUERY_STRING = "DELETE * FROM TRADING_POST_GOODS"
   QUERY_STRING = QUERY_STRING & " WHERE (((TRADING_POST_GOODS.TRIBE)='"
   QUERY_STRING = QUERY_STRING & CLANNUMBER & "'));"
   Set qdfCurrent = TVDB.CreateQueryDef("", QUERY_STRING)
   qdfCurrent.Execute
    
   Import_File = CurDir$ & "\trading_posts\" & CLANNUMBER & "_Trading_Post.xls"
   Import_Table = "Trading_Post_Goods"
    
   ' this will import the spreadsheet specified above into the table specified.
   ' the spreadsheet must have headings included.
   
   DoCmd.TransferSpreadsheet acImport, 8, Import_Table, Import_File, True
   Kill Import_File
  
ERR_File_Missing:
  TRIBECHECK.MoveNext
  If TRIBECHECK.EOF Then
     Exit Do
  End If
Loop

ERR_close:
   Exit Function

ERR_TABLES:
If (Err = 53) Then
   ' This error occurs when a file does not exist.
   
   Resume ERR_File_Missing
   
ElseIf (Err = 52) Then
   ' This error occurs when a filename is invalid.
   
   Resume ERR_File_Missing
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_close

End If

End Function


Public Function Importing_Player_Orders()
'The load procedure will read all emails in the Import folder of Outlook and then present the list to the user for selection
Dim objApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim colFolders As Outlook.Folders
Dim objFolder As Outlook.MAPIFolder
Dim objParentFolder As Outlook.MAPIFolder
Dim objMailItem As Outlook.MailItem
Dim OlImportedFolder As Outlook.MAPIFolder
Dim strSQL As String
Dim strBody As String
Dim strSubject As String
Const strDoubleQuote As String = """"
Dim strFolderPath As String
Dim i As Byte
Dim count As Integer

On Error GoTo Import_Error
'I use EmptyTable to run SQL code “Delete from tblImportEmails”. Search for EmptyTable below for the code
EmptyTable "tblImportEmails"
Set objApp = New Outlook.Application
Set objNS = objApp.GetNamespace("MAPI")
For i = 1 To objNS.Folders.count
   Set objParentFolder = objNS.Folders(i)
   'I use OutlookFolderNames to locate the folder in the user’s Outlook
   Set objFolder = OutlookFolderNames(objParentFolder, "Import")
   Set OlImportedFolder = OutlookFolderNames(objParentFolder, "Imported")
  
   'Set OlImportedFolder = objParentFolder.Folders("Imported")
  
   If Not objFolder Is Nothing Then
      'Once the folder is located exit For
      Exit For
   End If
   'what happens when there is no folder in outlook? code stops and here you need to check if user has Import folder in their Outlook
Next i

If objFolder Is Nothing Then
   MsgBox "Please make sure you have two folders in your Outlook:" & vbCrLf & "Import" & vbCrLf & _
   "Imported.", vbInformation, "Missing Folders in Outlook"
   Exit Function
End If

If Not objFolder Is Nothing Then
   For Each objMailItem In objFolder.Items
   With objMailItem
   'I build the SQL statement to save email details in temp folder, notice the user of replace to take care of single
   'quotes and other issues.
  
   strSQL = "From: " & .SenderName & " (" & Replace(.SenderEmailAddress, "'", "''") & ")" & vbCrLf
   strSQL = strSQL & "To: " & Replace(.To, "'", "''") & vbCrLf
   strSQL = strSQL & "CC: " & Replace(.CC, "'", "''") & vbCrLf
   strSubject = Replace(.Subject, "'", "''")
   'strSubject = Replace(strSubject, Chr(34), Chr(34) & Chr(34))
   strSQL = strSQL & "Subject: " & strSubject & vbCrLf
   strSQL = strSQL & "Date Emailed: " & .SentOn & vbCrLf
   strBody = Replace(.Body, "'", "''")
   strBody = Replace(strBody, Chr(34), Chr(34) & Chr(34))
   strSQL = strSQL & strBody
   strSQL = "Insert Into tblImportEmails(EmailFrom, EmailTo, EmailCC, EmailSubject, EmailBody, EmailDate, Message,OutlookID) " & _
   "Values('" & .SenderEmailAddress & "', '" & Replace(.To, "'", "''") & "','" & Replace(.CC, "'", "''") & "','" & strSubject & _
   "','" & strBody & "',#" & _
   .SentOn & "#,'" & strSQL & "','" & .EntryID & "')"
   CurrentDb.Execute strSQL
   End With

   Next
End If

Set objFolder = Nothing
Set objNS = Nothing
Set objApp = Nothing

Call Clear_Import_Folder

ExitProcedure:
On Error GoTo 0
Exit Function

Import_Error:
MsgBox "Error " & Err.NUMBER & " (" & Err.Description & ") in procedure Importing_Players_Orders"

End Function



Public Function EmptyTable(strTable As String)
'---------------------------------------------------------------------------------------
' Procedure : EmptyTable
' DateTime : 7/20/2006 14:20
' Author : Juan Soto/ AccessExperts.net/blog
' Purpose : Empty the table in the database represented by strTable. Usually used for temporary tables.
'---------------------------------------------------------------------------------------
'REVISED
'Date By Comment
'
Dim strSQL As String
On Error GoTo EmptyTable_Error
strSQL = "Delete * From " & strTable
CurrentDb.Execute strSQL, dbSeeChanges
On Error GoTo 0
Exit Function
EmptyTable_Error:
End Function
'Place this function in a global module




Public Function Clear_Import_Folder()
Dim objApp As Outlook.Application
Dim objNS As Outlook.NameSpace
Dim colFolders As Outlook.Folders
Dim objFolder As Outlook.MAPIFolder
Dim objParentFolder As Outlook.MAPIFolder
Dim objMailItem As Outlook.MailItem
Dim OlImportedFolder As Outlook.MAPIFolder
Dim strSQL As String
Dim strBody As String
Dim strSubject As String
Const strDoubleQuote As String = """"
Dim strFolderPath As String
Dim count As Integer
Dim i As Byte
On Error GoTo Import_Error
Set objApp = New Outlook.Application
Set objNS = objApp.GetNamespace("MAPI")
For i = 1 To objNS.Folders.count
   Set objParentFolder = objNS.Folders(i)
   'I use OutlookFolderNames to locate the folder in the user’s Outlook
   Set objFolder = OutlookFolderNames(objParentFolder, "Import")
   Set OlImportedFolder = OutlookFolderNames(objParentFolder, "Imported")
    
   If Not objFolder Is Nothing Then
      'Once the folder is located exit For
      Exit For
   End If
   'what happens when there is no folder in outlook? code stops and here you need to check if user has Import folder in their Outlook
Next i

If objFolder Is Nothing Then
   MsgBox "Please make sure you have two folders in your Outlook:" & vbCrLf & "Import" & vbCrLf & _
   "Imported.", vbInformation, "Missing Folders in Outlook"
   Exit Function
End If

If Not objFolder Is Nothing Then
count = objFolder.Items.count
   Do While count > 0
       objFolder.Items(count).Move OlImportedFolder
       count = count - 1
   Loop
End If

Set objFolder = Nothing
Set objNS = Nothing
Set objApp = Nothing

ExitProcedure:
On Error GoTo 0
Exit Function

Import_Error:
MsgBox "Error " & Err.NUMBER & " (" & Err.Description & ") in procedure Importing_Players_Orders"

End Function



Private Function OutlookFolderNames(objFolder As Outlook.MAPIFolder, strFolderName As String) As Object
'*********************************************************
On Error GoTo ErrorHandler
Dim objOneSubFolder As Outlook.MAPIFolder
Dim oFolder As Outlook.Folder

If Not objFolder Is Nothing Then
   If LCase(strFolderName) = LCase(objFolder.Name) Then
      Set OutlookFolderNames = objFolder
   Else
      ' Check if folders collection is not empty
      If objFolder.Folders.count > 0 And Not objFolder.Folders Is Nothing Then
         For Each oFolder In objFolder.Folders
         Set objOneSubFolder = oFolder
         ' only check mail item folder
         If objOneSubFolder.DefaultItemType = olMailItem Then
            If LCase(strFolderName) = LCase(objOneSubFolder.Name) Then
               Set OutlookFolderNames = objOneSubFolder
               Exit For
            Else
               If objOneSubFolder.Folders.count > 0 Then
                  Set OutlookFolderNames = OutlookFolderNames(objOneSubFolder, strFolderName)
                  If Not (OutlookFolderNames Is Nothing) Then
                     'MsgBox "It worked!"
                     Exit Function
                  End If
               End If
            End If
         End If
         Next
      End If
   End If
End If

Exit Function
ErrorHandler:
   Set OutlookFolderNames = Nothing
End Function


Private Function Parse_Imported_Orders()
'*********************************************************
'On Error GoTo ExitProcedure
Dim IMPORTEDEMAILS As Recordset
Dim EMAILLINE As String
Dim LENGTH As Integer
Dim START_POSITION As Integer
Dim NEXT_POSITION As Integer
Dim END_POSITION As Integer
Dim FIND_COMMA As Integer
Dim From_Clan As String
Dim From_Tribe As String
Dim To_Clan As String
Dim To_Tribe As String
Dim GOOD As String
Dim AMOUNT As String
Dim TVFILE As String

Set TVWKSPACE = DBEngine.Workspaces(0)

TVFILE = CurDir$ & "\TV.accdb"

Set TVDB = TVWKSPACE.OpenDatabase(TVFILE, False, False)
'Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set IMPORTEDEMAILS = TVDB.OpenRecordset("tblImportEmails")
IMPORTEDEMAILS.MoveFirst




' loop through all emails


' read each line in email into a string
START_POSITION = 1
END_POSITION = InStr(IMPORTEDEMAILS![EMAILBODY], vbCrLf)

EMAILLINE = Mid(IMPORTEDEMAILS![EMAILBODY], START_POSITION, END_POSITION)

'determine valid email and clan
If InStr(EMAILLINE, "ILOVEMONKEYS") Then
   'Find Start in first line and extract Clan
   From_Clan = "0" & Mid(IMPORTEDEMAILS![EMAILBODY], InStr(IMPORTEDEMAILS![EMAILBODY], "START") - 4, 3)
   
   'Transfers first
   
   If InStr(EMAILLINE, "TRANSFER FROM:") > 0 Then
      START_POSITION = InStr(EMAILLINE, "TRANSFER FROM:") + 14
      If Mid(EMAILLINE, START_POSITION, 1) = " " Then
         START_POSITION = START_POSITION + 1
      End If
      FIND_COMMA = InStr(Mid(EMAILLINE, START_POSITION, 10), ",")
      
      ' Determine From Tribe
      From_Tribe = Mid(EMAILLINE, START_POSITION, FIND_COMMA - 1)
      If Len(From_Tribe) = 3 Or Len(From_Tribe) = 5 Then
         From_Tribe = "0" & From_Tribe
      End If
      
      'Determine To Tribe
      START_POSITION = START_POSITION + FIND_COMMA
      NEXT_POSITION = START_POSITION + InStr(Mid(EMAILLINE, START_POSITION, 10), "TO:") + 3
      FIND_COMMA = InStr(Mid(EMAILLINE, NEXT_POSITION, 10), ",")
      To_Tribe = Mid(EMAILLINE, NEXT_POSITION, FIND_COMMA - 1)
      If Len(To_Tribe) = 3 Or Len(To_Tribe) = 5 Then
         To_Tribe = "0" & To_Tribe
      End If
   
      'DETERMINE TO_CLAN
      If InStr(To_Tribe, From_Clan) > 0 Then
         To_Clan = From_Clan
      ElseIf Len(To_Tribe) = 3 Then
         To_Clan = "0" & To_Tribe
      ElseIf Len(To_Tribe) = 5 Then
         To_Clan = "0" & Mid(To_Tribe, 1, 3)
      ElseIf Len(To_Tribe) = 6 Then
         To_Clan = "0" & Mid(To_Tribe, 2, 3)
      End If
      
      'RESET STRING
      'EMAILLINE = Mid(EMAILLINE, 1, 1)
      'FIND GOOD
      START_POSITION = FIND_COMMA + NEXT_POSITION + 1
      FIND_COMMA = InStr(Mid(EMAILLINE, START_POSITION, 10), ",")
      GOOD = Mid(EMAILLINE, START_POSITION, FIND_COMMA - 1)
      'FIND AMOUNT
      START_POSITION = FIND_COMMA + START_POSITION + 1
      FIND_COMMA = InStr(Mid(EMAILLINE, START_POSITION, 10), ",")
      If FIND_COMMA = 0 Then
         ' HOW TO FIND END OF LINE
         FIND_COMMA = InStr(Mid(EMAILLINE, START_POSITION, 10), ";")
         If FIND_COMMA = 0 Then
            'move forwards one step at a time until there isn't a integer
            
         Else
            AMOUNT = Mid(EMAILLINE, START_POSITION, FIND_COMMA - 1)
         End If
      Else
         AMOUNT = Mid(IMPORTEDEMAILS![EMAILBODY], START_POSITION, FIND_COMMA - 1)
      End If
      
    
      ' next transfer or next action
                 
                 
   End If
   
Else
   'go to next email
End If
'TRIBENET CLAN 330 START ILOVEMONKEYS
'TRANSFER FROM: 330E9, TO: 1330E1, WAGON, 5;
'ACTION FROM: 330,DEFENCE, USING: BOWS, 100;SPEARS,100;
'ACTION FROM: 330,HUNTING, USING: TRAPS. 1000, SNARES, 1000, BOWS, 100;
'MOVE FROM: 330, TO: N x 1, NE x 1
'SCOUT1 FROM: 330, TO: N x 1, NE x 1
'SCOUT2 FROM: 330, TO: NE x 1, N x 1
'TRIBENET CLAN 330 END

'LENGTH = Len(IMPORTEDEMAILS![EMAILBODY])

'Transfers first
'find start of line


'END_POSITION = InStr(IMPORTEDEMAILS![EMAILBODY], "End Transfers")
   
'   WORDLEN = Len(TRIBE_RESEARCH![TOPIC])
'   If Right(Mid(TRIBE_RESEARCH![TOPIC], 1, WORDLEN), 1) = " " Then
'      WORDLEN = WORDLEN - 1
'   End If
'   SEARCHVALUE = Mid(TRIBE_RESEARCH![TOPIC], 1, WORDLEN)
'   RESEARCH_TABLE.Seek "=", SEARCHVALUE






ExitProcedure:
On Error GoTo 0
Exit Function

Import_Error:
MsgBox "Error " & Err.NUMBER & " (" & Err.Description & ") in procedure Importing_Players_Orders"

End Function


Sub ImportTransfers(fileName As String, tableName As String)
    DebugOP "Importing>ImportTransfers: " & fileName & " " & tableName

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, tableName, fileName, True

End Sub

 






































