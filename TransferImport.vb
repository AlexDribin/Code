Attribute VB_Name = "TransferImport"
Option Compare Database

Sub ImportTransfers(fileName As String)
On Error GoTo ERR_ImportTransfers
Dim qdfCurrent As QueryDef
Dim TRIBE_FIRST As String
Dim CURRENT_TRIBE As String
Dim CURRENT_TO As String
Dim fName As String

DebugOP "TransferImport>ImportTransfers: " & fileName

fName = GetFileNameFromPath(fileName)

Set TVMWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVMWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst
 
FILEGM = CurDir$ & "\" & GMTABLE![FILE]
 
Set TVDBGM = TVMWKSPACE.OpenDatabase(FILEGM, False, False)
     
Set TRIBESINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESINFO.MoveFirst
TRIBESINFO.index = "PRIMARYKEY"

'=====================TRANSFERS========================
DebugOP "Importing Transfers......." & fName
Forms![IMPORT_TRANSFERS]![Status] = "Beginning Import "
Forms![IMPORT_TRANSFERS].Repaint

' clear the copy file
Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM MASSTRANSFERS_copy;")
qdfCurrent.Execute

Forms![IMPORT_TRANSFERS]![Status] = "Import Transfers"
Forms![IMPORT_TRANSFERS].Repaint

' pull data from excel and put into copy
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "MASSTRANSFERS_copy", fileName, True, "Transfers!A1:G1000"
    
Call CLEAN_UP_BLANK_ROWS("TRANSFERS_COPY")

Forms![IMPORT_TRANSFERS]![Status] = "Move Transfers"
Forms![IMPORT_TRANSFERS].Repaint

'move copy data into table
Set MASS_TRANSFER_COPY_TABLE = TVDB.OpenRecordset("MASSTRANSFERS_copy")
If MASS_TRANSFER_COPY_TABLE.EOF Then
   ' no transfers
   GoTo Tribe_Movement
Else
   MASS_TRANSFER_COPY_TABLE.MoveFirst
End If

Set Masstransfers = TVDBGM.OpenRecordset("MassTransfers")
If Masstransfers.RecordCount > 0 Then
   Masstransfers.MoveFirst
Else
   GoTo After_Transfer_Delete
End If

'clean up duplicates
Do Until MASS_TRANSFER_COPY_TABLE.EOF
   Masstransfers.MoveFirst
   Do Until Masstransfers.EOF
      If MASS_TRANSFER_COPY_TABLE!From = Masstransfers!From _
      And MASS_TRANSFER_COPY_TABLE![To] = Masstransfers!To _
      And MASS_TRANSFER_COPY_TABLE![ITEM] = Masstransfers!ITEM _
      And MASS_TRANSFER_COPY_TABLE![QUANTITY] = CStr(Masstransfers!QUANTITY) Then
          Masstransfers.Delete
          Exit Do
      End If
      Masstransfers.MoveNext
      If Masstransfers.EOF Then
         Exit Do
      End If
   Loop
   MASS_TRANSFER_COPY_TABLE.MoveNext
   If MASS_TRANSFER_COPY_TABLE.EOF Then
      Exit Do
   End If
Loop

After_Transfer_Delete:

MASS_TRANSFER_COPY_TABLE.MoveFirst
CURRENT_TRIBE = ""
CURRENT_TO = ""
TRIBE_FIRST = "YES"
Do Until MASS_TRANSFER_COPY_TABLE.EOF
   If MASS_TRANSFER_COPY_TABLE![PROCESSED] = "N" Or IsNull(MASS_TRANSFER_COPY_TABLE![PROCESSED]) Then
      If IsNull(MASS_TRANSFER_COPY_TABLE!From) Then
         ' ignore
      Else
         Masstransfers.AddNew
         Masstransfers![From] = MASS_TRANSFER_COPY_TABLE![From]
         Masstransfers![To] = MASS_TRANSFER_COPY_TABLE![To]
         Masstransfers![ITEM] = MASS_TRANSFER_COPY_TABLE![ITEM]
         If IsNull(MASS_TRANSFER_COPY_TABLE![QUANTITY]) Then
            Masstransfers![QUANTITY] = 0
         Else
            Masstransfers![QUANTITY] = MASS_TRANSFER_COPY_TABLE![QUANTITY]
         End If
         Masstransfers![TRANSFER_TIMING] = MASS_TRANSFER_COPY_TABLE![TRANSFER_TIMING]
         Masstransfers![NOTES] = MASS_TRANSFER_COPY_TABLE![NOTES]
         Masstransfers![PROCESSED] = "N"
         Masstransfers.UPDATE
      End If
   End If
   MASS_TRANSFER_COPY_TABLE.Edit
   MASS_TRANSFER_COPY_TABLE![PROCESSED] = "Y"
   MASS_TRANSFER_COPY_TABLE.UPDATE
   MASS_TRANSFER_COPY_TABLE.MoveNext
   If MASS_TRANSFER_COPY_TABLE.EOF Then
      Exit Do
   End If
Loop

'=====================MOVEMENT========================
Tribe_Movement:
DebugOP "Importing Movement........" & fName
Forms![IMPORT_TRANSFERS]![Status] = "Import Tribe Movement"
Forms![IMPORT_TRANSFERS].Repaint

    Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Process_Tribe_Movement_Copy;")
    qdfCurrent.Execute

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Process_Tribe_Movement_Copy", fileName, True, "Tribe_Movement!A1:AH1000"
    
Call CLEAN_UP_BLANK_ROWS("TRIBE_MOVEMENT_COPY")

Forms![IMPORT_TRANSFERS]![Status] = "Move Tribe Movement Copy"
Forms![IMPORT_TRANSFERS].Repaint

'move copy data into table
Set Process_Tribe_Movement_Copy = TVDB.OpenRecordset("Process_Tribe_Movement_Copy")
Process_Tribe_Movement_Copy.index = "PRIMARYKEY"
If Process_Tribe_Movement_Copy.EOF Then
   ' no Movement
   GoTo Scouting
Else
   Process_Tribe_Movement_Copy.MoveFirst
End If

Set Process_Tribe_Movement = TVDBGM.OpenRecordset("Process_Tribe_Movement")
Process_Tribe_Movement.index = "PRIMARYKEY"
Process_Tribe_Movement.MoveFirst
   
CURRENT_TRIBE = ""
    TRIBE_FIRST = "YES"
    If Not Process_Tribe_Movement_Copy.NoMatch Then
       Do Until Process_Tribe_Movement_Copy.EOF
          If CURRENT_TRIBE = Process_Tribe_Movement_Copy![TRIBE] Then
             'DO NOTHING
          Else
             If IsNull(Process_Tribe_Movement_Copy![TRIBE]) Then
                Exit Do
             Else
                CURRENT_TRIBE = Process_Tribe_Movement_Copy![TRIBE]
                TRIBE_FIRST = "YES"
             End If
          End If
          
          If TRIBE_FIRST = "YES" Then
             Process_Tribe_Movement.Seek "=", Process_Tribe_Movement_Copy!TRIBE
             If Process_Tribe_Movement.NoMatch Then
                'DO NOTHING
             Else
                Do
                  Process_Tribe_Movement.Delete
                  Process_Tribe_Movement.MoveNext
                  If Process_Tribe_Movement.EOF Then
                     Exit Do
                  End If
                  If Not Process_Tribe_Movement![TRIBE] = Process_Tribe_Movement_Copy![TRIBE] Then
                      Exit Do
                  End If
                  
                Loop
             End If
             TRIBE_FIRST = "NO"
          End If
       If Process_Tribe_Movement_Copy![PROCESSED] = "N" Or IsNull(Process_Tribe_Movement_Copy![PROCESSED]) Then
          If IsNull(Process_Tribe_Movement_Copy!TRIBE) Then
        ' ignore
          Else
              Process_Tribe_Movement.AddNew
              Process_Tribe_Movement![TRIBE] = Process_Tribe_Movement_Copy![TRIBE]
              Process_Tribe_Movement![Follow_Tribe] = Process_Tribe_Movement_Copy![Follow_Tribe]
              Process_Tribe_Movement![MOVEMENT_1] = Process_Tribe_Movement_Copy![MOVEMENT_1]
              Process_Tribe_Movement![MOVEMENT_2] = Process_Tribe_Movement_Copy![MOVEMENT_2]
              Process_Tribe_Movement![MOVEMENT_3] = Process_Tribe_Movement_Copy![MOVEMENT_3]
              Process_Tribe_Movement![MOVEMENT_4] = Process_Tribe_Movement_Copy![MOVEMENT_4]
              Process_Tribe_Movement![MOVEMENT_5] = Process_Tribe_Movement_Copy![MOVEMENT_5]
              Process_Tribe_Movement![MOVEMENT_6] = Process_Tribe_Movement_Copy![MOVEMENT_6]
              Process_Tribe_Movement![MOVEMENT_7] = Process_Tribe_Movement_Copy![MOVEMENT_7]
              Process_Tribe_Movement![MOVEMENT_8] = Process_Tribe_Movement_Copy![MOVEMENT_8]
              Process_Tribe_Movement![MOVEMENT_9] = Process_Tribe_Movement_Copy![MOVEMENT_9]
              Process_Tribe_Movement![MOVEMENT_10] = Process_Tribe_Movement_Copy![MOVEMENT_10]
              Process_Tribe_Movement![MOVEMENT_11] = Process_Tribe_Movement_Copy![MOVEMENT_11]
              Process_Tribe_Movement![MOVEMENT_12] = Process_Tribe_Movement_Copy![MOVEMENT_12]
              Process_Tribe_Movement![MOVEMENT_13] = Process_Tribe_Movement_Copy![MOVEMENT_13]
              Process_Tribe_Movement![MOVEMENT_14] = Process_Tribe_Movement_Copy![MOVEMENT_14]
              Process_Tribe_Movement![MOVEMENT_15] = Process_Tribe_Movement_Copy![MOVEMENT_15]
              Process_Tribe_Movement![MOVEMENT_16] = Process_Tribe_Movement_Copy![MOVEMENT_16]
              Process_Tribe_Movement![MOVEMENT_17] = Process_Tribe_Movement_Copy![MOVEMENT_17]
              Process_Tribe_Movement![MOVEMENT_18] = Process_Tribe_Movement_Copy![MOVEMENT_18]
              Process_Tribe_Movement![MOVEMENT_19] = Process_Tribe_Movement_Copy![MOVEMENT_19]
              Process_Tribe_Movement![MOVEMENT_20] = Process_Tribe_Movement_Copy![MOVEMENT_20]
              Process_Tribe_Movement![MOVEMENT_21] = Process_Tribe_Movement_Copy![MOVEMENT_21]
              Process_Tribe_Movement![MOVEMENT_22] = Process_Tribe_Movement_Copy![MOVEMENT_22]
              Process_Tribe_Movement![MOVEMENT_23] = Process_Tribe_Movement_Copy![MOVEMENT_23]
              Process_Tribe_Movement![MOVEMENT_24] = Process_Tribe_Movement_Copy![MOVEMENT_24]
              Process_Tribe_Movement![MOVEMENT_25] = Process_Tribe_Movement_Copy![MOVEMENT_25]
              Process_Tribe_Movement![MOVEMENT_26] = Process_Tribe_Movement_Copy![MOVEMENT_26]
              Process_Tribe_Movement![MOVEMENT_27] = Process_Tribe_Movement_Copy![MOVEMENT_27]
              Process_Tribe_Movement![MOVEMENT_28] = Process_Tribe_Movement_Copy![MOVEMENT_28]
              Process_Tribe_Movement![MOVEMENT_29] = Process_Tribe_Movement_Copy![MOVEMENT_29]
              Process_Tribe_Movement![MOVEMENT_30] = Process_Tribe_Movement_Copy![MOVEMENT_30]
              Process_Tribe_Movement![PROCESSED] = "N"
              If Mid(Process_Tribe_Movement_Copy![HEX], 3, 1) = " " Then
                 Process_Tribe_Movement![HEX] = Process_Tribe_Movement_Copy![HEX]
              Else
                 Process_Tribe_Movement![HEX] = Mid(Process_Tribe_Movement_Copy![HEX], 1, 2) & " " & Mid(Process_Tribe_Movement_Copy![HEX], 3, 4)
              End If
              Process_Tribe_Movement.UPDATE
          End If
       End If
       Process_Tribe_Movement_Copy.Edit
       Process_Tribe_Movement_Copy![PROCESSED] = "Y"
       Process_Tribe_Movement_Copy.UPDATE
       Process_Tribe_Movement_Copy.MoveNext
       If Process_Tribe_Movement_Copy.EOF Then
          Exit Do
       End If
    Loop
    End If

'=====================SCOUTING========================
Scouting:
DebugOP "Importing Scouting........" & fName
Forms![IMPORT_TRANSFERS]![Status] = "Import Scout Movement Copy"
Forms![IMPORT_TRANSFERS].Repaint

    Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM scout_movement_copy;")
    qdfCurrent.Execute

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Scout_Movement_Copy", fileName, True, "Scout_Movement!A1:P1000"

Call CLEAN_UP_BLANK_ROWS("SCOUT_MOVEMENT")

Forms![IMPORT_TRANSFERS]![Status] = "Transfer Scout Movement Copy"
Forms![IMPORT_TRANSFERS].Repaint

    'now to transfer from the copy to the original

    Set SCOUT_MOVEMENT_TABLE = TVDBGM.OpenRecordset("SCOUT_MOVEMENT")
    SCOUT_MOVEMENT_TABLE.index = "PRIMARYKEY"
    
    Set SCOUT_MOVEMENT_COPY = TVDB.OpenRecordset("SCOUT_MOVEMENT_COPY")
    If SCOUT_MOVEMENT_COPY.EOF Then
       ' no transfers
       GoTo Activities
    Else
       SCOUT_MOVEMENT_COPY.MoveFirst
    End If
    CURRENT_TRIBE = ""
    TRIBE_FIRST = "YES"
    If Not SCOUT_MOVEMENT_COPY.NoMatch Then
       Do Until SCOUT_MOVEMENT_COPY.EOF
          If CURRENT_TRIBE = SCOUT_MOVEMENT_COPY![TRIBE] Then
             'DO NOTHING
          Else
             If IsNull(SCOUT_MOVEMENT_COPY![TRIBE]) Then
                Exit Do
             Else
                CURRENT_TRIBE = SCOUT_MOVEMENT_COPY![TRIBE]
                TRIBE_FIRST = "YES"
             End If
          End If
          If TRIBE_FIRST = "YES" Then
             SCOUT_MOVEMENT_TABLE.Seek "=", SCOUT_MOVEMENT_COPY!TRIBE
             If SCOUT_MOVEMENT_TABLE.NoMatch Then
                'DO NOTHING
             Else
                Do
                  SCOUT_MOVEMENT_TABLE.Delete
                  SCOUT_MOVEMENT_TABLE.MoveNext
                  If SCOUT_MOVEMENT_TABLE.EOF Then
                     Exit Do
                  End If
                  If Not SCOUT_MOVEMENT_TABLE![TRIBE] = SCOUT_MOVEMENT_COPY![TRIBE] Then
                      Exit Do
                  End If
                  
                Loop
             End If
          TRIBE_FIRST = "NO"
       End If
       If SCOUT_MOVEMENT_COPY![PROCESSED] = "N" Or IsNull(SCOUT_MOVEMENT_COPY![PROCESSED]) Then
          If IsNull(SCOUT_MOVEMENT_COPY!TRIBE) Then
        ' ignore
          Else
              SCOUT_MOVEMENT_TABLE.AddNew
              SCOUT_MOVEMENT_TABLE![TRIBE] = SCOUT_MOVEMENT_COPY![TRIBE]
              SCOUT_MOVEMENT_TABLE![No_of_Scouts] = SCOUT_MOVEMENT_COPY![No_of_Scouts]
              SCOUT_MOVEMENT_TABLE![No_of_Horses] = SCOUT_MOVEMENT_COPY![No_of_Horses]
              SCOUT_MOVEMENT_TABLE![No_of_Elephants] = SCOUT_MOVEMENT_COPY![No_of_Elephants]
              SCOUT_MOVEMENT_TABLE![No_of_Camels] = SCOUT_MOVEMENT_COPY![No_of_Camels]
              SCOUT_MOVEMENT_TABLE![MISSION] = SCOUT_MOVEMENT_COPY![MISSION]
              SCOUT_MOVEMENT_TABLE![Movement1] = SCOUT_MOVEMENT_COPY![Movement1]
              SCOUT_MOVEMENT_TABLE![Movement2] = SCOUT_MOVEMENT_COPY![Movement2]
              SCOUT_MOVEMENT_TABLE![Movement3] = SCOUT_MOVEMENT_COPY![Movement3]
              SCOUT_MOVEMENT_TABLE![Movement4] = SCOUT_MOVEMENT_COPY![Movement4]
              SCOUT_MOVEMENT_TABLE![Movement5] = SCOUT_MOVEMENT_COPY![Movement5]
              SCOUT_MOVEMENT_TABLE![Movement6] = SCOUT_MOVEMENT_COPY![Movement6]
              SCOUT_MOVEMENT_TABLE![Movement7] = SCOUT_MOVEMENT_COPY![Movement7]
              SCOUT_MOVEMENT_TABLE![Movement8] = SCOUT_MOVEMENT_COPY![Movement8]
              SCOUT_MOVEMENT_TABLE![Movement9] = SCOUT_MOVEMENT_COPY![Movement9]
              SCOUT_MOVEMENT_TABLE![PROCESSED] = "N"
              SCOUT_MOVEMENT_TABLE.UPDATE
          End If
       End If
       SCOUT_MOVEMENT_COPY.Edit
       SCOUT_MOVEMENT_COPY![PROCESSED] = "Y"
       SCOUT_MOVEMENT_COPY.UPDATE
       SCOUT_MOVEMENT_COPY.MoveNext
       If SCOUT_MOVEMENT_COPY.EOF Then
          Exit Do
       End If
    Loop
    End If
    
'=====================ACTIVITIES========================

Activities:
DebugOP "Importing Activities......" & fName
Forms![IMPORT_TRANSFERS]![Status] = "Import Tribe_Activities"
Forms![IMPORT_TRANSFERS].Repaint

    Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM TRIBE_ACTIVITIES_IMPLEMENTS;")
    qdfCurrent.Execute

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "TRIBE_ACTIVITIES_IMPLEMENTS", fileName, True, "Tribes_Activities!A1:BA1000"

Call CLEAN_UP_BLANK_ROWS("TRIBES_ACTIVITY")

Forms![IMPORT_TRANSFERS]![Status] = "Transfers Tribe_Activities"
Forms![IMPORT_TRANSFERS].Repaint

    Set PROCESSACTIVITY = TVDBGM.OpenRecordset("Process_Tribes_Activity")
    PROCESSACTIVITY.index = "PRIMARYKEY"
  
    Set PROCESSITEMS = TVDBGM.OpenRecordset("Process_Tribes_Item_Allocation")
    PROCESSITEMS.index = "PRIMARYKEY"
   
    ' Now load the data from Tribes_Activity_Implements into Process_Tribes_Activtities and Process_Tribes_Item_Allocation

    Set TRIBES_ACTIVITY_IMPLEMENTS = TVDB.OpenRecordset("TRIBE_ACTIVITIES_IMPLEMENTS")
    If TRIBES_ACTIVITY_IMPLEMENTS.EOF Then
       ' no transfers
       GoTo SKILLS
    Else
       TRIBES_ACTIVITY_IMPLEMENTS.MoveFirst
    End If
    TRIBE_FIRST = "YES"
    CURRENT_TRIBE = "EMPTY"
    If Not TRIBES_ACTIVITY_IMPLEMENTS.NoMatch Then
       Do Until TRIBES_ACTIVITY_IMPLEMENTS.EOF
          If IsNull(TRIBES_ACTIVITY_IMPLEMENTS!TRIBE) Then
             GoTo END_TRIBE_ACTIVITIES_IMPLEMENTS
          End If
          If CURRENT_TRIBE = TRIBES_ACTIVITY_IMPLEMENTS!TRIBE Then
             'IGNORE
          Else
             TRIBE_FIRST = "YES"
             CURRENT_TRIBE = TRIBES_ACTIVITY_IMPLEMENTS!TRIBE
          End If
          If TRIBE_FIRST = "YES" Then
             PROCESSACTIVITY.Seek "=", TRIBES_ACTIVITY_IMPLEMENTS!TRIBE
             If PROCESSACTIVITY.NoMatch Then
                'DO NOTHING
             Else
                Do
                  PROCESSACTIVITY.Delete
                  PROCESSACTIVITY.MoveNext
                  If PROCESSACTIVITY.EOF Then
                     Exit Do
                  End If
                  If Not PROCESSACTIVITY![TRIBE] = TRIBES_ACTIVITY_IMPLEMENTS![TRIBE] Then
                      Exit Do
                  End If
                  
                Loop
             End If
          End If
          If TRIBE_FIRST = "YES" Then
             PROCESSITEMS.Seek "=", TRIBES_ACTIVITY_IMPLEMENTS!TRIBE
             If PROCESSITEMS.NoMatch Then
                'DO NOTHING
             Else
                Do
                  PROCESSITEMS.Delete
                  PROCESSITEMS.MoveNext
                  If PROCESSITEMS.EOF Then
                     Exit Do
                  End If
                  If Not PROCESSITEMS![TRIBE] = TRIBES_ACTIVITY_IMPLEMENTS![TRIBE] Then
                     Exit Do
                  End If
                 
                Loop
             End If
             TRIBE_FIRST = "NO"
          End If
       If IsNull(TRIBES_ACTIVITY_IMPLEMENTS![PROCESSED]) Or TRIBES_ACTIVITY_IMPLEMENTS![PROCESSED] = "N" Then
          If IsNull(TRIBES_ACTIVITY_IMPLEMENTS!TRIBE) Then
             ' ignore
          Else
              PROCESSACTIVITY.AddNew
              PROCESSACTIVITY![TRIBE] = TRIBES_ACTIVITY_IMPLEMENTS![TRIBE]
              PROCESSACTIVITY![ACTIVITY] = TRIBES_ACTIVITY_IMPLEMENTS![ACTIVITY]
              PROCESSACTIVITY![ITEM] = TRIBES_ACTIVITY_IMPLEMENTS![ITEM]
              PROCESSACTIVITY![DISTINCTION] = TRIBES_ACTIVITY_IMPLEMENTS![DISTINCTION]
              PROCESSACTIVITY![PEOPLE] = TRIBES_ACTIVITY_IMPLEMENTS![PEOPLE]
              PROCESSACTIVITY![Slaves] = TRIBES_ACTIVITY_IMPLEMENTS![Slaves]
              PROCESSACTIVITY![SPECIALISTS] = TRIBES_ACTIVITY_IMPLEMENTS![SPECIALISTS]
              PROCESSACTIVITY![JOINT] = TRIBES_ACTIVITY_IMPLEMENTS![JOINT]
              
              If Left(TRIBES_ACTIVITY_IMPLEMENTS![OWNING_TRIBE], 2) = "  " Then
                 PROCESSACTIVITY![OWNING_TRIBE] = ""
              Else
                 PROCESSACTIVITY![OWNING_TRIBE] = TRIBES_ACTIVITY_IMPLEMENTS![OWNING_TRIBE]
              End If
              
              PROCESSACTIVITY![Number_of_Seeking_Groups] = 0
              PROCESSACTIVITY![Whale_Size] = TRIBES_ACTIVITY_IMPLEMENTS![Whale_Size]
              PROCESSACTIVITY![MINING_DIRECTION] = TRIBES_ACTIVITY_IMPLEMENTS![MINING_DIRECTION]
              PROCESSACTIVITY![PROCESSED] = "N"
              PROCESSACTIVITY![Building] = TRIBES_ACTIVITY_IMPLEMENTS![Building]
              PROCESSACTIVITY.UPDATE
              count = 1
              Do While count < 21
                 stext1 = "[IMPLEMENT_" & CStr(count) & "]"
                 stext2 = "[IMP_" & CStr(count) & "_NUMBER]"
                 If IsNull(TRIBES_ACTIVITY_IMPLEMENTS(stext1).Value) Then
                    Exit Do
                 End If
                 PROCESSITEMS.AddNew
                 PROCESSITEMS![TRIBE] = TRIBES_ACTIVITY_IMPLEMENTS![TRIBE]
                 PROCESSITEMS![ACTIVITY] = TRIBES_ACTIVITY_IMPLEMENTS![ACTIVITY]
                 PROCESSITEMS![ITEM] = TRIBES_ACTIVITY_IMPLEMENTS![ITEM]
                 PROCESSITEMS![ITEM_USED] = TRIBES_ACTIVITY_IMPLEMENTS(stext1).Value
                 PROCESSITEMS![QUANTITY] = TRIBES_ACTIVITY_IMPLEMENTS(stext2).Value
                 PROCESSITEMS![PROCESSED] = "N"
                 PROCESSITEMS.UPDATE
                 count = count + 1
                 If count = 21 Then
                    Exit Do
                 End If
              Loop
          End If
       End If
END_TRIBE_ACTIVITIES_IMPLEMENTS:
       TRIBES_ACTIVITY_IMPLEMENTS.Edit
       TRIBES_ACTIVITY_IMPLEMENTS![PROCESSED] = "Y"
       TRIBES_ACTIVITY_IMPLEMENTS.UPDATE
       TRIBES_ACTIVITY_IMPLEMENTS.MoveNext
       If TRIBES_ACTIVITY_IMPLEMENTS.EOF Then
          Exit Do
       End If
    Loop
    End If
   
'=====================SKILLS========================
SKILLS:
DebugOP "Importing Skills.........." & fName
Forms![IMPORT_TRANSFERS]![Status] = "Import Skills"
Forms![IMPORT_TRANSFERS].Repaint

    Set qdfCurrent = TVDB.CreateQueryDef("", "DELETE * FROM Process_Skills_Copy;")
    qdfCurrent.Execute
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Process_Skills_Copy", fileName, True, "Skill_Attempts!A1:D100"

Call CLEAN_UP_BLANK_ROWS("SKILLS")

    Set PROCESS_SKILLS = TVDBGM.OpenRecordset("PROCESS_SKILLS")
    PROCESS_SKILLS.index = "PRIMARYKEY"
    
    Set Process_Skills_Copy = TVDB.OpenRecordset("PROCESS_SKILLS_COPY")
    If Process_Skills_Copy.EOF Then
       ' no transfers
       GoTo research
    Else
       Process_Skills_Copy.MoveFirst
    End If
    CURRENT_TRIBE = "EMPTY"
    TRIBE_FIRST = "YES"
    If Not Process_Skills_Copy.NoMatch Then
       Do Until Process_Skills_Copy.EOF
          If CURRENT_TRIBE = Process_Skills_Copy![TRIBE] Then
             'DO NOTHING
          Else
             If IsNull(Process_Skills_Copy![TRIBE]) Then
                Exit Do
             Else
                CURRENT_TRIBE = Process_Skills_Copy![TRIBE]
                TRIBE_FIRST = "YES"
             End If
          End If
          If TRIBE_FIRST = "YES" Then
             PROCESS_SKILLS.Seek "=", Process_Skills_Copy!TRIBE
             If PROCESS_SKILLS.NoMatch Then
                'DO NOTHING
             Else
                Do
                  PROCESS_SKILLS.Delete
                  PROCESS_SKILLS.MoveNext
                  If PROCESS_SKILLS.EOF Then
                     Exit Do
                  End If
                  If Not PROCESS_SKILLS![TRIBE] = Process_Skills_Copy![TRIBE] Then
                      Exit Do
                  End If
                  
                Loop
             End If
          TRIBE_FIRST = "NO"
       End If
       If Process_Skills_Copy![PROCESSED] = "N" Or IsNull(Process_Skills_Copy![PROCESSED]) Then
          If IsNull(Process_Skills_Copy!TRIBE) Then
        ' ignore
          Else
              PROCESS_SKILLS.AddNew
              PROCESS_SKILLS![TRIBE] = Process_Skills_Copy![TRIBE]
              PROCESS_SKILLS![Order] = Process_Skills_Copy![Order]
              PROCESS_SKILLS![TOPIC] = Process_Skills_Copy![TOPIC]
              PROCESS_SKILLS![PROCESSED] = "N"
              PROCESS_SKILLS![Comment] = Process_Skills_Copy![Comment]
              PROCESS_SKILLS.UPDATE
          End If
       End If
       Process_Skills_Copy.Edit
       Process_Skills_Copy![PROCESSED] = "Y"
       Process_Skills_Copy.UPDATE
       Process_Skills_Copy.MoveNext
       If Process_Skills_Copy.EOF Then
          Exit Do
       End If
    Loop
    End If
    
'=====================RESEARCH========================
research:
DebugOP "Importing Research........" & fName
Forms![IMPORT_TRANSFERS]![Status] = "Import Research"
Forms![IMPORT_TRANSFERS].Repaint

    Dim sSQL As String
    
    sSQL = "DELETE * FROM Process_Research_Copy;"
    
    Set qdfCurrent = TVDB.CreateQueryDef("", sSQL)
    qdfCurrent.Execute
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "Process_Research_Copy", fileName, True, "Research_Attempts!A1:C1000"
    
    sSQL = "DELETE * From Process_Research_Copy " & _
            "WHERE Process_Research_Copy.TRIBE Is Null " & _
            "OR Process_Research_Copy.TOPIC Is Null;"
            
    Set qdfCurrent = TVDB.CreateQueryDef("", sSQL)
    qdfCurrent.Execute

Call CLEAN_UP_BLANK_ROWS("RESEARCH")

    Set PROCESS_RESEARCH = TVDBGM.OpenRecordset("PROCESS_RESEARCH")
    PROCESS_RESEARCH.index = "TRIBE"
    
    Set Process_Research_Copy = TVDB.OpenRecordset("PROCESS_RESEARCH_COPY")
    If Process_Research_Copy.EOF Then
       ' no transfers
       GoTo Clean_up
    Else
       Process_Research_Copy.MoveFirst
    End If
    CURRENT_TRIBE = "EMPTY"
    TRIBE_FIRST = "YES"
    If Not Process_Research_Copy.NoMatch Then
       Do Until Process_Research_Copy.EOF
          If CURRENT_TRIBE = Process_Research_Copy![TRIBE] Then
             'DO NOTHING
          Else
             If IsNull(Process_Research_Copy![TRIBE]) Then
                Exit Do
             Else
                CURRENT_TRIBE = Process_Research_Copy![TRIBE]
                TRIBE_FIRST = "YES"
             End If
          End If
          If TRIBE_FIRST = "YES" Then
             PROCESS_RESEARCH.Seek "=", Process_Research_Copy!TRIBE
             If PROCESS_RESEARCH.NoMatch Then
                'DO NOTHING
             Else
                Do
                  PROCESS_RESEARCH.Delete
                  PROCESS_RESEARCH.MoveNext
                  If PROCESS_RESEARCH.EOF Then
                     Exit Do
                  End If
                  If Not PROCESS_RESEARCH![TRIBE] = Process_Research_Copy![TRIBE] Then
                      Exit Do
                  End If
                  
                Loop
             End If
          TRIBE_FIRST = "NO"
       End If
       If Process_Research_Copy![PROCESSED] = "N" Or IsNull(Process_Research_Copy![PROCESSED]) Then
          If IsNull(Process_Research_Copy!TRIBE) Then
        ' ignore
          Else
              PROCESS_RESEARCH.AddNew
              PROCESS_RESEARCH![TRIBE] = Process_Research_Copy![TRIBE]
              PROCESS_RESEARCH![TOPIC] = Process_Research_Copy![TOPIC]
              PROCESS_RESEARCH![PROCESSED] = "N"
              PROCESS_RESEARCH![Comment] = Process_Research_Copy![Comment]
              PROCESS_RESEARCH.UPDATE
          End If
       End If
       Process_Research_Copy.Edit
       Process_Research_Copy![PROCESSED] = "Y"
       Process_Research_Copy.UPDATE
       Process_Research_Copy.MoveNext
       If Process_Research_Copy.EOF Then
          Exit Do
       End If
    Loop
    End If
    
Clean_up:
Forms![IMPORT_TRANSFERS]![Status] = "Clean Up Blank Rows"
Forms![IMPORT_TRANSFERS].Repaint

    Call CLEAN_UP_BLANK_ROWS("ALL")

    TRIBESINFO.Close
    SCOUT_MOVEMENT_COPY.Close
    SCOUT_MOVEMENT_TABLE.Close
    TRIBES_ACTIVITY_IMPLEMENTS.Close
    PROCESSITEMS.Close
    PROCESSACTIVITY.Close

Forms![IMPORT_TRANSFERS]![Status] = "Finished"
Forms![IMPORT_TRANSFERS].Repaint

DebugOP "Importing Tribe Orders Finished........" & fName

ERR_ImportTransfers_CLOSE:
   Exit Sub


ERR_ImportTransfers:
If (Err = 3125) Or Err = 3011 Then
   Err = 0
   Resume Next
   
Else
   Msg = "Error # " & Err & " " & Error$
   DebugOP "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
   DebugOP "!!!!====!!!!==== " & Msg & "====!!!!====!!!!"
   DebugOP "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
   MsgBox (Msg)
   Resume ERR_ImportTransfers_CLOSE
End If


End Sub

 
