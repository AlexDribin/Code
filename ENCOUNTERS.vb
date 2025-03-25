Attribute VB_Name = "ENCOUNTERS"
Option Compare Database   'Use database order for string comparisons
Option Explicit
Global LINENUMBERS As Recordset
Global HEXSTOP As String
Global TRADING_POST_FOUND As String
Global LINENUMBER As Long


'*===============================================================================*'
'*****                      MAINTENANCE LOG                                  *****'
'*-------------------------------------------------------------------------------*'
'**   DATE    *  DESCRIPTION                                                    **'
'*-------------------------------------------------------------------------------*'
'** 17/01/96  *  Insert Maintenance Log                                         **'
'*===============================================================================*'
 

Function ENCOUNTERS_PROCESS()
On Error GoTo ERR_ENCOUNTERS
TRIBE_STATUS = "Encounters Process"

Set MYFORM = Forms![ENCOUNTERS]
  
Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)

Set TRIBEINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBEINFO.index = "PRIMARYKEY"
TRIBEINFO.Seek "=", MYFORM![CLAN NUMBER], MYFORM![TRIBE NUMBER]

LINENUMBER = 1
CLANNUMBER = Forms!ENCOUNTERS![CLAN NUMBER]
TRIBENUMBER = Forms!ENCOUNTERS![TRIBE NUMBER]


Set ActivitiesTable = TVDBGM.OpenRecordset("Turns_Activities")
ActivitiesTable.index = "PRIMARYKEY"
ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", CLANNUMBER, TRIBENUMBER, "COMMENTS", LINENUMBER

If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![CLAN] = CLANNUMBER
   ActivitiesTable![TRIBE] = TRIBENUMBER
   ActivitiesTable![Section] = "COMMENTS"
   ActivitiesTable![LINE NUMBER] = LINENUMBER
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![Comments1]
   ActivitiesTable.UPDATE
Else
   ActivitiesTable.Edit
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![Comments1]
   ActivitiesTable.UPDATE
End If

LINENUMBER = 1

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", CLANNUMBER, TRIBENUMBER, "ENCOUNTERS", LINENUMBER

If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![CLAN] = Forms!ENCOUNTERS![CLAN NUMBER]
   ActivitiesTable![TRIBE] = Forms!ENCOUNTERS![TRIBE NUMBER]
   ActivitiesTable![Section] = "ENCOUNTERS"
   ActivitiesTable![LINE NUMBER] = LINENUMBER
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![ENCOUNTER1]
   ActivitiesTable.UPDATE
Else
   ActivitiesTable.Edit
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![ENCOUNTER1]
   ActivitiesTable.UPDATE
End If

LINENUMBER = LINENUMBER + 1

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", CLANNUMBER, TRIBENUMBER, "ENCOUNTERS", LINENUMBER

If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![CLAN] = Forms!ENCOUNTERS![CLAN NUMBER]
   ActivitiesTable![TRIBE] = Forms!ENCOUNTERS![TRIBE NUMBER]
   ActivitiesTable![Section] = "ENCOUNTERS"
   ActivitiesTable![LINE NUMBER] = LINENUMBER
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![ENCOUNTER2]
   ActivitiesTable.UPDATE
Else
   ActivitiesTable.Edit
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![ENCOUNTER2]
   ActivitiesTable.UPDATE
End If

LINENUMBER = LINENUMBER + 1

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", CLANNUMBER, TRIBENUMBER, "ENCOUNTERS", LINENUMBER

If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![CLAN] = Forms!ENCOUNTERS![CLAN NUMBER]
   ActivitiesTable![TRIBE] = Forms!ENCOUNTERS![TRIBE NUMBER]
   ActivitiesTable![Section] = "ENCOUNTERS"
   ActivitiesTable![LINE NUMBER] = LINENUMBER
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![ENCOUNTER3]
   ActivitiesTable.UPDATE
Else
   ActivitiesTable.Edit
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![ENCOUNTER3]
   ActivitiesTable.UPDATE
End If

LINENUMBER = 1

ActivitiesTable.MoveFirst
ActivitiesTable.Seek "=", CLANNUMBER, TRIBENUMBER, "RESPONSE", LINENUMBER

If ActivitiesTable.NoMatch Then
   ActivitiesTable.AddNew
   ActivitiesTable![CLAN] = Forms!ENCOUNTERS![CLAN NUMBER]
   ActivitiesTable![TRIBE] = Forms!ENCOUNTERS![TRIBE NUMBER]
   ActivitiesTable![Section] = "RESPONSE"
   ActivitiesTable![LINE NUMBER] = LINENUMBER
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![Response1]
   ActivitiesTable.UPDATE
Else
   ActivitiesTable.Edit
   ActivitiesTable![line detail] = Forms![ENCOUNTERS]![Response1]
   ActivitiesTable.UPDATE
End If

ActivitiesTable.Close

  Forms![ENCOUNTERS]![Comments1] = "EMPTY"
  Forms![ENCOUNTERS]![ENCOUNTER1] = "EMPTY"
  Forms![ENCOUNTERS]![ENCOUNTER2] = "EMPTY"
  Forms![ENCOUNTERS]![ENCOUNTER3] = "EMPTY"
  Forms![ENCOUNTERS]![Response1] = "EMPTY"

EXIT_FORMS ("ENCOUNTERS")
OPEN_FORMS ("ENCOUNTERS")

ERR_ENCOUNTERS_CLOSE:
   Exit Function


ERR_ENCOUNTERS:
If (Err = 3021) Then
   Resume Next
   
Else
  Msg = "Error # " & Err & " " & Error$
  MsgBox (Msg)
  Resume ERR_ENCOUNTERS_CLOSE
  
End If

End Function

