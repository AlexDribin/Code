Attribute VB_Name = "GM_ACTIONS"
Option Compare Database
Option Explicit

Function GetRandRecord(sKeyField As String, sTable As String) As Variant
       
    Dim rs As DAO.Recordset
    Dim n As Long
    Dim i As Long
    Dim OP As Variant
    'Debug.Print strSQL
    
    Randomize
    
            
    Set rs = CurrentDb.OpenRecordset(sTable)
    
     If Not (rs.EOF And rs.BOF) Then
        rs.MoveLast
        i = rs.RecordCount
        n = Int(i * Rnd + 1)
     
        rs.MoveFirst 'Unnecessary in this case, but still a good habit

        rs.Move (n - 1)

        OP = rs(sKeyField)
    Else
        MsgBox "There are no records in the recordset."
    End If
    
    GetRandRecord = OP
    
    rs.Close 'Close the recordset
    Set rs = Nothing 'Clean up
End Function

Sub testrandom()
    Debug.Print GetRandRecord("VDC_Name", "Valid_desiredCommodities")
End Sub

Function GMA_New_Clan(strClanNo As String, strStartHex As String, sPlayerName As String)
    Dim HEX_MAP As String
    
    Set TVWKSPACE = DBEngine.Workspaces(0)
    
    Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
    Set GMTABLE = TVDB.OpenRecordset("GM")
    GMTABLE.index = "PRIMARYKEY"
    GMTABLE.MoveFirst
    
    FILEGM = CurDir$ & "\" & GMTABLE![FILE]
    
    Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
       
    TCLANNUMBER = strClanNo
    HEX_MAP = strStartHex
    
    
    
    Set TRIBESINFO = TVDBGM.OpenRecordset("Tribes_General_Info")
    TRIBESINFO.index = "PRIMARYKEY"
    TRIBESINFO.AddNew
    TRIBESINFO![CLAN] = TCLANNUMBER
    TRIBESINFO![TRIBE] = TCLANNUMBER
    TRIBESINFO![Village] = "Tribe"
    TRIBESINFO![CURRENT TERRAIN] = "PRAIRIE"
    TRIBESINFO![OWNER] = sPlayerName
    TRIBESINFO![GOODS_CLAN] = TCLANNUMBER
    TRIBESINFO![GOODS TRIBE] = TCLANNUMBER
    TRIBESINFO![COST CLAN] = TCLANNUMBER
    TRIBESINFO![WARRIORS] = 5890
    TRIBESINFO![ACTIVES] = 5890
    TRIBESINFO![INACTIVES] = 5890
    TRIBESINFO![MORALE] = 1
    TRIBESINFO![CREDIT] = 0                     '!!!! ========== needs changing
    TRIBESINFO![Current Hex] = HEX_MAP
    TRIBESINFO.UPDATE
    TRIBESINFO.Close
    
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "CATTLE", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "GOAT", "ADD", 3700)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "HORSE", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "COAL", "ADD", 3000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "IRON", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SILVER", "ADD", 10000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "CLUB", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SHIELD", "ADD", 30)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "JERKIN", "ADD", 200)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SWORD", "ADD", 30)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "PROVS", "ADD", 50000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SLING", "ADD", 300)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "WAGON", "ADD", 300)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BARK", "ADD", 1000)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BONES", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "GUT", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "LEATHER", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "WAX", "ADD", 20)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "SKIN", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BRONZE", "ADD", 400)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "LOG", "ADD", 100)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "BRASS", "ADD", 500)
    Call UPDATE_TRIBES_GOODS_TABLES(TCLANNUMBER, TCLANNUMBER, "TRAP", "ADD", 500)


    
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ADMINISTRATION", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "BONEWORK", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "BONING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "CURING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "DIPLOMACY", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ECONOMICS", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "ENGINEERING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "FORESTRY", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "GARRISON", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "GUTTING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "HERDING", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "HUNTING", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "LEADERSHIP", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "LEATHERWORK", 3)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "QUARRYING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "SCOUTING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "SKINNING", 2)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "TANNING", 1)
    Call UPDATE_TRIBES_SKILLS_TABLE(TCLANNUMBER, "WOODWORK", 3)
    
    'OutLine = "You may distribute a further 30 skill points (new, or add to existing skills) but no skill level may exceed 7."
    'OutLine = OutLine & "{enter}"
    '
    'Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 1, OutLine, "No")
    '
    '
    'OutLine = "You may choose to take either 900 iron or 1200 bronze."
    'OutLine = OutLine & "{enter}{enter}"
    '
    'Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 2, OutLine, "No")
    '
    'OutLine = "Special skills / items:"
    'OutLine = OutLine & "{enter}"
    '
    'Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 3, OutLine, "No")
    '
    'OutLine = "Horse Bow:  can make with Wpn6, horse archers can participate in missile phase and charge phase."
    'OutLine = OutLine & "{enter}"
    '
    'Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 4, OutLine, "No")
    'OutLine = "Elephants:      can handle, may use elephants for transport."
    'OutLine = OutLine & "{enter}"
    '
    'Call WRITE_TURN_ACTIVITY(TCLANNUMBER, TCLANNUMBER, "NEWCLAN", 5, OutLine, "No")
    
    'Farming insterted by AB from unit creation
    Set Globaltable = TVDBGM.OpenRecordset("GLOBAL")
    Globaltable.MoveFirst
    Current_Turn = Globaltable![CURRENT TURN]
    Globaltable.Close
    
    Set FarmingTable = TVDBGM.OpenRecordset("TRIBE_FARMING")
    FarmingTable.index = "PRIMARYKEY"
    FarmingTable.Seek "=", strClanNo, strClanNo, Current_Turn
    
    If FarmingTable.NoMatch Then
       FarmingTable.AddNew
       FarmingTable![CLAN] = strClanNo
       FarmingTable![TRIBE] = strClanNo
       FarmingTable![TURN] = Current_Turn
       FarmingTable![ITEM] = "START"
       FarmingTable.UPDATE
       FarmingTable.Close
    End If
    
    Call Tribe_Checking("Update_All", "", "", "")
    
End Function

Function GMA_NEW_GROUP()

Set TVWKSPACE = DBEngine.Workspaces(0)

Set TVDB = TVWKSPACE.OpenDatabase("tvdatapr.accdb", False, False)
Set GMTABLE = TVDB.OpenRecordset("GM")
GMTABLE.index = "PRIMARYKEY"
GMTABLE.MoveFirst

FILEGM = CurDir$ & "\" & GMTABLE![FILE]

Set TVDBGM = TVWKSPACE.OpenDatabase(FILEGM, False, False)
   
Set TRIBESINFO = TVDBGM.OpenRecordset("TRIBES_GENERAL_INFO")
TRIBESINFO.index = "PRIMARYKEY"
TRIBESINFO.Seek "=", Forms![NEW GROUP]![PARENT CLAN], Forms![NEW GROUP]![PARENT TRIBE]
CURRENT_HEX = TRIBESINFO![Current Hex]
CURRENT_COST_CLAN = TRIBESINFO![COST CLAN]
CURRENT_TERRAIN = TRIBESINFO![CURRENT TERRAIN]
If IsNull(TRIBESINFO![RELIGION]) Then
   RELIGION = ""
Else
   RELIGION = TRIBESINFO![RELIGION]
End If
If IsNull(TRIBESINFO![CULT]) Then
   CULT = ""
Else
   CULT = TRIBESINFO![CULT]
End If

TRIBESINFO.AddNew
TRIBESINFO![CLAN] = Forms![NEW GROUP]![New Clan]
TRIBESINFO![TRIBE] = Forms![NEW GROUP]![NEW TRIBE]
TRIBESINFO![TRIBE NAME] = Null
TRIBESINFO![Current Hex] = CURRENT_HEX
TRIBESINFO![CURRENT TERRAIN] = CURRENT_TERRAIN
TRIBESINFO![Village] = Forms![NEW GROUP]![Village]
If Not IsNull(RELIGION) And Not (RELIGION = "") Then
   TRIBESINFO![RELIGION] = RELIGION
End If
If Not IsNull(CULT) And Not (CULT = "") Then
   TRIBESINFO![CULT] = CULT
End If
TRIBESINFO![CREDIT] = 0
TRIBESINFO![AMT RECEIVED] = 0
TRIBESINFO![Cost] = 0
TRIBESINFO![OWNER] = Null
TRIBESINFO![EMAIL] = "N"
TRIBESINFO![GOODS_CLAN] = Null
TRIBESINFO![GOODS TRIBE] = Null
TRIBESINFO![POP TRIBE] = Null
TRIBESINFO![COST CLAN] = CURRENT_COST_CLAN
TRIBESINFO![MORALE] = 1
TRIBESINFO.UPDATE
TRIBESINFO.Close

Set TRIBES_CHECKING = TVDBGM.OpenRecordset("TRIBE_CHECKING")
TRIBES_CHECKING.index = "PRIMARYKEY"
TRIBES_CHECKING.AddNew
TRIBES_CHECKING![CLAN] = Forms![NEW GROUP]![New Clan]
TRIBES_CHECKING![TRIBE] = Forms![NEW GROUP]![NEW TRIBE]
TRIBES_CHECKING![Current Hex] = CURRENT_HEX
TRIBES_CHECKING.UPDATE
TRIBES_CHECKING.Close

Set Globaltable = TVDBGM.OpenRecordset("GLOBAL")
Globaltable.MoveFirst
Current_Turn = Globaltable![CURRENT TURN]
Globaltable.Close

Set FarmingTable = TVDBGM.OpenRecordset("TRIBE_FARMING")
FarmingTable.index = "PRIMARYKEY"
FarmingTable.Seek "=", Forms![NEW GROUP]![New Clan], Forms![NEW GROUP]![NEW TRIBE], Current_Turn

If FarmingTable.NoMatch Then
   FarmingTable.AddNew
   FarmingTable![CLAN] = Forms![NEW GROUP]![New Clan]
   FarmingTable![TRIBE] = Forms![NEW GROUP]![NEW TRIBE]
   FarmingTable![TURN] = Current_Turn
   FarmingTable![ITEM] = "START"
   FarmingTable.UPDATE
   FarmingTable.Close
End If

Call EXIT_FORMS("NEW GROUP")
Call OPEN_FORMS("NEW GROUP")

End Function

Public Function GMA_AddDeclaration(sSourceClan, sDeclaration, sDestClan)
    Dim sSQLD As String
    Dim sSQL As String
    Dim OP As String
    
    sSQLD = "DELETE * " & _
            "FROM CLAN_DECLARATIONS " & _
            "WHERE CLAN_DECLARATIONS.CLAN_SOURCE='" & _
            sSourceClan & _
            "' " & _
            "AND CLAN_DECLARATIONS.CLAN_DESTINATION='" & _
            sDestClan & _
            "';"
    sSQL = "INSERT INTO CLAN_DECLARATIONS " & _
            "( CLAN_SOURCE, DECLARATION, CLAN_DESTINATION ) " & _
            "SELECT '" & _
            sSourceClan & "', '" & _
            sDeclaration & _
            "', '" & _
            sDestClan & _
            "';"
            
    CurrentDb.Execute sSQLD, dbFailOnError
    
    If sSourceClan = sDestClan Then
        OP = "Cannot make declaration on same Clan."
    ElseIf sDeclaration = "Cease War" Or sDeclaration = "End Truce" Then
        OP = "Declaration of " & sDeclaration & " from " & _
            sSourceClan & " with " & sDestClan & " registered."
    ElseIf sDeclaration = "War" Or sDeclaration = "Truce" Then
        CurrentDb.Execute sSQL, dbFailOnError
        OP = "Declaration of " & sDeclaration & " from " & _
            sSourceClan & " with " & sDestClan & " registered."
    Else
        OP = "Error - declaration ran into problems"
    End If
            
    GMA_AddDeclaration = OP
        
    
End Function

Public Sub ResolveDeclarations()
    Dim rs As DAO.Recordset
    Dim sSQL As String
    Dim MirrorRecordCount As Variant
    
    sSQL = "SELECT * FROM CLAN_DECLARATIONS WHERE DECLARATION = 'Truce';"
    Set rs = CurrentDb.OpenRecordset(sSQL)
    
    'Check to see if the recordset actually contains rows
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
            'check if there's a parallel TRUCE record
            MirrorRecordCount = DCount("ID", "CLAN_DECLARATIONS", _
                "CLAN_SOURCE = '" & rs!CLAN_DESTINATION & _
                "' AND CLAN_DESTINATION = '" & _
                rs!CLAN_SOURCE & "'" & _
                "AND DECLARATION = 'Truce'")
            Debug.Print MirrorRecordCount

            
            
            If MirrorRecordCount = 0 Then
                'if no parallel TRUCE record then delete
                DebugOP rs!CLAN_SOURCE & _
                        " truce with " & _
                        rs!CLAN_DESTINATION & _
                        " is unilateral and will be deleted."
                rs.Delete
                
            End If
    
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        'MsgBox "There are no records in the recordset."
    End If
    
    rs.Close 'Close the recordset
    Set rs = Nothing 'Clean up
    DebugOP "All Declarations Resolved."
    Call UpdateTribeTableWithTruceWarStrings
    
End Sub

Public Sub UpdateTribeTableWithTruceWarStrings()
    Dim rs As DAO.Recordset
    Dim sSQL As String
    Dim sTruce As Variant
    Dim sWar As Variant
    sSQL = "SELECT * FROM TRIBES_general_info " & _
            "WHERE VILLAGE = 'Tribe' " & _
            "AND Left(TRIBE, 1) = '0';"
    Set rs = CurrentDb.OpenRecordset(sSQL)
    
    'Check to see if the recordset actually contains rows
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
            'Perform an edit
            sTruce = ConcatRelated("CLAN_DESTINATION", "CLAN_DECLARATIONS", "CLAN_SOURCE = '" & rs!CLAN & "' AND DECLARATION = 'Truce'")
            sTruce = Left(sTruce, 200)
            sWar = ConcatRelated("CLAN_DESTINATION", "CLAN_DECLARATIONS", "CLAN_SOURCE = '" & rs!CLAN & "' AND DECLARATION = 'War'")
            sWar = Left(sWar, 50)
            rs.Edit
            rs!TRUCES = sTruce
            rs!War = sWar
            rs.UPDATE
    
            'Save contact name into a variable

            Debug.Print rs!TRIBE
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        'MsgBox "There are no records in the recordset."
    End If
    
    'MsgBox "Finished looping through records."
    
    rs.Close 'Close the recordset
    Set rs = Nothing 'Clean up
    DebugOP "Truces and War strings entered into TRIBES_general_info Table."
End Sub

