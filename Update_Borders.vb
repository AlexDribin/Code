Attribute VB_Name = "Update_Borders"
Option Compare Database
Option Explicit

Function BorderN(sBorders As String, n As Integer) As String
    Dim aB() As String
    aB = BorderStringToArray(sBorders)
    
    BorderN = aB(n)
End Function

Function BorderStringToArray(sBorders As String) As Variant
    Dim aB(6) As String
    aB(0) = Left(sBorders, 2)
    aB(1) = Mid(sBorders, 3, 2)
    aB(2) = Mid(sBorders, 5, 2)
    aB(3) = Mid(sBorders, 7, 2)
    aB(4) = Mid(sBorders, 9, 2)
    aB(5) = Mid(sBorders, 11, 2)
    
    BorderStringToArray = aB
End Function

Public Sub BorderUpdate()
    Dim db As DAO.Database
    
    Dim sSQL As String
    Dim vA(5, 1) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim t As String
    Dim AffectedRows As Long
    
    Set db = CurrentDb
    
    
    
    vA(0, 0) = "N"
    vA(0, 1) = "0, -1"
    vA(1, 0) = "NE"
    vA(1, 1) = "1, -1"
    vA(2, 0) = "SE"
    vA(2, 1) = "1, 0"
    vA(3, 0) = "S"
    vA(3, 1) = "0, 1"
    vA(4, 0) = "SW"
    vA(4, 1) = "-1, 0"
    vA(5, 0) = "NW"
    vA(5, 1) = "-1, -1"
    
    For i = 0 To 5
        ifTableExistsDrop ("AABorderUpdate" & i & vA(i, 0))
    Next i
    
    
    For i = 0 To 5
        sSQL = "SELECT HEX_MAP.MAP, Mid(HEX_MAP.Borders, " & _
                2 * i + 1 & _
                ", 2) AS Bord, HexMove([MAP]," & _
                vA(i, 1) & _
                ") AS AdjHex " & _
                "INTO AABorderUpdate" & _
                i & vA(i, 0) & _
                " FROM HEX_MAP " & _
                "WHERE Borders <> 'NNNNNNNNNNNN';"
        Debug.Print sSQL
        db.Execute sSQL, dbFailOnError
    Next i
    
    AffectedRows = 0
    For i = 0 To 5
        j = (i + 3) Mod 6 'apply change to opposite side

        t = "AABorderUpdate" & i & vA(i, 0) 'joined table name
    
        sSQL = "UPDATE HEX_MAP INNER JOIN " & _
                t & " " & _
                "ON HEX_MAP.MAP = " & t & ".AdjHex " & _
                "SET HEX_MAP.Borders = " & _
                "Mid([Borders],1," & _
                2 * j & _
                ") & [Bord] & Mid([Borders]," & _
                2 * j + 3 & _
                ") " & _
                "WHERE (((" & t & ".Bord) Not In ('NN','No'))) " & _
                "AND (Mid(HEX_MAP.Borders, " & _
                2 * j + 1 & _
                ", 2) = 'NN' OR Mid(HEX_MAP.Borders, " & _
                2 * j + 1 & _
                ", 2) = 'No');"

        db.Execute sSQL, dbFailOnError

        AffectedRows = AffectedRows + db.RecordsAffected
        Debug.Print db.RecordsAffected
        Debug.Print sSQL
    Next i

    For i = 0 To 5
        ifTableExistsDrop ("AABorderUpdate" & i & vA(i, 0))
    Next i
    MsgBox "Updated " & AffectedRows & " borders."
End Sub

Public Sub ifTableExistsDrop(tblName As String)
    Dim sSQL As String
    

    If DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "'") = 1 Then
        sSQL = "DROP TABLE " & tblName & ";"
        CurrentDb.Execute sSQL, dbFailOnError

    End If

End Sub

Sub GetBorders()
    Dim rs As DAO.Recordset
    Dim sSQL As String
    Dim sMapSheet As String
    Dim n As Integer
    Dim m As Integer
    Dim frm As Form
    
    Set frm = Forms("Map3021")
    sMapSheet = frm.ComboMapSheet
    
    sSQL = "SELECT * From HEX_MAP WHERE Left(MAP, 2) = '" & _
            sMapSheet & _
            "' " & _
            "AND Borders <> 'NNNNNNNNNNNN' " & _
            "AND Borders <> 'NoNoNoNoNoNo'" & _
            ";"
    Set rs = CurrentDb.OpenRecordset(sSQL)
    n = 1
    'Check to see if the recordset actually contains rows
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
            'Perform an edit


            m = PlaceBorderPosition(rs!MAP, rs!Borders, n)
            n = m
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        'MsgBox "There are no records in the recordset."
    End If
    
    'MsgBox "Finished looping through records."
    
    rs.Close 'Close the recordset
    Set rs = Nothing 'Clean up
End Sub

Function PlaceBorderPosition(sMap As String, sBorders As String, n)
    Dim aB As Variant
    Dim i As Integer
    Dim C As Integer
    Dim r As Integer
    Dim frm As Form
    Dim ctl As Control
    Dim intX As Integer, intY As Integer
    Dim intH As Integer, intW As Integer
    Dim lH As Integer, lW As Integer
    Dim adjX As Integer, adjY As Integer
    Dim x As Integer, Y As Integer
    Dim h As Integer, w As Integer

    Set frm = Forms("Map3021")
    intH = 500
    intW = 500

    intX = 0
    intY = 0
    
    aB = BorderStringToArray(sBorders)
    'Debug.Print sMap
    
    C = CInt(Mid(sMap, 4, 2))
    r = CInt(Mid(sMap, 6, 2))
    For i = 0 To 5
        If aB(i) <> "NN" And aB(i) <> "No" Then
            
            x = intX + C * intW
            Y = intY + r * intH + ((C + 1) Mod 2) * intH / 2
            w = 0
            h = 500
            
            Select Case i
            Case 0
                w = intW
                h = 0
            Case 1
                x = x + intW
                w = 0
                h = intH / 2
            Case 2
                x = x + intW
                Y = Y + intH / 2
                h = intH / 2
            Case 3
                Y = Y + intH
                w = intW
                h = 0
            Case 4
                Y = Y + intH / 2
                w = 0
                h = intH / 2
            Case 5
                w = 0
                h = intH / 2
            End Select
            
            
            'Debug.Print "Column " & c & " Row " & r & " - Border " & i & " is " & aB(i)
            Set ctl = frm.Controls("Line" & n)
            ctl.Move x, Y, w, h
            ctl.BorderWidth = 4
            ctl.BorderColor = HEXtoLong("0000EC")
            n = n + 1
        End If
    Next i
    PlaceBorderPosition = n
    
    Set frm = Nothing
    Set ctl = Nothing
End Function

Sub ResetBorders()
    Dim frm As Form
    Dim ctl As Control
    Dim i As Integer
    Dim x As Integer
    Dim Y As Integer
    
    x = 16000
    Y = 10000
    
    Set frm = Forms("Map3021")
    i = 0
    For Each ctl In frm
        If ctl.ControlType = acLine Then
            
            'Debug.Print ctl.Name
            ctl.Move x, Y, 0, 500
            ctl.BorderWidth = 4
            ctl.BorderColor = HEXtoLong("000000")
            
            i = i + 1
        End If
    Next ctl


End Sub

Sub UpdateHexesSingle(sXXccrr As String)
    Dim vA() As String
    Dim r As Integer
    Dim C As Integer
    Dim sA As String
    Dim sT As String
    Dim sC As String
    Dim sM As String
'    Dim XXccrr As String
    Dim ccrr As String
    
        
            'XXccrr = Me.MAP
            ccrr = Mid(sXXccrr, 4)
            
            
'            sA = ELookup("PipeData", "Map2130q", "MAP = '" & _
'                        XXccrr & "'")
                        
            sA = Nz(ELookup("PipeData", "Map2130q", "MAP = '" & _
                        sXXccrr & "'"), _
                        "X|C2C2C2|NNNNNNNNNNNN|")
                        
            vA = Split(sA, "|")
            
            sT = vA(0)
            sC = vA(1)
            sM = vA(3)
            If Len(sM) > 0 Then
                sT = sT & vbCrLf & sM
            End If

            Forms!Map3021.Controls("HEX" & ccrr).Caption = sT
            Forms!Map3021.Controls("HEX" & ccrr).BackColor = HEXtoLong(sC)
            Forms!Map3021.Controls("HEX" & ccrr).ForeColor = HEXtoFontColor(sC)
            
            'HEXtoLong
            
End Sub

Public Sub UpdateSingleHexSurroundingBorders(sHex As String)

    Dim aB As Variant
    Dim sBorder As String
    
    'declare hex strings
    Dim h0 As String
    Dim h1 As String
    Dim h2 As String
    Dim h3 As String
    Dim h4 As String
    Dim h5 As String
    
    'declare border strings
    Dim b0 As String
    Dim b1 As String
    Dim b2 As String
    Dim b3 As String
    Dim b4 As String
    Dim b5 As String
    
    sBorder = ELookup("Borders", "HEX_MAP", _
                "MAP = '" & sHex & "'")
                
    aB = BorderStringToArray(sBorder)
    
    
    'find adjacent hex strings
    h0 = HexMove(sHex, 0, -1)
    h1 = HexMove(sHex, 1, -1)
    h2 = HexMove(sHex, 1, 0)
    h3 = HexMove(sHex, 0, 1)
    h4 = HexMove(sHex, -1, 0)
    h5 = HexMove(sHex, -1, -1)
    
    Debug.Print "h0: " & h0
    Debug.Print "h1: " & h1
    Debug.Print "h2: " & h2
    Debug.Print "h3: " & h3
    Debug.Print "h4: " & h4
    Debug.Print "h5: " & h5
    
    
    'find border strings of adjacent hexes
    If Not IsNull(h0) Then
        b0 = Nz(ELookup("Borders", "HEX_MAP", _
                "MAP = '" & h0 & "'"), "")
        If Len(b0) > 0 Then
            b0 = ReplaceOneBorder(b0, 3, CStr(aB(0)))
            Call UpdateBorderString(h0, b0)
        End If
    End If
    
    If Not IsNull(h1) Then
        b1 = Nz(ELookup("Borders", "HEX_MAP", _
                "MAP = '" & h1 & "'"), "")
        If Len(b1) > 0 Then
            b1 = ReplaceOneBorder(b1, 4, CStr(aB(1)))
            Call UpdateBorderString(h1, b1)
        End If
    End If
    
    If Not IsNull(h2) Then
        b2 = Nz(ELookup("Borders", "HEX_MAP", _
                "MAP = '" & h2 & "'"), "")
        If Len(b2) > 0 Then
            b2 = ReplaceOneBorder(b2, 5, CStr(aB(2)))
            Call UpdateBorderString(h2, b2)
        End If
    End If
    
    If Not IsNull(h3) Then
        b3 = Nz(ELookup("Borders", "HEX_MAP", _
                "MAP = '" & h3 & "'"), "")
        If Len(b3) > 0 Then
            b3 = ReplaceOneBorder(b3, 0, CStr(aB(3)))
            Call UpdateBorderString(h3, b3)
        End If
    End If
    
    If Not IsNull(h4) Then
        b4 = Nz(ELookup("Borders", "HEX_MAP", _
                "MAP = '" & h4 & "'"), "")
        If Len(b4) > 0 Then
            b4 = ReplaceOneBorder(b4, 1, CStr(aB(4)))
            Call UpdateBorderString(h4, b4)
        End If
    End If
    
    If Not IsNull(h5) Then
        b5 = Nz(ELookup("Borders", "HEX_MAP", _
                "MAP = '" & h5 & "'"), "")
        If Len(b5) > 0 Then
            b5 = ReplaceOneBorder(b5, 2, CStr(aB(5)))
            Call UpdateBorderString(h5, b5)
        End If
    End If
    
    Debug.Print "b0: " & b0
    Debug.Print "b1: " & b1
    Debug.Print "b2: " & b2
    Debug.Print "b3: " & b3
    Debug.Print "b4: " & b4
    Debug.Print "b5: " & b5
    
End Sub

Public Sub UpdateBorderString(sHex As String, sBorderString As String)
    Dim sSQL As String
    
    sSQL = "UPDATE HEX_MAP SET HEX_MAP.Borders = '" & _
            sBorderString & "' " & _
            "WHERE HEX_MAP.MAP='" & _
            sHex & "';"
            
    CurrentDb.Execute sSQL, dbFailOnError
End Sub

Function SplitBorder(sBorders As String, i As Integer) As Variant
    SplitBorder = Mid(sBorders, 1 + 2 * i, 2)
End Function

Function ReplaceOneBorder(sBorder As String, _
                            i As Integer, _
                            sReplace As String)
    Dim aB As Variant
    
    aB = BorderStringToArray(sBorder)
    aB(i) = sReplace
    
    ReplaceOneBorder = Join(aB, "")
    Debug.Print ReplaceOneBorder
                            
                            
End Function
Private Sub oiwerroiuweoi()
    Call UpdateSingleHexSurroundingBorders("GG 0302")
End Sub

Private Sub asjklfdkljf()
    Call UpdateBorderString("AA 0101", "NNNNNNNNNNNN")
End Sub










