Attribute VB_Name = "AB_GASH"
Option Compare Database
Option Explicit

Sub TT123dhjdhj3737dnd()

    Dim sSQL As String
    sSQL = "SELECT  ""("" & [DC_Index] & "")"" & [DC_DesiredCommodity] AS ConcatOP " & vbCrLf & _
            "From Clan_DesiredCommodities " & vbCrLf & _
            "WHERE Clan_DesiredCommodities.DC_CLAN=""0123"";"
            
    Debug.Print ConcatRelatedSQL(sSQL, ", ")

End Sub

Sub akljsdh()
        Dim sOutputTitle As String
    Dim sOutputType As String
    Dim sOutputSubType As String
    Dim sOutputDescription As String
    Dim sCurrentHex As String
    
    sCurrentHex = "AA 0101"
    
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
    End If
        
End Sub

Sub GetSpecialHexText()
    Dim s As String
    Dim sCurrentHex As String
    
    sCurrentHex = "AA 0101"
    
    If Not IsNull(ELookup("SH_Hex", "VALID_SPECIAL_HEXES", "SH_Hex = '" & sCurrentHex & "'")) Then
        s = ELookup("SH_Description", _
                    "VALID_SPECIAL_HEXES", _
                    "SH_Hex = '" & sCurrentHex & "'")
    Else
        s = "No Special Hex"
    End If
    Debug.Print s
End Sub

Sub dhdhd898()
    Debug.Print ChangeTurn2("12/700", 1)
End Sub

Function ChangeTurn2(sTurn As String, iAdj As Integer) As String
    Dim m As Integer
    Dim Y As Integer
    Dim mm As Integer
    
    m = CInt(Left(sTurn, 2)) - 1
    Y = CInt(Mid(sTurn, 4, 3))
    
    mm = 12 * Y + m + iAdj
    
    
    
    m = mm Mod 12 + 1
    
    Y = (mm - m) / 12
    

    
    ChangeTurn2 = Format(m, "00") & "/" & Y
End Function

Sub FormatControls2()
    Dim frm As Form
    Dim ctl As Control
    Dim intX As Integer, intY As Integer
    Dim intH As Integer, intW As Integer
    Dim r As Integer
    Dim C As Integer
    Dim n As Integer
    Dim i As Integer
    Dim el As String 'WHERE string for elookup
    Dim ccrr As String
    
    Set frm = Forms("Map3021")
    If frm.HEX0101.Caption <> "X" Then
    
        For r = 1 To 21
            For C = 1 To 30
    
                ccrr = Format(C, "00") & Format(r, "00")
    
    
                frm.Controls("HEX" & ccrr).BackColor = HEXtoLong("C2C2C2")
                frm.Controls("HEX" & ccrr).Caption = "X"
    
    
            Next C
    
        Next r
    End If
    
    Set frm = Nothing

End Sub

Function FindReplaceNo(sString As String)
    Dim s0 As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    Dim s5 As String
    
    s0 = Mid(sString, 1, 2)
    s1 = Mid(sString, 3, 2)
    s2 = Mid(sString, 5, 2)
    s3 = Mid(sString, 7, 2)
    s4 = Mid(sString, 9, 2)
    s5 = Mid(sString, 11, 2)
    
    If s0 = "No" Then s0 = "NN"
    If s1 = "No" Then s1 = "NN"
    If s2 = "No" Then s2 = "NN"
    If s3 = "No" Then s3 = "NN"
    If s4 = "No" Then s4 = "NN"
    If s5 = "No" Then s5 = "NN"
    
    FindReplaceNo = s0 & s1 & s2 & s3 & s4 & s5
    
 
End Function

Sub weoiruweoiru()
    Debug.Print FindReplaceNo("OnOcNoNoNoNo")
End Sub







