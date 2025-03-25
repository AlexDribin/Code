Attribute VB_Name = "TurnFunctions"
Option Compare Database
Option Explicit


Public Function ChangeTurn(sTurn As String, iAdj As Integer) As String
    'mm/YYY
    Dim m As Integer
    Dim Y As Integer
    Dim mm As Integer
    
    m = CInt(Left(sTurn, 2)) - 1
    Y = CInt(Mid(sTurn, 4, 3))

    mm = 12 * Y + m + iAdj
    
    m = mm Mod 12 + 1
    
    Y = (mm - m) / 12
        
    ChangeTurn = Format(m, "00") & "/" & Y
End Function

Public Function ChangeTurnYYYMM(sTurn As String, iAdj As Integer) As String
    'YYY-MM
    Dim m As Integer
    Dim Y As Integer
    Dim mm As Integer
    
    m = CInt(Mid(sTurn, 5, 2)) - 1
    Y = CInt(Mid(sTurn, 1, 3))

    
    mm = 12 * Y + m + iAdj
    
    m = mm Mod 12 + 1
    
    Y = (mm - m) / 12
        
    ChangeTurnYYYMM = Y & "-" & Format(m, "00")
End Function

Public Function GetCurrentTurn() As String
    'Gets the current turn and reformats to YYY-MM
    Dim sTurn As String
    sTurn = DLookup("[CURRENT TURN]", _
                    "[GLOBAL]", _
                    "[GLOBAL] = 'GLOBAL'")
                    
    sTurn = Mid(sTurn, 4) & "-" & Left(sTurn, 2)
                    
    GetCurrentTurn = sTurn
End Function

Public Function GetCurrentTurnNo() As Integer
    'Gets the current turn and reformats to increments from 900-01
    Dim sTurn As String
    Dim vY As Variant
    Dim vM As Variant
    
    sTurn = DLookup("[CURRENT TURN]", _
                    "[GLOBAL]", _
                    "[GLOBAL] = 'GLOBAL'")
                    
    vY = Mid(sTurn, 4, 3)
    vM = Left(sTurn, 2)
    
    vY = CInt(vY)
    vM = CInt(vM)
    
                    
    GetCurrentTurnNo = (vY * 12) + vM - 10800
End Function



Sub asdljkdfslkj()
    
    Debug.Print ChangeTurnYYYMM(GetCurrentTurn(), 1)

    
End Sub

