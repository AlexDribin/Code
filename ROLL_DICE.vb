Attribute VB_Name = "ROLL_DICE"
Option Compare Database   'Use database order for string comparisons
Option Explicit

Function ROLL_DICE()
Dim PERCENT As Long
Dim x As String

PERCENT = DICE_ROLL("030", "030")

Msg = PERCENT
Response = MsgBox(Msg, 0)

PERCENT = DICE_ROLL("030", "030")

Msg = PERCENT
Response = MsgBox(Msg, 0)


x = DCount("[VILLAGE]", "TRIBES", "[VILLAGE] = 'TRIBE'")

MsgBox (x)

End Function

Sub DiceTesting123123()
    Dim i As Integer
    For i = 1 To 100
        Debug.Print DROLL(6, 1, 100, 0, 0, 0, 0)

    Next
End Sub


Function DROLL(typ, lvl, sds, res, trb, pre, tmod)
    '=============================
    ' provided by Kingsley March 2023 to replace tvutil.dll
    ' updated April 2023 by Andy to fix error (From i = 1 to i = 11)
    '(see GLOBAL FUNCTIONS)
    '
    'typ = roll_type 1 to 6
    '   1 = Skill 1
    '   2 = Skill 2
    '   3 = Skill 3
    '   4 = First attempt at research DL0
    '   5 = Research attempts DL1 and updwards
    '   6 = Generic dice roll
    '
    'lvl = level
    'sds = dice_sides = 100
    
    'res = reset_roll ..... NOT USED
    'trb = TRIBE .......... NOT USED
    'pre = PRESET ......... NOT USED
    
    'tmod = MODIFY
    '=============================================================================
    Dim research, roll, chk(4), primary(11), dl(21), i As Integer
    Randomize
    research = 5
    For i = 0 To 11
        primary(i) = 110 - (i * 10)
    Next i
    
    'Odds agreed with Peter GM on 2023-04-15
    dl(0) = 0
    dl(1) = 50
    dl(2) = 42
    dl(3) = 35
    dl(4) = 29
    dl(5) = 24
    dl(6) = 20
    dl(7) = 17
    dl(8) = 15
    dl(9) = 12
    dl(10) = 10
    dl(11) = 10
    dl(12) = 10
    dl(13) = 10
    dl(14) = 10
    dl(15) = 10
    dl(16) = 10
    dl(17) = 10
    dl(18) = 10
    dl(19) = 10
    dl(20) = 10
    
    roll = Int(Rnd * sds + 1) - tmod
    DROLL = 0
    If roll < 1 Then
        roll = 1
    ElseIf roll > sds Then
        roll = sds
    End If
    'Debug.Print "roll: " & roll
    
    If typ > 0 And typ < 4 And lvl > 0 And lvl < 11 Then
        'Skill rolls
        chk(1) = primary(lvl)
        chk(2) = chk(1) / 2
        chk(3) = chk(1) / 4
        If roll <= chk(typ) Then
            DROLL = 1
        End If
    ElseIf typ = 4 And roll <= research Then
        'First attempt at research DL0
        DROLL = 1
    ElseIf typ = 5 And lvl > 0 And lvl < 21 And roll <= dl(lvl) Then
        'Research Roll
        DROLL = 1
    ElseIf typ = 6 Then
        'Generating number (1 to sds) - modifier
        DROLL = roll
    End If

End Function

