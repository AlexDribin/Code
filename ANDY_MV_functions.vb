Attribute VB_Name = "ANDY_MV_functions"
Option Compare Database
Option Explicit


Public Function AbsGridCol(sGridRef As String) As Long
'With AA0101 = (1,1) find the absolute column number
    Dim sColumnAlpha As Variant
    Dim iColumnNum As Variant
    
    sColumnAlpha = Mid(sGridRef, 2, 1)
    iColumnNum = CInt(Mid(sGridRef, 4, 2))
    
    AbsGridCol = (Asc(sColumnAlpha) - 65) * 30 + iColumnNum
                    
End Function

Public Function AbsGridRow(sGridRef As String) As Long
'With AA0101 = (1,1) find the absolute row number
    Dim sRowAlpha As Variant
    Dim iRowNum As Variant
    
    

    
    sRowAlpha = Mid(sGridRef, 1, 1)
    iRowNum = CInt(Mid(sGridRef, 6, 2))
    
    AbsGridRow = (Asc(sRowAlpha) - 65) * 21 + iRowNum
                    
End Function

Public Function CRtoGrid(iC As Integer, iR As Integer) As String
    CRtoGrid = Chr(Int((iR - 1) / 21) + 65) & _
                Chr(Int((iC - 1) / 30) + 65) & _
                " " & _
                Format(iC - (Int((iC - 1) / 30)) * 30, "00") & _
                Format(iR - (Int((iR - 1) / 21)) * 21, "00")

End Function

Public Function HexMove(sGridRef As String, _
                        iCm As Integer, _
                        iRm As Integer)
    
    Dim iCol As Integer
    Dim iRow As Integer
    
    
    iCol = AbsGridCol(sGridRef)
    iRow = AbsGridRow(sGridRef)
    
    'The mod functions here ensures that if the column is odd
    'and the columns added are odd then
    'the row is incremented by 1 to allow for staggered
    'nature of hex grids
    iRow = iRow + iRm + ((iCol + 1) Mod 2) * (Abs(iCm) Mod 2)
    
    iCol = iCol + iCm
    
   
    HexMove = CRtoGrid(iCol, iRow)
    
End Function

Public Function GridTest(s As String) As String
    GridTest = CRtoGrid(AbsGridCol(s), AbsGridRow(s))
End Function

'==================Map Flipping================
Sub TT921837()
    Debug.Print AbsGridCol("AP 3001")
End Sub

Public Function MapSwitcheroo(sHex As String) As String
    Dim OP As Variant
    Dim C As Integer
    Dim r As Integer
    
    C = AbsGridCol(sHex)
    r = AbsGridRow(sHex)
    
    If C = 1 Or C > 480 Then
        'nothing, leave as it is
    Else
        C = 241 - (C - 241)
    End If
    
    MapSwitcheroo = CRtoGrid(C, r)
    
End Function

Public Function HexFlip12(sBord As String) As String
    Dim NN As String
    Dim NE As String
    Dim SE As String
    Dim SS As String
    Dim SW As String
    Dim NW As String
    
    
    NN = Left(sBord, 2)
    NE = Mid(sBord, 3, 2)
    SE = Mid(sBord, 5, 2)
    SS = Mid(sBord, 7, 2)
    SW = Mid(sBord, 9, 2)
    NW = Mid(sBord, 11, 2)
    
    HexFlip12 = NN & NW & SW & SS & SE & NE
    
End Function

Public Function HexFlip6(sBord As String) As String
    Dim NN As String
    Dim NE As String
    Dim SE As String
    Dim SS As String
    Dim SW As String
    Dim NW As String
    
    
    NN = Left(sBord, 1)
    NE = Mid(sBord, 2, 1)
    SE = Mid(sBord, 3, 1)
    SS = Mid(sBord, 4, 1)
    SW = Mid(sBord, 5, 1)
    NW = Mid(sBord, 6, 1)
    
    HexFlip6 = NN & NW & SW & SS & SE & NE
    
End Function






