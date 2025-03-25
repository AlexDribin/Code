Attribute VB_Name = "MapBMPCreate"
Option Compare Database
Option Explicit

Type typHEADER
    strType As String * 2  ' Signature of file = "BM"
    lngSize As Long        ' File size
    intRes1 As Integer     ' reserved = 0
    intRes2 As Integer     ' reserved = 0
    lngOffset As Long      ' offset to the bitmap data (bits)
End Type

Type typINFOHEADER
    lngSize As Long        ' Size
    lngWidth As Long       ' Height
    lngHeight As Long      ' Length
    intPlanes As Integer   ' Number of image planes in file
    intBits As Integer     ' Number of bits per pixel
    lngCompression As Long ' Compression type (set to zero)
    lngImageSize As Long   ' Image size (bytes, set to zero)
    lngxResolution As Long ' Device resolution (set to zero)
    lngyResolution As Long ' Device resolution (set to zero)
    lngColorCount As Long  ' Number of colors (set to zero for 24 bits)
    lngImportantColors As Long ' "Important" colors (set to zero)
End Type

Type typPIXEL
    bytB As Byte    ' Blue
    bytG As Byte    ' Green
    bytR As Byte    ' Red
End Type

Type typBITMAPFILE
    bmfh As typHEADER
    bmfi As typINFOHEADER
    bmbits() As Byte
End Type

Sub testowy()
    Dim bmpFile As typBITMAPFILE
    Dim lngRowSize As Long
    Dim lngPixelArraySize As Long
    Dim lngFileSize As Long
    Dim j, k, l, x As Long
    Dim w, h As Long
    Dim bytRed, bytGreen, bytBlue As Integer
    Dim lngRGBColor() As Long
    Dim sHexCol As String
    Dim vRGB() As String

    Dim strBMP As String
    DebugOP "BMP Output Started"
    w = 480
    h = 546

    With bmpFile

        With .bmfh
            .strType = "BM"
            .lngSize = 0
            .intRes1 = 0
            .intRes2 = 0
            .lngOffset = 54
        End With
        With .bmfi
            .lngSize = 40
            .lngWidth = w
            .lngHeight = h
            .intPlanes = 1
            .intBits = 24
            .lngCompression = 0
            .lngImageSize = 0
            .lngxResolution = 0
            .lngyResolution = 0
            .lngColorCount = 0
            .lngImportantColors = 0
        End With
        lngRowSize = Round(.bmfi.intBits * .bmfi.lngWidth / 32) * 4
        lngPixelArraySize = lngRowSize * .bmfi.lngHeight

        ReDim .bmbits(lngPixelArraySize)
        ReDim lngRGBColor(w, h)
        
        k = -1
        For j = h To 1 Step -1
        ' For each row, starting at the bottom and working up...
            'each column starting at the left
            For x = 1 To w
'                sHexCol = Nz(ELookup("TerrainColor", "MapBMPq", "MAP = '" & _
'                            "XX " & Format(x, "00") & Format(j, "00") & "'"), "000000")
                sHexCol = Nz(ELookup("TerrainColor", "AA_BMPMap", _
                                    "Col = " & x & _
                                    " AND Row = " & j), "C2C2C2")

                vRGB = Split(HEXtoRGB(sHexCol), ",")
                'Debug.Print sHexCol & " - " & vRGB(0) & vRGB(1) & vRGB(2)
                
                    k = k + 1
                    .bmbits(k) = CInt(vRGB(2))
                    k = k + 1
                    .bmbits(k) = CInt(vRGB(1))
                    k = k + 1
                    .bmbits(k) = CInt(vRGB(0))

            Next x
    
            If (w * .bmfi.intBits / 8 < lngRowSize) Then   ' Add padding if required
                For l = w * .bmfi.intBits / 8 + 1 To lngRowSize
                    k = k + 1
                    .bmbits(k) = 0
                Next l
            End If
        Next j
        .bmfh.lngSize = 14 + 40 + lngPixelArraySize
     End With ' Defining bmpFile
    strBMP = "C:\TV\xxx.BMP"
    Open strBMP For Binary Access Write As 1 Len = 1
        Put 1, 1, bmpFile.bmfh
        Put 1, , bmpFile.bmfi
        Put 1, , bmpFile.bmbits
    Close
    MsgBox "BMP map exported"
    DebugOP "BMP Output Complete"
End Sub
