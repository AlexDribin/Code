Attribute VB_Name = "ColourConvert"
Option Compare Database
'Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : RGBtoHEX
' DateTime  : 2006-Nov-17 13:58
' Author    : CARDA Consultants Inc. - Main
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function RGBtoHEX(r As Byte, G As Byte, B As Byte) As String
On Error GoTo RGBtoHEX_Error

    If r < 16 Then
        HEX1 = 0 & HEX(r)
    Else
        HEX1 = HEX(r)
    End If
    
    If G < 16 Then
        HEX2 = 0 & HEX(G)
    Else
        HEX2 = HEX(G)
    End If
    
    If B < 16 Then
        HEX3 = 0 & HEX(B)
    Else
        HEX3 = HEX(B)
    End If
    
    RGBtoHEX = HEX1 & HEX2 & HEX3

ExitFunction:
   Exit Function

RGBtoHEX_Error:

    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.NUMBER & vbCrLf & "Error Source: ColorConvert / RGBtoHEX" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo ExitFunction

End Function

Function HEXtoLong(HEX As String) As Long
On Error GoTo HEXtoLong_Error

    r = CByte("&H" & Left(HEX, 2))
    G = CByte("&H" & Mid(HEX, 3, 2))
    B = CByte("&H" & Mid(HEX, 5, 2))

    HEXtoLong = r + (G * 256) + (B * 65536)

ExitFunction:
   Exit Function

HEXtoLong_Error:

    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.NUMBER & vbCrLf & "Error Source: ColorConvert / HEXtoLong" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo ExitFunction
End Function



'---------------------------------------------------------------------------------------
' Procedure : HEXtoRGB
' DateTime  : 2006-Nov-17 14:07
' Author    : CARDA Consultants Inc. - Main
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function HEXtoRGB(HEX As String) As String
On Error GoTo HEXtoRGB_Error

    r = CByte("&H" & Left(HEX, 2))
    G = CByte("&H" & Mid(HEX, 3, 2))
    B = CByte("&H" & Mid(HEX, 5, 2))

HEXtoRGB = r & "," & G & "," & B

ExitFunction:
   Exit Function

HEXtoRGB_Error:

    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.NUMBER & vbCrLf & "Error Source: ColorConvert / HEXtoRGB" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo ExitFunction
End Function

'---------------------------------------------------------------------------------------
' Procedure : HEXtoRGB
' DateTime  : 2006-Nov-17 14:07
' Author    : CARDA Consultants Inc. - Main
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function HEXtoFontColor(HEX As String) As Long
On Error GoTo HEXtoFontColor_Error

    Dim sFontColor As String

    r = CByte("&H" & Left(HEX, 2))
    G = CByte("&H" & Mid(HEX, 3, 2))
    B = CByte("&H" & Mid(HEX, 5, 2))

    If (r + G + B < 350) And (G < 200) Then
        sFontColor = "FFFFFF"
    Else
        sFontColor = "000000"
    End If
    
    HEXtoFontColor = HEXtoLong(sFontColor)

ExitFunction:
   Exit Function

HEXtoFontColor_Error:

    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.NUMBER & vbCrLf & "Error Source: ColorConvert / HEXtoFontColor" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo ExitFunction
End Function

Sub tt1923847()
    Debug.Print HEXtoFontColor("000000")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RGBtoOLE
' DateTime  : 2006-Nov-17 14:23
' Author    : CARDA Consultants Inc. - Main
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function RGBtoOLE(r As Long, G As Long, B As Long) As Long
On Error GoTo RGBtoOLE_Error

    RGBtoOLE = r + (G * 256) + (B * 65536)

ExitFunction:
   Exit Function

RGBtoOLE_Error:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.NUMBER & vbCrLf & "Error Source: ColorConvert / RGBtoOLE" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo ExitFunction
End Function

'---------------------------------------------------------------------------------------
' Procedure : OLEtoRGB
' DateTime  : 2006-Nov-17 14:37
' Author    : CARDA Consultants Inc. - Main
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function OLEtoRGB(OLE As Long) As String
On Error GoTo OLEtoRGB_Error

    r = OLE And 255
    G = (OLE \ 256) And 255
    B = (OLE \ 65536) And 255
    
    OLEtoRGB = r & "," & G & "," & B

ExitFunction:
   Exit Function

OLEtoRGB_Error:

    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.NUMBER & vbCrLf & "Error Source: ColorConvert / OLEtoRGB" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occured!"
    GoTo ExitFunction
End Function


' Convert an RGB value into an HLS value.
Function RgbToHls(ByVal r As Double, ByVal G As Double, _
    ByVal B As Double, ByRef h As Double, ByRef l As _
    Double, ByRef s As Double)
'Source: http://www.vb-helper.com/howto_rgb_to_hls.html
Dim max As Double
Dim min As Double
Dim diff As Double
Dim r_dist As Double
Dim g_dist As Double
Dim b_dist As Double

    ' Get the maximum and minimum RGB components.
    max = r
    If max < G Then max = G
    If max < B Then max = B

    min = r
    If min > G Then min = G
    If min > B Then min = B

    diff = max - min
    l = (max + min) / 2
    If Abs(diff) < 0.00001 Then
        s = 0
        h = 0   ' H is really undefined.
    Else
        If l <= 0.5 Then
            s = diff / (max + min)
        Else
            s = diff / (2 - max - min)
        End If

        r_dist = (max - r) / diff
        g_dist = (max - G) / diff
        b_dist = (max - B) / diff

        If r = max Then
            h = b_dist - g_dist
        ElseIf G = max Then
            h = 2 + r_dist - b_dist
        Else
            h = 4 + g_dist - r_dist
        End If

        h = h * 60
        If h < 0 Then h = h + 360
    End If
End Function


' Convert an HLS value into an RGB value.
Function HlsToRgb(ByVal h As Double, ByVal l As Double, _
    ByVal s As Double, ByRef r As Double, ByRef G As _
    Double, ByRef B As Double)
'Source: http://www.vb-helper.com/howto_rgb_to_hls.html
Dim p1 As Double
Dim p2 As Double

    If l <= 0.5 Then
        p2 = l * (1 + s)
    Else
        p2 = l + s - l * s
    End If
    p1 = 2 * l - p2
    If s = 0 Then
        r = l
        G = l
        B = l
    Else
        r = QqhToRgb(p1, p2, h + 120)
        G = QqhToRgb(p1, p2, h)
        B = QqhToRgb(p1, p2, h - 120)
    End If
End Function

Function QqhToRgb(ByVal q1 As Double, ByVal q2 As _
    Double, ByVal hue As Double) As Double
'Source: http://www.vb-helper.com/howto_rgb_to_hls.html
    If hue > 360 Then
        hue = hue - 360
    ElseIf hue < 0 Then
        hue = hue + 360
    End If
    If hue < 60 Then
        QqhToRgb = q1 + (q2 - q1) * hue / 60
    ElseIf hue < 180 Then
        QqhToRgb = q2
    ElseIf hue < 240 Then
        QqhToRgb = q1 + (q2 - q1) * (240 - hue) / 60
    Else
        QqhToRgb = q1
    End If
End Function


