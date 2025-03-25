Attribute VB_Name = "Debugging"
Option Compare Database
Option Explicit
'=================================================================
'Used to send debug info to DebugOutput form
'
'andrew.d.bentley@gmail.com
'=================================================================
Public Sub DebugOP(s As String)
    Dim m As Long
    m = 999999
    If FormIsOpen("DebugOutput") Then
        Forms.DebugOutput.OutputTxt = Format(Now(), "HH:MM:SS") & " - " & s & vbCrLf & Forms.DebugOutput.OutputTxt
        If Len(Forms.DebugOutput.OutputTxt) > m Then
            Forms.DebugOutput.OutputTxt = Left(Forms.DebugOutput.OutputTxt, m)
        End If
        DoEvents
    End If
End Sub

Public Function FormIsOpen(ByVal strFormName As String) As Boolean

    FormIsOpen = False
    ' is form open?

    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> 0 Then
        ' if so make sure its not in design view
        If Forms(strFormName).CurrentView <> 0 Then
            FormIsOpen = True
        End If
    End If

    Exit Function

End Function

Public Sub WriteTxtFile(strFileName, strContent)
    'Use to save text in debugOutput form to a txt file
    Open strFileName For Output As #1
    Print #1, strContent
    Close #1
End Sub

Public Sub CalcRun()
    Dim sFirstTimeStamp As String
    Dim sLastTimeStamp As String
    Dim sDebugOP As String
    Dim tFirst As Date
    Dim tLast As Date
    Dim t As Double
    
    
    
    sDebugOP = Forms.DebugOutput.OutputTxt
    
    If Left(sDebugOP, 8) = "||======" Then
        Forms.DebugOutput.OutputTxt = "||======NO RUN SINCE LAST TIME CALC======||" & _
                                vbCrLf & _
                                sDebugOP
    Else
        sLastTimeStamp = Trim(Left(sDebugOP, 8))
        Debug.Print sLastTimeStamp
        tLast = TimeValue(sLastTimeStamp)
    
    
        Dim iLastRunPoint As Long
        iLastRunPoint = InStr(1, sDebugOP, "||======")
    
    
        sFirstTimeStamp = Trim(Mid(sDebugOP, InStrRev(sDebugOP, vbCrLf, InStrRev(sDebugOP, vbCrLf, iLastRunPoint - 1) - 1) + 2, 8))
        Debug.Print sFirstTimeStamp
        tFirst = TimeValue(sFirstTimeStamp)
        'Debug.Print Len(sFirstTimeStamp)
        
        
        t = Round(((tLast - tFirst) * 24 * 60 * 60), 0)
        Debug.Print t
        'MsgBox "====Run completed in: " & Format((t / 86400), "hh:nn:ss") & "===="
        
        Forms.DebugOutput.OutputTxt = "||======Run completed in: " & _
                                    Format((t / 86400), "hh:nn:ss") & _
                                    "======||" & _
                                    vbCrLf & _
                                    sDebugOP
    End If
    
    

    
    
End Sub

Public Sub TestDebugOP()
    Dim i As Long
    
    For i = 1 To 3000
        DebugOP (i) & " - Test string"
    Next i
    
End Sub


