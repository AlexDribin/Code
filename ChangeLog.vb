Attribute VB_Name = "ChangeLog"
Option Compare Database
Option Explicit

Public Function LastChangeLog()
    LastChangeLog = DLookup("VL_No", "A2_VersionLog", "VL_Date = #" & _
                    DMax("VL_Date", "A2_VersionLog") & "#") & " - " & _
                    DLookup("VL_Description", "A2_VersionLog", "VL_Date = #" & _
                    DMax("VL_Date", "A2_VersionLog") & "#")
End Function

Public Sub tt124()
    Debug.Print LastChangeLog()
End Sub
