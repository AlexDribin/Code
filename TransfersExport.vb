Attribute VB_Name = "TransfersExport"
Option Compare Database
Option Explicit

Public Sub Export_Transfers()
    Call ExportUnitTransfer(GetCurrentTurn(), "0636e9", "Andy1")
    Call ExportUnitTransfer(GetCurrentTurn(), "1636e1", "Andy2")
    Application.FollowHyperlink CurrentProject.Path & "\TransferExport\"
End Sub




Public Sub ExportUnitTransfer(sTurn As String, sUnit As String, sPlayer As String)

    'Create destination Folder if is doesn't exist
    Dim sDestinationFolder As String
    Dim sFileName As String
    sDestinationFolder = CurrentProject.Path & "\TransferExport\"
    If Len(Dir$(sDestinationFolder, vbDirectory)) > 0 Then
        Debug.Print "Folder exists - " & sDestinationFolder
    Else
        Debug.Print "Folder does not exist - " & sDestinationFolder
        MkDir sDestinationFolder
    End If
    
    Dim XL As Object
    Dim WB As Object
    Dim WKS As Object
    Dim WKS2 As Object
    
    Dim db As DAO.Database, rec As DAO.Recordset, f As DAO.field
    Dim i As Integer, j As Integer
    
    
    Set XL = New Excel.Application

    XL.Visible = False
    Set WB = XL.Workbooks.Add
    Set WKS = WB.Worksheets(1)
    Set WKS2 = WB.Sheets.Add
    WKS.Name = "Transfers"
    WKS2.Name = "GoodsHeld"


    '===Transfers===
    Dim rs As Recordset
    Dim sSQL As String
    
    sSQL = "SELECT [FROM], TO, ITEM, QUANTITY, TRANSFER_TIMING, NOTES, PROCESS_MSG " & _
        "FROM MassTransfers " & _
        "WHERE [FROM] = '" & _
        sUnit & "' " & _
        "OR TO = '" & _
        sUnit & _
        "' ORDER BY [FROM] ASC;"
    
    Set rs = CurrentDb.OpenRecordset(sSQL)
    WKS.Range("A2").CopyFromRecordset rs
    WKS.Range("A1").Value = "FROM"
    WKS.Range("B1").Value = "TO"
    WKS.Range("C1").Value = "ITEM"
    WKS.Range("D1").Value = "QUANTITY"
    WKS.Range("E1").Value = "TRANSFER_TIMING"
    WKS.Range("F1").Value = "NOTES"
    WKS.Range("G1").Value = "PROCESS_MSG"
    
    '===Goods Held===
    
    sSQL = "SELECT TRIBE, ITEM_TYPE, ITEM, ITEM_NUMBER " & _
        "FROM TRIBES_GOODS " & _
        "WHERE TRIBE = '" & _
        sUnit & "';"
    
    Set rs = CurrentDb.OpenRecordset(sSQL)
    WKS2.Range("A2").CopyFromRecordset rs
    WKS2.Range("A1").Value = "TRIBE"
    WKS2.Range("B1").Value = "ITEM_TYPE"
    WKS2.Range("C1").Value = "ITEM"
    WKS2.Range("D1").Value = "ITEM_NUMBER"
    
    
    WKS.Columns("A:G").AutoFit
    WKS2.Columns("A:D").AutoFit

    
    sFileName = sTurn & " - " & sUnit & " - " & sPlayer

    WB.SaveAs fileName:=sDestinationFolder & sFileName & ".xlsx", _
        AccessMode:=xlExclusive, _
        ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        
    rs.Close
    Set rs = Nothing

    WB.Close 'SaveChanges:=True
    Set WB = Nothing
    XL.Quit
    Set XL = Nothing
    
End Sub
