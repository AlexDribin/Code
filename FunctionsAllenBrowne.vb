Attribute VB_Name = "FunctionsAllenBrowne"
Option Compare Database
Option Explicit

Public Function ELookup(Expr As String, Domain As String, Optional Criteria As Variant, _
    Optional OrderClause As Variant) As Variant
On Error GoTo Err_ELookup
    'Purpose:   Faster and more flexible replacement for DLookup()
    'Arguments: Same as DLookup, with additional Order By option.
    'Return:    Value of the Expr if found, else Null.
    '           Delimited list for multi-value field.
    'Author:    Allen Browne. allen@allenbrowne.com
    'Updated:   December 2006, to handle multi-value fields (Access 2007 and later.)
    'Examples:
    '           1. To find the last value, include DESC in the OrderClause, e.g.:
    '               ELookup("[Surname] & [FirstName]", "tblClient", , "ClientID DESC")
    '           2. To find the lowest non-null value of a field, use the Criteria, e.g.:
    '               ELookup("ClientID", "tblClient", "Surname Is Not Null" , "Surname")
    'Note:      Requires a reference to the DAO library.
    Dim db As DAO.Database          'This database.
    Dim rs As DAO.Recordset         'To retrieve the value to find.
    Dim rsMVF As DAO.Recordset      'Child recordset to use for multi-value fields.
    Dim varResult As Variant        'Return value for function.
    Dim strSQL As String            'SQL statement.
    Dim strOut As String            'Output string to build up (multi-value field.)
    Dim lngLen As Long              'Length of string.
    Const strcSep = ","             'Separator between items in multi-value list.

    'Initialize to null.
    varResult = Null

    'Build the SQL string.
    strSQL = "SELECT TOP 1 " & Expr & " FROM " & Domain
    If Not IsMissing(Criteria) Then
        strSQL = strSQL & " WHERE " & Criteria
    End If
    If Not IsMissing(OrderClause) Then
        strSQL = strSQL & " ORDER BY " & OrderClause
    End If
    strSQL = strSQL & ";"

    'Lookup the value.
    Set db = DBEngine(0)(0)
    Set rs = db.OpenRecordset(strSQL, dbOpenForwardOnly)
    If rs.RecordCount > 0 Then
        'Will be an object if multi-value field.
        If VarType(rs(0)) = vbObject Then
            Set rsMVF = rs(0).Value
            Do While Not rsMVF.EOF
                If rs(0).TYPE = 101 Then        'dbAttachment
                    strOut = strOut & rsMVF!fileName & strcSep
                Else
                    strOut = strOut & rsMVF![Value].Value & strcSep
                End If
                rsMVF.MoveNext
            Loop
            'Remove trailing separator.
            lngLen = Len(strOut) - Len(strcSep)
            If lngLen > 0& Then
                varResult = Left(strOut, lngLen)
            End If
            Set rsMVF = Nothing
        Else
            'Not a multi-value field: just return the value.
            varResult = rs(0)
        End If
    End If
    rs.Close

    'Assign the return value.
    ELookup = varResult

Exit_ELookup:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_ELookup:
    MsgBox Err.Description, vbExclamation, "ELookup Error " & Err.NUMBER
    Resume Exit_ELookup
End Function

Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWhere As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = ", ") As Variant
On Error GoTo Err_Handler
    'Purpose:   Generate a concatenated string of related records.
    'Return:    String variant, or Null if no matches.
    'Arguments: strField = name of field to get results from and concatenate.
    '           strTable = name of a table or query.
    '           strWhere = WHERE clause to choose the right values.
    '           strOrderBy = ORDER BY clause, for sorting the values.
    '           strSeparator = characters to use between the concatenated values.
    'Notes:     1. Use square brackets around field/table names with spaces or odd characters.
    '           2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
    '           3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
    '           4. Returning more than 255 characters to a recordset triggers this Access bug:
    '               http://allenbrowne.com/bug-16.html
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).TYPE > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.NUMBER & ": " & Err.Description, vbExclamation, "ConcatRelated()"
    Resume Exit_Handler
End Function

Public Function ConcatRelatedSQL(strSQL As String, _
                                Optional strSeparator = ", ") As Variant
On Error GoTo Err_Handler
    'Purpose:   Generate a concatenated string of related records.
    'Return:    String variant, or Null if no matches.
    'Arguments: strField = name of field to get results from and concatenate.
    '           strTable = name of a table or query.
    '           strWhere = WHERE clause to choose the right values.
    '           strOrderBy = ORDER BY clause, for sorting the values.
    '           strSeparator = characters to use between the concatenated values.
    'Notes:     1. Use square brackets around field/table names with spaces or odd characters.
    '           2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
    '           3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
    '           4. Returning more than 255 characters to a recordset triggers this Access bug:
    '               http://allenbrowne.com/bug-16.html
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
   
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelatedSQL = ""
    
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).TYPE > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelatedSQL = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.NUMBER & ": " & Err.Description, vbExclamation, "ConcatRelatedSQL()"
    Resume Exit_Handler
End Function


