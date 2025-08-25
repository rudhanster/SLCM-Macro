' Reads absentees for the selected date from the "Attendance" sheet.
' - Expects date headers in row 1; "Reg. No." column in row 1; dates from col F onward.
' - Student data starts from row 16.
' - Marks absentees when cell has "ab" or "absent" (case-insensitive).
Private Function GetAbsenteesForDate(ByVal selectedDate As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim dateCol As Long, regNoCol As Long
    Dim lastRow As Long, i As Long
    Dim absentees As String
    Dim cellValue As String
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Attendance")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Find the date column by scanning row 2 (where your dates are)
    dateCol = 0
    For i = 1 To ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        cellValue = CStr(ws.Cells(2, i).Value)
        ' Try to match the selected date with the header
        If InStr(1, cellValue, selectedDate, vbTextCompare) > 0 Or cellValue = selectedDate Then
            dateCol = i
            Exit For
        End If
        ' Also try formatting the cell as date and comparing
        On Error Resume Next
        If IsDate(ws.Cells(2, i).Value) Then
            Dim headerDate As String
            headerDate = Format(CDate(ws.Cells(2, i).Value), "m/d/yyyy")
            If headerDate = selectedDate Then
                dateCol = i
                Exit For
            End If
        End If
        On Error GoTo ErrorHandler
    Next i
    
    If dateCol = 0 Then
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Find "Reg. No." column in row 2 (fuzzy match: contains "REG" and "NO")
    regNoCol = 0
    For i = 1 To ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        cellValue = UCase(CStr(ws.Cells(2, i).Value))
        If (InStr(1, cellValue, "REG", vbTextCompare) > 0) And (InStr(1, cellValue, "NO", vbTextCompare) > 0) Then
            regNoCol = i
            Exit For
        End If
    Next i
    
    If regNoCol = 0 Then
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Collect absentees from row 3 downward (where your student data starts)
    lastRow = ws.Cells(ws.Rows.Count, regNoCol).End(xlUp).Row
    absentees = ""
    
    For i = 3 To lastRow
        cellValue = UCase(CStr(ws.Cells(i, dateCol).Value))
        If cellValue = "AB" Or cellValue = "A" Then
            Dim regNo As String
            regNo = CStr(ws.Cells(i, regNoCol).Value)
            ' Strip trailing ".0" or anything after a dot, if the column is numeric-formatted
            If InStr(regNo, ".") > 0 Then
                regNo = Split(regNo, ".")(0)
            End If
            regNo = Trim(regNo)
            If Len(regNo) > 0 Then
                If Len(absentees) > 0 Then absentees = absentees & ","
                absentees = absentees & regNo
            End If
        End If
    Next i
    
    GetAbsenteesForDate = absentees
    Exit Function

ErrorHandler:
    GetAbsenteesForDate = "ERROR"
End Function

