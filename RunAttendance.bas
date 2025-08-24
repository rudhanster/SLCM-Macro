Attribute VB_Name = "RunAttendanceModule"
Option Explicit

' ====== EDIT THESE TWO PATHS ======
Private Const PYTHON_EXE As String = "C:\Path\To\Python\python.exe"
Private Const PY_SCRIPT  As String = "C:\Path\To\maa.py"
' ==================================

Public Sub RunAttendanceForActiveWorkbook()
    Dim theDate As String, wbPath As String
    Dim absenteesList As String, subjectDetails As String
    Dim cmd As String, args As String
    Dim preview As String
    
    ' --- Read selected date (prefer m/d/yyyy so Python treats X/Y/Z as MM/DD/YYYY) ---
    theDate = GetSelectedDateString()
    If Len(theDate) = 0 Then
        MsgBox "Please select the date cell (Attendance column header) and try again.", vbExclamation
        Exit Sub
    End If
    
    ' --- Workbook path ---
    wbPath = ThisWorkbook.FullName
    If Len(wbPath) = 0 Then
        MsgBox "Please save the workbook and try again.", vbExclamation
        Exit Sub
    End If
    
    ' --- Gather absentees for that date ---
    absenteesList = GetAbsenteesForDate(theDate)
    If absenteesList = "ERROR" Then
        MsgBox "Could not find absentees for the selected date.", vbExclamation
        Exit Sub
    End If
    
    ' --- Read subject details from Initial Setup (B1..B5) ---
    subjectDetails = GetSubjectDetails()
    If subjectDetails = "ERROR" Then
        MsgBox "Could not read subject details from 'Initial Setup' sheet.", vbExclamation
        Exit Sub
    End If
    
    ' --- Quick confirmation prompt (optional) ---
    preview = "📅 Date: " & theDate & vbCrLf & vbCrLf & _
              "📂 Workbook: " & wbPath & vbCrLf & vbCrLf
    If Len(absenteesList) = 0 Then
        preview = preview & "✅ No absentees found." & vbCrLf
    Else
        preview = preview & "❌ Absentees (" & CountCsv(absenteesList) & "): " & absenteesList & vbCrLf
    End If
    preview = preview & vbCrLf & "Proceed to update SLCM attendance?"
    
    If MsgBox(preview, vbQuestion + vbOKCancel, "Confirm") <> vbOK Then Exit Sub
    
    ' --- Build command line (properly quoted) ---
    args = Join(Array( _
        QuoteArg(theDate), _
        QuoteArg(wbPath), _
        QuoteArg(absenteesList), _
        QuoteArg(subjectDetails) _
    ), " ")
    
    cmd = QuoteArg(PYTHON_EXE) & " " & QuoteArg(PY_SCRIPT) & " " & args
    
    ' --- Validate paths exist ---
    If Dir(PYTHON_EXE, vbNormal) = "" Then
        MsgBox "Python not found at:" & vbCrLf & PYTHON_EXE & vbCrLf & vbCrLf & "Edit PYTHON_EXE at the top of the module.", vbCritical
        Exit Sub
    End If
    If Dir(PY_SCRIPT, vbNormal) = "" Then
        MsgBox "Python script not found at:" & vbCrLf & PY_SCRIPT & vbCrLf & vbCrLf & "Edit PY_SCRIPT at the top of the module.", vbCritical
        Exit Sub
    End If
    
    ' --- Launch in a visible console so you can complete SSO and see logs ---
    ' /K keeps the window open; use /C to close after it finishes
    Shell "cmd.exe /K " & cmd, vbNormalFocus
End Sub


' ===== Helpers =====

' Returns selected cell as m/d/yyyy when it's a date; otherwise the cell text trimmed.
Private Function GetSelectedDateString() As String
    On Error GoTo Fallback
    If TypeName(Selection) = "Range" Then
        Dim v As Variant
        v = Selection.Cells(1, 1).Value
        If IsDate(v) Then
            ' Force US-style m/d/yyyy so Python treats X/Y/Z as MM/DD/YYYY
            GetSelectedDateString = Format$(CDate(v), "m/d/yyyy")
            Exit Function
        Else
            Dim s As String
            s = Trim$(CStr(v))
            If Len(s) > 0 Then
                GetSelectedDateString = s
                Exit Function
            End If
        End If
    End If
Fallback:
    GetSelectedDateString = ""
End Function

' Quote a single argument for Windows command line (handles spaces & embedded quotes)
Private Function QuoteArg(ByVal s As String) As String
    If Len(s) = 0 Then
        QuoteArg =  &  ' empty ""
        Exit Function
    End If
    ' Escape internal quotes by doubling them
    s = Replace$(s, , "")
    QuoteArg =  & s & 
End Function

' Count CSV elements (simple split on commas)
Private Function CountCsv(ByVal csv As String) As Long
    Dim arr As Variant
    If Len(Trim$(csv)) = 0 Then
        CountCsv = 0
    Else
        arr = Split(csv, ",")
        CountCsv = UBound(arr) - LBound(arr) + 1
    End If
End Function

' ==== Your sheet parsers (same logic as Mac) ====

' Reads absentees for the selected date from the "Attendance" sheet.
' - Expects headers in row 2; "Reg. No." column somewhere in row 2; dates from col 7 onward.
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
    
    ' Find the date column by scanning row 2
    dateCol = 0
    For i = 1 To ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        cellValue = CStr(ws.Cells(2, i).Value)
        If InStr(1, cellValue, selectedDate, vbTextCompare) > 0 Or cellValue = selectedDate Then
            dateCol = i
            Exit For
        End If
    Next i
    If dateCol = 0 Then
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Find "Reg. No." column (fuzzy match: contains "REG" and "NO")
    regNoCol = 0
    For i = 1 To ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
        cellValue = UCase$(CStr(ws.Cells(2, i).Value))
        If (InStr(1, cellValue, "REG", vbTextCompare) > 0) And (InStr(1, cellValue, "NO", vbTextCompare) > 0) Then
            regNoCol = i
            Exit For
        End If
    Next i
    If regNoCol = 0 Then
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Collect absentees from row 3 downward
    lastRow = ws.Cells(ws.Rows.Count, regNoCol).End(xlUp).Row
    absentees = ""
    For i = 3 To lastRow
        cellValue = UCase$(CStr(ws.Cells(i, dateCol).Value))
        If cellValue = "AB" Or cellValue = "ABSENT" Then
            Dim regNo As String
            regNo = CStr(ws.Cells(i, regNoCol).Value)
            ' Strip trailing ".0" or anything after a dot, if the column is numeric-formatted
            regNo = Split(regNo, ".")(0)
            regNo = Trim$(regNo)
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

' Reads subject details from "Initial Setup" sheet:
'   B1: Course Name
'   B2: Course Code
'   B3: Semester
'   B4: Class Section (e.g., B or B-1)
'   B5: Session No (optional; leave blank if not used)
Private Function GetSubjectDetails() As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim courseName As String, courseCode As String
    Dim semester As String, classSection As String, sessionNo As String
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Initial Setup")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetSubjectDetails = "ERROR"
        Exit Function
    End If
    
    courseName = CleanOneLine(CStr(ws.Cells(1, 2).Value))    ' B1
    courseCode = CleanOneLine(CStr(ws.Cells(2, 2).Value))    ' B2
    semester = CleanOneLine(CStr(ws.Cells(3, 2).Value))      ' B3
    classSection = CleanOneLine(CStr(ws.Cells(4, 2).Value))  ' B4
    sessionNo = CleanOneLine(CStr(ws.Cells(5, 2).Value))     ' B5
    
    GetSubjectDetails = courseName & "|" & courseCode & "|" & semester & "|" & classSection & "|" & sessionNo
    Exit Function

ErrorHandler:
    GetSubjectDetails = "ERROR"
End Function

' Remove CR/LF and trim (avoids breaking command-line args)
Private Function CleanOneLine(ByVal s As String) As String
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    CleanOneLine = Trim$(s)
End Function
