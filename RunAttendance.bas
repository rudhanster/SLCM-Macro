Option Explicit

' ====== EDIT THESE TWO PATHS ======
Private Const PYTHON_EXE As String = "C:\Users\anirudhanadukkathaya\AppData\Local\Microsoft\WindowsApps\python.exe"
Private Const PY_SCRIPT  As String = "C:\Mac\Home\Documents\win\maa.py"
' ==================================

' Delimiter used between subject fields (VBA ? Python)
Private Const DETAILS_DELIM As String = "|"

' ==============================================================
' == MAIN ENTRY ================================================
' ==============================================================

Public Sub RunAttendanceForActiveWorkbook()
    Dim theDate As String, wbPath As String
    Dim absenteesList As String, subjectDetails As String
    Dim args As String, runLine As String
    Dim preview As String

    ' --- Get selected date ---
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

    ' --- Absentees ---
    absenteesList = GetAbsenteesForDate(theDate)
    If Left$(absenteesList, 2) = "E:" Then
        MsgBox Mid$(absenteesList, 3), vbExclamation, "Attendance date detection"
        Exit Sub
    End If

    ' --- Subject details ---
    subjectDetails = GetSubjectDetails()
    If subjectDetails = "ERROR" Then Exit Sub

    ' --- Confirmation (plain text, no emojis) ---
    preview = "Date: " & theDate & vbCrLf & _
              "Workbook: " & wbPath & vbCrLf & vbCrLf
    If Len(absenteesList) = 0 Then
        preview = preview & "No absentees found." & vbCrLf
    Else
        preview = preview & "Absentees (" & CountCsv(absenteesList) & "): " & absenteesList & vbCrLf
    End If
    preview = preview & vbCrLf & "Proceed to update SLCM attendance?"
    If MsgBox(preview, vbQuestion + vbOKCancel, "Confirm") <> vbOK Then Exit Sub

    ' --- Path check ---
    If Dir(PYTHON_EXE, vbNormal) = "" Then
        MsgBox "Python not found: " & PYTHON_EXE, vbCritical
        Exit Sub
    End If
    If Dir(PY_SCRIPT, vbNormal) = "" Then
        MsgBox "Python script not found: " & PY_SCRIPT, vbCritical
        Exit Sub
    End If

    ' --- Create a temporary batch file (most reliable approach) ---
    Dim tempDir As String
    Dim batFile As String
    
    tempDir = Environ("TEMP")
    batFile = tempDir & "\slcm_automation_" & Format(Now, "yyyymmdd_hhnnss") & ".bat"
    
    ' Build batch file content
    Dim batContent As String
    batContent = "@echo off" & vbCrLf
    batContent = batContent & "title SLCM Attendance Automation" & vbCrLf
    batContent = batContent & "echo Starting SLCM Attendance Automation..." & vbCrLf
    batContent = batContent & "echo." & vbCrLf
    batContent = batContent & """" & PYTHON_EXE & """ """ & PY_SCRIPT & """ """ & theDate & """ """ & wbPath & """ """ & absenteesList & """ """ & subjectDetails & """" & vbCrLf
    batContent = batContent & "echo." & vbCrLf
    batContent = batContent & "echo Script execution completed." & vbCrLf
    batContent = batContent & "echo Press any key to close this window..." & vbCrLf
    batContent = batContent & "pause > nul" & vbCrLf
    batContent = batContent & "del """ & batFile & """" & vbCrLf
    
    ' Write batch file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open batFile For Output As #fileNum
    Print #fileNum, batContent
    Close #fileNum
    
    ' Execute batch file
    runLine = """" & batFile & """"
    
    Debug.Print "Executing: " & runLine

    ' --- Launch ---
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run runLine, 1, False ' 1 = Normal window, False = Don't wait
    

End Sub

' ==============================================================
' == POWERSHELL HELPERS ========================================
' ==============================================================

' Properly quote strings for PowerShell
Private Function PowerShellQuote(ByVal s As String) As String
    ' Escape single quotes by doubling them, then wrap in single quotes
    PowerShellQuote = "'" & Replace(s, "'", "''") & "'"
End Function

' Escape arguments for PowerShell ArgumentList (comma-separated)
Private Function EscapeForPowerShell(ByVal s As String) As String
    ' For ArgumentList, we need to quote each argument properly
    EscapeForPowerShell = PowerShellQuote(s)
End Function

' ==============================================================
' == ORIGINAL HELPERS ==========================================
' ==============================================================

' Returns selected cell as m/d/yyyy when it's a date; otherwise plain text
Private Function GetSelectedDateString() As String
    On Error GoTo Fallback
    If TypeName(Selection) = "Range" Then
        Dim v As Variant: v = Selection.Cells(1, 1).Value
        If IsDate(v) Then
            GetSelectedDateString = Format$(CDate(v), "m/d/yyyy")
            Exit Function
        Else
            Dim s As String: s = Trim$(CStr(v))
            If Len(s) > 0 Then GetSelectedDateString = s: Exit Function
        End If
    End If
Fallback:
    GetSelectedDateString = ""
End Function

' Count CSV elements
Private Function CountCsv(ByVal csv As String) As Long
    If Len(Trim$(csv)) = 0 Then
        CountCsv = 0
    Else
        CountCsv = UBound(Split(csv, ",")) + 1
    End If
End Function

' Find absentees for a given date
Private Function GetAbsenteesForDate(ByVal selectedDate As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Attendance")
    If ws Is Nothing Then
        GetAbsenteesForDate = "E: Sheet 'Attendance' not found.": Exit Function
    End If

    Dim headerRow As Long: headerRow = 2
    Dim dateCol As Long: dateCol = FindDateColumn(ws, headerRow, selectedDate)
    If dateCol = 0 Then
        GetAbsenteesForDate = "E: Could not find date column matching '" & selectedDate & "'."
        Exit Function
    End If

    Dim regNoCol As Long, c As Long
    For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If InStr(1, UCase$(CStr(ws.Cells(headerRow, c).Value)), "REG") > 0 And _
           InStr(1, UCase$(CStr(ws.Cells(headerRow, c).Value)), "NO") > 0 Then
            regNoCol = c: Exit For
        End If
    Next
    If regNoCol = 0 Then
        GetAbsenteesForDate = "E: Could not find 'Reg. No.' column.": Exit Function
    End If

    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, regNoCol).End(xlUp).Row
    Dim absentees As String, r As Long, val As String, regNo As String
    For r = headerRow + 1 To lastRow
        val = UCase$(Trim$(CStr(ws.Cells(r, dateCol).Value)))
        If val = "AB" Or val = "ABSENT" Then
            regNo = Trim$(CStr(ws.Cells(r, regNoCol).Value))
            If InStr(regNo, ".") > 0 Then
                regNo = Split(regNo, ".")(0) ' strip trailing .0 if any
            End If
            If absentees <> "" Then absentees = absentees & ","
            absentees = absentees & regNo
        End If
    Next

    GetAbsenteesForDate = absentees
End Function

' Find the correct date column in header row
Private Function FindDateColumn(ws As Worksheet, headerRow As Long, selectedDate As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, v As Variant
    For c = 1 To lastCol
        v = ws.Cells(headerRow, c).Value
        If IsDate(v) And IsDate(selectedDate) Then
            If DateValue(v) = DateValue(CDate(selectedDate)) Then FindDateColumn = c: Exit Function
        ElseIf CStr(v) = selectedDate Then
            FindDateColumn = c: Exit Function
        End If
    Next
End Function

' Build subject details string from Initial Setup sheet
Private Function GetSubjectDetails() As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Initial Setup")
    If ws Is Nothing Then
        MsgBox "Sheet 'Initial Setup' not found.", vbCritical
        GetSubjectDetails = "ERROR": Exit Function
    End If

    Dim courseName As String:   courseName = Trim$(CStr(ws.Cells(1, 2).Value)) ' B1
    Dim courseCode As String:   courseCode = Trim$(CStr(ws.Cells(2, 2).Value)) ' B2
    Dim semester As String:     semester = Trim$(CStr(ws.Cells(3, 2).Value))   ' B3
    Dim classSection As String: classSection = Trim$(CStr(ws.Cells(4, 2).Value)) ' B4
    Dim sessionNo As String:    sessionNo = Trim$(CStr(ws.Cells(5, 2).Value))  ' B5 (optional)

    If Len(courseCode) = 0 Or Len(semester) = 0 Or Len(classSection) = 0 Then
        MsgBox "Please fill B2 (Course Code), B3 (Semester), and B4 (Class Section) on 'Initial Setup'.", vbCritical, "Subject details incomplete"
        GetSubjectDetails = "ERROR": Exit Function
    End If

    ' Ensure the delimiter never appears in the fields
    Dim badDelim As String: badDelim = DETAILS_DELIM
    If InStr(courseName, badDelim) Or InStr(courseCode, badDelim) Or _
       InStr(semester, badDelim) Or InStr(classSection, badDelim) Or InStr(sessionNo, badDelim) Then
        MsgBox "Fields must not contain '" & DETAILS_DELIM & "'.", vbCritical
        GetSubjectDetails = "ERROR": Exit Function
    End If

    GetSubjectDetails = courseName & DETAILS_DELIM & courseCode & DETAILS_DELIM & _
                        semester & DETAILS_DELIM & classSection & DETAILS_DELIM & sessionNo
End Function


