Option Explicit

' ====== EDIT THESE TWO PATHS ======
Private Const PYTHON_EXE As String = "C:\Windows\py.exe"   ' or full path to python.exe
Private Const PY_SCRIPT  As String = "C:\Mac\Home\Documents\win\maa.py"
' ==================================

' Delimiter used between subject fields (VBA ↔ Python)
Private Const DETAILS_DELIM As String = "|"

' ==============================================================
' == MAIN ENTRY ================================================
' ==============================================================

Public Sub RunAttendanceForActiveWorkbook()
    On Error GoTo Fail

    Dim theDate As String, wbPath As String
    Dim absenteesList As String, subjectDetails As String
    Dim preview As String, runLine As String

    ' --- Get selected date ---
    theDate = GetSelectedDateString()
    If Len(theDate) = 0 Then
        MsgBox "Please select the date cell (Attendance column header) and try again.", vbExclamation
        Exit Sub
    End If

    ' --- Workbook path (local copy if on SharePoint/OneDrive) ---
    wbPath = GetWorkbookPathForPython(ThisWorkbook)
    If Len(Dir$(wbPath, vbNormal)) = 0 Then
        MsgBox "Could not find or create a local copy of the workbook for Python." & vbCrLf & _
               "Path: " & wbPath, vbCritical
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

    ' --- Confirmation ---
    preview = "Date: " & theDate & vbCrLf & _
              "Workbook: " & wbPath & vbCrLf & vbCrLf
    If Len(absenteesList) = 0 Then
        preview = preview & "No absentees found." & vbCrLf
    Else
        preview = preview & "Absentees (" & CountCsv(absenteesList) & "): " & absenteesList & vbCrLf
    End If
    preview = preview & vbCrLf & "Proceed to update SLCM attendance?"
    If MsgBox(preview, vbQuestion + vbOKCancel, "Confirm") <> vbOK Then Exit Sub

    ' --- Path checks ---
    If Dir(PYTHON_EXE, vbNormal) = "" And LCase$(PYTHON_EXE) <> LCase$("C:\Windows\py.exe") Then
        MsgBox "Python not found: " & PYTHON_EXE & vbCrLf & "Tip: set to C:\Windows\py.exe", vbCritical
        Exit Sub
    End If
    If Dir(PY_SCRIPT, vbNormal) = "" Then
        MsgBox "Python script not found: " & PY_SCRIPT, vbCritical
        Exit Sub
    End If

    ' --- Build command line ---
    runLine = """" & PYTHON_EXE & """ """ & PY_SCRIPT & """ """ & theDate & _
              """ """ & wbPath & """ """ & absenteesList & """ """ & subjectDetails & """"

    Debug.Print "Executing: " & runLine

    ' --- Launch directly (no .bat, avoids Sophos block) ---
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.Run runLine, 1, False    ' 1=normal window, False=don’t wait

    Exit Sub

Fail:
    MsgBox "Error: " & Err.Description, vbCritical, "RunAttendanceForActiveWorkbook"
End Sub

' ==============================================================
' == HELPERS ===================================================
' ==============================================================

' Get a local file path for Python (handles SharePoint/OneDrive)
Private Function GetWorkbookPathForPython(ByVal wb As Workbook) As String
    Dim fullName As String, localName As String
    On Error Resume Next
    fullName = wb.FullName
    localName = wb.FullNameLocal
    On Error GoTo 0

    If Len(localName) > 0 And LCase$(Left$(localName, 4)) <> "http" Then
        GetWorkbookPathForPython = localName
        Exit Function
    End If

    ' Cloud-only: save a temp local copy
    Dim tempDir As String, tempPath As String, baseName As String, ext As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = "C:\Temp"

    baseName = wb.Name
    If InStrRev(baseName, ".") > 0 Then
        ext = Mid$(baseName, InStrRev(baseName, "."))
        baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    Else
        ext = ".xlsx"
    End If

    tempPath = tempDir & "\" & baseName & "_pycopy_" & Format(Now, "yyyymmdd_hhnnss") & ext
    wb.SaveCopyAs tempPath
    GetWorkbookPathForPython = tempPath
End Function

' Returns selected cell as m/d/yyyy if date; otherwise trimmed text
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
    Dim ws As Worksheet: On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Attendance")
    On Error GoTo 0
    If ws Is Nothing Then
        GetAbsenteesForDate = "E: Sheet 'Attendance' not found.": Exit Function
    End If

    Dim headerRow As Long: headerRow = 2
    Dim dateCol As Long: dateCol = FindDateColumn(ws, headerRow, selectedDate)
    If dateCol = 0 Then
        GetAbsenteesForDate = "E: Could not find date column matching '" & selectedDate & "'."
        Exit Function
    End If

    ' Locate "Reg. No." column
    Dim regNoCol As Long, c As Long
    For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        Dim hdr As String: hdr = UCase$(Trim$(CStr(ws.Cells(headerRow, c).Value)))
        If (InStr(hdr, "REG") > 0 And InStr(hdr, "NO") > 0) Or hdr = "REG. NO." Then
            regNoCol = c: Exit For
        End If
    Next
    If regNoCol = 0 Then
        GetAbsenteesForDate = "E: Could not find 'Reg. No.' column.": Exit Function
    End If

    ' Collect absentees
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, regNoCol).End(xlUp).Row
    Dim absentees As String, r As Long, val As String, regNo As String
    For r = headerRow + 1 To lastRow
        val = Trim$(CStr(ws.Cells(r, dateCol).Value))
        If LCase$(val) = "ab" Or LCase$(val) = "absent" Then
            regNo = Trim$(CStr(ws.Cells(r, regNoCol).Value))
            If InStr(regNo, ".") > 0 Then regNo = Split(regNo, ".")(0)
            If Len(regNo) > 0 Then
                If Len(absentees) > 0 Then absentees = absentees & ","
                absentees = absentees & regNo
            End If
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
        ElseIf Trim$(CStr(v)) = Trim$(selectedDate) Then
            FindDateColumn = c: Exit Function
        End If
    Next
End Function

' Build subject details string from Initial Setup sheet
Private Function GetSubjectDetails() As String
    Dim ws As Worksheet: On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Initial Setup")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'Initial Setup' not found.", vbCritical
        GetSubjectDetails = "ERROR": Exit Function
    End If

    Dim courseName As String:   courseName = Trim$(CStr(ws.Cells(1, 2).Value)) ' B1
    Dim courseCode As String:   courseCode = Trim$(CStr(ws.Cells(2, 2).Value)) ' B2
    Dim semester As String:     semester = Trim$(CStr(ws.Cells(3, 2).Value))   ' B3
    Dim classSection As String: classSection = Trim$(CStr(ws.Cells(4, 2).Value)) ' B4
    Dim sessionNo As String:    sessionNo = Trim$(CStr(ws.Cells(5, 2).Value))  ' B5 optional

    If Len(courseCode) = 0 Or Len(semester) = 0 Or Len(classSection) = 0 Then
        MsgBox "Please fill B2 (Course Code), B3 (Semester), and B4 (Class Section) on 'Initial Setup'.", vbCritical, "Subject details incomplete"
        GetSubjectDetails = "ERROR": Exit Function
    End If

    If InStr(courseName, DETAILS_DELIM) Or InStr(courseCode, DETAILS_DELIM) Or _
       InStr(semester, DETAILS_DELIM) Or InStr(classSection, DETAILS_DELIM) Or InStr(sessionNo, DETAILS_DELIM) Then
        MsgBox "Fields must not contain '" & DETAILS_DELIM & "'.", vbCritical
        GetSubjectDetails = "ERROR": Exit Function
    End If

    GetSubjectDetails = courseName & DETAILS_DELIM & courseCode & DETAILS_DELIM & _
                        semester & DETAILS_DELIM & classSection & DETAILS_DELIM & sessionNo
End Function
