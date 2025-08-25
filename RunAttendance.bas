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
        QuoteArg = """" & """"  ' -> ""
        Exit Function
    End If
    s = Replace$(s, """", """""") ' escape internal quotes
    QuoteArg = """" & s & """"
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


' -------- Robust header detection & date matching --------

' Try to auto-detect the header row by finding a "Reg. No."-like header
Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, lastCol As Long, c As Long
    For r = 1 To 5
        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastCol < 2 Then GoTo NextRow
        For c = 1 To lastCol
            Dim hdr As String
            hdr = UCase$(NormalizeHeaderText(CStr(ws.Cells(r, c).Value)))
            If (InStr(1, hdr, "REG", vbTextCompare) > 0) And (InStr(1, hdr, "NO", vbTextCompare) > 0) Then
                FindHeaderRow = r
                Exit Function
            End If
        Next c
NextRow:
    Next r
    FindHeaderRow = 0
End Function

' Normalize header text (remove CR/LF and trim)
Private Function NormalizeHeaderText(ByVal s As String) As String
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    NormalizeHeaderText = Trim$(s)
End Function

Private Function SameDay(d1 As Date, d2 As Date) As Boolean
    SameDay = (DateValue(d1) = DateValue(d2))
End Function

' Find the date column by comparing actual dates (ignores formatting/time)
Private Function FindDateColumn(ws As Worksheet, headerRow As Long, selectedDateText As String) As Long
    Dim lastCol As Long, c As Long
    Dim hdr As String, v As Variant
    Dim selIsDate As Boolean, selD As Date
    Dim hdrIsDate As Boolean, hdrD As Date
    
    FindDateColumn = 0
    
    selIsDate = IsDate(selectedDateText)
    If selIsDate Then selD = DateValue(CDate(selectedDateText))
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        v = ws.Cells(headerRow, c).Value
        hdr = NormalizeHeaderText(CStr(v))
        hdrIsDate = IsDate(v)
        If hdrIsDate Then hdrD = DateValue(CDate(v))
        
        ' 1) true date vs true date by day
        If selIsDate And hdrIsDate Then
            If SameDay(hdrD, selD) Then
                FindDateColumn = c
                Exit Function
            End If
        End If
        
        ' 2) selected is date; header is text — try common formats
        If selIsDate And Not hdrIsDate Then
            If InStr(1, hdr, Format$(selD, "m/d/yyyy"), vbTextCompare) > 0 _
            Or InStr(1, hdr, Format$(selD, "d-mmm-yy"), vbTextCompare) > 0 _
            Or InStr(1, hdr, Format$(selD, "dd-mmm-yyyy"), vbTextCompare) > 0 _
            Or InStr(1, hdr, Format$(selD, "mmmm d, yyyy"), vbTextCompare) > 0 _
            Or InStr(1, hdr, Format$(selD, "ddd, dd mmm yyyy"), vbTextCompare) > 0 Then
                FindDateColumn = c
                Exit Function
            End If
        End If
        
        ' 3) selected is text; fallback to tolerant contains
        If Not selIsDate Then
            If Len(selectedDateText) > 0 Then
                If InStr(1, hdr, selectedDateText, vbTextCompare) > 0 Or hdr = selectedDateText Then
                    FindDateColumn = c
                    Exit Function
                End If
            End If
        End If
    Next c
End Function


' Reads absentees for the selected date from the "Attendance" sheet.
' Works with your layout: headers on row 2; dates start at F2 onward; data from row 3.
Private Function GetAbsenteesForDate(ByVal selectedDate As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim dateCol As Long, regNoCol As Long
    Dim lastRow As Long, r As Long
    Dim c As Long
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
    
    ' Auto-detect header row (default to 2)
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then headerRow = 2
    
    ' Robust date column detection on headerRow
    dateCol = FindDateColumn(ws, headerRow, selectedDate)
    If dateCol = 0 Then
        Dim lastHeaderCol As Long, c As Long, hdr As String, sample As String
        lastHeaderCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        sample = ""
        For c = 1 To Application.Min(lastHeaderCol, 20)
            hdr = NormalizeHeaderText(CStr(ws.Cells(headerRow, c).Value))
            If Len(sample) < 700 Then sample = sample & "[" & c & "] " & hdr & vbCrLf
        Next
        MsgBox "Could not find date column matching: " & selectedDate & vbCrLf & _
               "(Header row detected as row " & headerRow & ")." & vbCrLf & vbCrLf & _
               "First headers scanned:" & vbCrLf & sample, vbExclamation
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Find "Reg. No." column on same header row
    regNoCol = 0
    For c = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        cellValue = UCase$(NormalizeHeaderText(CStr(ws.Cells(headerRow, c).Value)))
        If (InStr(1, cellValue, "REG", vbTextCompare) > 0) And (InStr(1, cellValue, "NO", vbTextCompare) > 0) Then
            regNoCol = c
            Exit For
        End If
    Next c
    If regNoCol = 0 Then
        MsgBox "Could not find 'Reg. No.' in header row " & headerRow & ".", vbExclamation
        GetAbsenteesForDate = "ERROR"
        Exit Function
    End If
    
    ' Collect absentees from rows below headerRow
    lastRow = ws.Cells(ws.Rows.Count, regNoCol).End(xlUp).Row
    absentees = ""
    For r = headerRow + 1 To lastRow
        cellValue = UCase$(Trim$(CStr(ws.Cells(r, dateCol).Value)))
        If cellValue = "AB" Or cellValue = "ABSENT" Then
            Dim regNo As String
            regNo = CStr(ws.Cells(r, regNoCol).Value)
            regNo = Split(regNo, ".")(0)
            regNo = Trim$(regNo)
            If Len(regNo) > 0 Then
                If Len(absentees) > 0 Then absentees = absentees & ","
                absentees = absentees & regNo
            End If
        End If
    Next r
    
    GetAbsenteesForDate = absentees
    Exit Function

ErrorHandler:
    GetAbsenteesForDate = "ERROR"
End Function


' Reads subject details from "Initial Setup" sheet: B1..B5
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


