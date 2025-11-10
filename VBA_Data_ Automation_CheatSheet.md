# VBA Data Automation Cheat Sheet

> _Created by Fábio Silva._  
> A full-spectrum reference for analysts, BI developers, and automation engineers using **VBA for Excel, Power BI, Python & SQL integration.**

---

## Table of Contents

1. [Project Setup & Hygiene](#1-project-setup--hygiene)  
2. [Fast Ranges, Arrays, and Writing Back](#2-fast-ranges-arrays-and-writing-back)  
3. [Text, RegEx, and Parsing](#3-text-regex-and-parsing)  
4. [File I/O & Folder Automation](#4-file-io--folder-automation)  
5. [HTTP, JSON, and APIs](#5-http-json-and-apis)  
6. [ADO: Querying CSV & Excel](#6-ado-querying-csv--excel)  
7. [Power Query Orchestration](#7-power-query-orchestration)  
8. [PivotTables & Evaluate Tricks](#8-pivottables--evaluate-tricks)  
9. [Date/Time, Scheduling, and Batching](#9-datetime-scheduling-and-batching)  
10. [Outlook / Email Automation](#10-outlook--email-automation)  
11. [PowerPoint / Word Reporting](#11-powerpoint--word-reporting)  
12. [Events & Application Hooks](#12-events--application-hooks)  
13. [Data Quality & Cleaning](#13-data-quality--cleaning)  
14. [Grouping, Aggregation & Joins](#14-grouping-aggregation--joins)  
15. [Worksheet & Table Utilities](#15-worksheet--table-utilities)  
16. [Quality-of-Life Tricks](#16-quality-of-life-tricks)  
17. [Protection, Versioning & Config](#17-protection-versioning--config)  
18. [Robust CSV Export (UTF-8)](#18-robust-csv-export-utf-8)  
19. [Memory-Friendly Patterns](#19-memory-friendly-patterns)  
20. [Testing & Maintainability](#20-testing--maintainability)  
21. [VBA ⇄ Python Interop (Call Python from Excel)](#21-vba--python-interop-call-python-from-excel)  
22. [Python ⇄ Excel Interop (xlwings / COM)](#22-python--excel-interop-xlwings--com)  
23. [SQL Integration (Secure + Performant)](#23-sql-integration-secure--performant)  
24. [Power Query + SQL + VBA (Hybrid Pipelines)](#24-power-query--sql--vba-hybrid-pipelines)  
25. [Pandas → Excel → VBA Reporting](#25-pandas--excel--vba-reporting)  
26. [Environment, Security & Packaging](#26-environment-security--packaging)  
27. [Changelog](#27-changelog)  
28. [Environment Tested](#28-environment-tested)

---

## 1) Project Setup & Hygiene

### Turn off “Excel lag” (and always turn it back on) — *now restores all states*
```vba
' Stores & restores full application state (calc, events, screen, statusbar)
Private Type AppState
    Calc As XlCalculation
    ScreenUpdate As Boolean
    Events As Boolean
    StatusBar As Variant
End Type
Private st As AppState

Public Sub SpeedUp(Optional ByVal enable As Boolean = True)
    If enable Then
        With Application
            st.Calc = .Calculation
            st.ScreenUpdate = .ScreenUpdating
            st.Events = .EnableEvents
            st.StatusBar = .StatusBar
            .ScreenUpdating = False
            .EnableEvents = False
            .StatusBar = "Running…"
            .Calculation = xlCalculationManual
        End With
    Else
        Application.Calculate
        With Application
            .ScreenUpdating = st.ScreenUpdate
            .EnableEvents = st.Events
            .StatusBar = st.StatusBar
            .Calculation = st.Calc
        End With
    End If
End Sub
```

### Error handling + logging scaffold
```vba
Public Sub ExampleRunner()
    On Error GoTo EH
    SpeedUp True

    ' ... your code ...

CleanExit:
    SpeedUp False
    Exit Sub
EH:
    LogMsg "ERROR", "ExampleRunner", Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub

Public Sub LogMsg(ByVal level As String, ByVal where As String, ByVal msg As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss"), level, where, msg
    ' Optionally append to a hidden "LOG" sheet:
    ' With ThisWorkbook.Sheets("LOG")
    '     .Cells(.Rows.Count, 1).End(xlUp).Offset(1).Resize(1, 4).Value = Array(Now, level, where, msg)
    ' End With
End Sub
```

### Common References
- **Microsoft Scripting Runtime** — `Dictionary`, `FileSystemObject`
- **Microsoft VBScript Regular Expressions 5.5**
- **Microsoft XML, v6.0** (or WinHTTP Services)
- **ADO 2.x** — database/CSV connections
- **Outlook/PowerPoint/Word Libraries** — for Office automation

---

## 2) Fast Ranges, Arrays, and Writing Back

### Read → Array → Process → Write Back
```vba
Public Sub RangeToArrayToRange()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Data")
    Dim arr As Variant, r As Long

    arr = ws.Range("A1").CurrentRegion.Value2
    For r = 2 To UBound(arr, 1)
        arr(r, 2) = UCase$(CStr(arr(r, 2)))
    Next
    ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value2 = arr
End Sub
```

### Column Index by Header
```vba
Public Function ColIndex(ByVal header As String, ByVal headerRange As Range) As Long
    Dim m As Variant
    m = Application.Match(header, headerRange, 0)
    If IsError(m) Then Err.Raise 5, , "Header not found: " & header
    ColIndex = CLng(m)
End Function
```

### Unique Values (Dictionary)
```vba
Public Function UniqueValues(ByVal rng As Range) As Variant
    Dim d As Scripting.Dictionary: Set d = New Scripting.Dictionary
    Dim v, i&
    v = rng.Value2
    For i = 1 To UBound(v, 1)
        If Not d.Exists(CStr(v(i, 1))) Then d.Add CStr(v(i, 1)), Empty
    Next
    UniqueValues = d.Keys
End Function
```

---

## 3) Text, RegEx, and Parsing

### Regex Replace
```vba
Public Function RegexReplace(ByVal text As String, ByVal pattern As String, ByVal repl As String) As String
    Dim re As VBScript_RegExp_55.RegExp
    Set re = New VBScript_RegExp_55.RegExp
    re.Global = True: re.MultiLine = True: re.IgnoreCase = True
    re.Pattern = pattern
    RegexReplace = re.Replace(text, repl)
End Function
```

### Split CSV Line Safely
```vba
Public Function SplitCsv(ByVal line As String) As Variant
    Dim re As New VBScript_RegExp_55.RegExp, m As VBScript_RegExp_55.Match, out() As String, i&
    re.Pattern = "(""[^""]*""|[^,]*)"
    re.Global = True
    For Each m In re.Execute(line)
        ReDim Preserve out(i)
        out(i) = Trim$(m.SubMatches(0))
        If Left$(out(i), 1) = """" And Right$(out(i), 1) = """" Then
            out(i) = Mid$(out(i), 2, Len(out(i)) - 2)
            out(i) = Replace(out(i), """""", """")
        End If
        i = i + 1
    Next
    SplitCsv = out
End Function
```

---

## 4) File I/O & Folder Automation

### FileSystemObject Basics
```vba
Public Sub FsoDemo()
    Dim fso As New Scripting.FileSystemObject
    Dim f As Scripting.TextStream, path$
    path = ThisWorkbook.Path & "\output.txt"
    Set f = fso.CreateTextFile(path, True, False)
    f.WriteLine "Hello, log."
    f.Close
End Sub
```

### Pick Folder Dialog
```vba
Public Function PickFolder(Optional ByVal title As String = "Select folder") As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function
```

### Loop Files
```vba
Public Sub ForEachCsvInFolder(ByVal folder As String)
    Dim f As String
    f = Dir$(folder & "\*.csv")
    Do While Len(f) > 0
        Debug.Print "Got:", f
        f = Dir$
    Loop
End Sub
```

---

## 5) HTTP, JSON, and APIs

### HTTP GET (with timeouts)
```vba
Public Function HttpGet(ByVal url As String) As String
    Dim x As New MSXML2.ServerXMLHTTP60
    x.setTimeouts 5000, 5000, 5000, 10000 ' resolve, connect, send, receive (ms)
    x.Open "GET", url, False
    x.send
    If x.Status <> 200 Then Err.Raise 5, , "HTTP " & x.Status & " " & x.statusText
    HttpGet = x.responseText
End Function
```

### HTTP POST with JSON Body (with timeouts)
```vba
Public Function HttpPostJson(ByVal url As String, ByVal json As String) As String
    Dim x As New MSXML2.ServerXMLHTTP60
    x.setTimeouts 5000, 5000, 5000, 10000
    x.Open "POST", url, False
    x.setRequestHeader "Content-Type", "application/json"
    x.send json
    If x.Status < 200 Or x.Status >= 300 Then Err.Raise 5, , "HTTP " & x.Status & " " & x.statusText
    HttpPostJson = x.responseText
End Function
```

### Parse JSON (using VBA-JSON)
```vba
Public Function JsonExample()
    Dim t$, j As Object
    t = HttpGet("https://api.github.com/repos/octocat/Hello-World")
    Set j = JsonConverter.ParseJson(t)
    Debug.Print j("full_name"), j("stargazers_count")
End Function
```

---

## 6) ADO: Querying CSV & Excel

> **Note on UTF-8 CSV via ACE**: `CharacterSet=65001` can be flaky. Prefer **Power Query**, a `schema.ini`, or read via `ADODB.Stream` then parse.

### Query a CSV Folder
```vba
Public Sub QueryCsvFolder()
    Dim cn As Object, rs As Object, sql$, folder$
    folder = "C:\Data\csvfolder"
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & folder & ";" & _
            "Extended Properties=""text;HDR=Yes;FMT=Delimited;CharacterSet=65001;"""
    sql = "SELECT * FROM [myfile.csv] WHERE [Country]='PT'"
    Set rs = cn.Execute(sql)
    Sheet1.Range("A1").CopyFromRecordset rs
    rs.Close: cn.Close
End Sub
```

### Query a Closed Excel File
```vba
Public Sub QueryClosedXlsx()
    Dim cn As Object, rs As Object, p$, sql$
    p = "C:\Data\Sales.xlsx"
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p & ";" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1;"""
    sql = "SELECT * FROM [SalesData$A1:Z100000]"
    Set rs = cn.Execute(sql)
    Sheet1.Range("A1").CopyFromRecordset rs
    rs.Close: cn.Close
End Sub
```

---

## 7) Power Query Orchestration

### Refresh All, wait safely, then post-process
```vba
Public Sub RefreshAllAndPost()
    ThisWorkbook.RefreshAll
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
    Do
        DoEvents
    Loop While Application.CalculationState <> xlDone Or Not Application.Ready
    ' Post-processing...
End Sub
```

### Refresh a specific PQ query and read result table (now safe on empty tables)
```vba
Public Sub RefreshQueryToTable(ByVal queryName As String, ByVal listObjectName As String, ByVal ws As Worksheet)
    ws.ListObjects(listObjectName).Refresh
    Dim lo As ListObject: Set lo = ws.ListObjects(listObjectName)

    If lo.DataBodyRange Is Nothing Then
        Debug.Print "Rows: 0"
        Exit Sub
    End If

    Dim arr
    arr = lo.DataBodyRange.Value2
    Debug.Print "Rows:", UBound(arr, 1)
End Sub
```

---

## 8) PivotTables & Evaluate Tricks

### Create Pivot from a Table
```vba
Public Sub MakePivot()
    Dim wsSrc As Worksheet: Set wsSrc = Sheet1
    Dim wsPvt As Worksheet: Set wsPvt = Sheet2
    Dim pc As PivotCache, pt As PivotTable

    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, wsSrc.ListObjects(1).Range)
    Set pt = pc.CreatePivotTable(TableDestination:=wsPvt.Range("A3"), TableName:="PT1")

    With pt
        .PivotFields("Country").Orientation = xlRowField
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Sales").Orientation = xlDataField
        .PivotFields("Sales").Function = xlSum
        .RowAxisLayout xlTabularRow
    End With
End Sub
```

### Evaluate for Vectorized Operations
```vba
Sheet1.Range("C2:C1001").Value = Evaluate("IF(A2:A1001>10, A2:A1001*B2:B1001, 0)")
```

---

## 9) Date/Time, Scheduling, and Batching

### ISO Format
```vba
Public Function ToISO(ByVal d As Date) As String
    ToISO = Format$(d, "yyyy-mm-dd\Thh:nn:ss")
End Function
```

### Timers
```vba
Private nextRun As Date
Public Sub StartEvery15Min()
    nextRun = Now + TimeSerial(0, 15, 0)
    Application.OnTime nextRun, "Job"
End Sub
Public Sub Job()
    ' do stuff...
    StartEvery15Min ' reschedule
End Sub
Public Sub StopTimer()
    On Error Resume Next
    Application.OnTime nextRun, "Job", , False
End Sub
```

---

## 10) Outlook / Email Automation

### Send Email with Attachment
```vba
Public Sub SendMail()
    Dim ol As Outlook.Application, mail As Outlook.MailItem
    Set ol = New Outlook.Application
    Set mail = ol.CreateItem(olMailItem)
    With mail
        .To = "team@company.com"
        .CC = ""
        .Subject = "Daily Report"
        .Body = "Attached."
        .Attachments.Add ThisWorkbook.FullName
        .Send  ' or .Display
    End With
End Sub
```

---

## 11) PowerPoint / Word Reporting

### Export Range to PowerPoint (no Selection dependency)
```vba
Public Sub RangeToPpt()
    Dim rng As Range: Set rng = Sheet1.Range("A1:F20")
    rng.CopyPicture xlScreen, xlPicture

    Dim ppt As PowerPoint.Application, pres As PowerPoint.Presentation, slide As PowerPoint.Slide
    Dim shp As PowerPoint.Shape

    Set ppt = New PowerPoint.Application
    ppt.Visible = msoTrue
    Set pres = ppt.Presentations.Add
    Set slide = pres.Slides.Add(1, ppLayoutBlank)

    Set shp = slide.Shapes.Paste(1)
    shp.Align msoAlignCenters, True
    shp.Align msoAlignMiddles, True

    pres.SaveAs ThisWorkbook.Path & "\report.pptx"
    pres.Close: ppt.Quit
End Sub
```

---

## 12) Events & Application Hooks

### Workbook Events
```vba
Private Sub Workbook_Open()
    LogMsg "INFO", "Workbook_Open", "Hello!"
End Sub
```

### Application Events
```vba
' ==== Class Module: CAppEvents ====
Public WithEvents App As Application
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Debug.Print "Changed:", Sh.Name, Target.Address
End Sub

' ==== Standard Module ====
Public AE As CAppEvents
Public Sub HookApp()
    Set AE = New CAppEvents
    Set AE.App = Application
End Sub
```

---

## 13) Data Quality & Cleaning

### Null/Blank Normalizer
```vba
Public Function Nz(ByVal v As Variant, Optional ByVal fallback As Variant = "") As Variant
    If IsError(v) Then Nz = fallback ElseIf LenB(v & vbNullString) = 0 Then Nz = fallback Else Nz = v
End Function
```

### Trim & Collapse Whitespace
```vba
Public Sub CleanColumn(ByVal rng As Range)
    Dim arr, r&, s$
    arr = rng.Value2
    For r = 1 To UBound(arr, 1)
        s = CStr(arr(r, 1))
        s = WorksheetFunction.Trim(RegexReplace(s, "\s+", " "))
        arr(r, 1) = s
    Next
    rng.Value2 = arr
End Sub
```

---

## 14) Grouping, Aggregation & Joins (Dictionary-Based)

### Group by Key and Sum
```vba
Public Function GroupSum(ByVal keys As Range, ByVal vals As Range) As Object
    Dim d As New Scripting.Dictionary, i&, k$, v, a, b
    a = keys.Value2: b = vals.Value2
    For i = 1 To UBound(a, 1)
        k = CStr(a(i, 1)): v = CDbl(Nz(b(i, 1), 0))
        d(k) = IIf(d.Exists(k), d(k) + v, v)
    Next
    Set GroupSum = d
End Function
```

### Join Two Tables by Key (Left Join)
```vba
Public Sub LeftJoin()
    Dim ws As Worksheet: Set ws = Sheet1
    Dim L, R, i&, d As New Scripting.Dictionary, k$, colsL&, colsR&, out(), r&, c&

    L = ws.Range("A1").CurrentRegion.Value2 ' Left table
    R = ws.Range("H1").CurrentRegion.Value2 ' Right table

    For i = 2 To UBound(R, 1)
        d(CStr(R(i, 1))) = Application.Index(R, i, 0)
    Next

    colsL = UBound(L, 2): colsR = UBound(R, 2)
    ReDim out(1 To UBound(L, 1), 1 To colsL + colsR - 1)

    For c = 1 To colsL: out(1, c) = L(1, c): Next
    For c = 2 To colsR: out(1, colsL + c - 1) = R(1, c): Next

    For r = 2 To UBound(L, 1)
        For c = 1 To colsL: out(r, c) = L(r, c): Next
        k = CStr(L(r, 1))
        If d.Exists(k) Then
            Dim rowR: rowR = d(k): rowR = rowR ' coerce layout across versions
            For c = 2 To colsR: out(r, colsL + c - 1) = rowR(1, c): Next
        End If
    Next
    Sheet2.Range("A1").Resize(UBound(out, 1), UBound(out, 2)).Value2 = out
End Sub
```

---

## 15) Worksheet & Table Utilities

### Clear Table Body, Keep Headers
```vba
Public Sub ClearTableBody(ByVal lo As ListObject)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
End Sub
```

### Add or Get a Table Safely
```vba
Public Function EnsureTable(ByVal ws As Worksheet, ByVal name As String, ByVal rng As Range) As ListObject
    On Error Resume Next
    Set EnsureTable = ws.ListObjects(name)
    On Error GoTo 0
    If EnsureTable Is Nothing Then
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        EnsureTable.Name = name
    End If
End Function
```

---

## 16) Quality-of-Life Tricks

### Status Bar Progress (Keep Excel Responsive)
```vba
Public Sub ProgressDemo()
    Dim i&
    For i = 1 To 100
        Application.StatusBar = "Processing " & i & "%"
        DoEvents
    Next
    Application.StatusBar = False
End Sub
```

### Copy Values Only (Safe Clipboard)
```vba
Public Sub CopyValuesOnly(ByVal src As Range, ByVal dst As Range)
    dst.Resize(src.Rows.Count, src.Columns.Count).Value2 = src.Value2
End Sub
```

**Tip:** Always use `.Value2` — it’s faster and avoids unwanted date conversions.

---

## 17) Protection, Versioning & Config

### Configuration Sheet Pattern
```vba
Public Function Cfg(ByVal key As String) As String
    Dim m
    m = Application.Match(key, Sheets("_CONFIG").Range("A:A"), 0)
    If IsError(m) Then Err.Raise 5, , "Config key not found: " & key
    Cfg = Sheets("_CONFIG").Cells(m, 2).Value2
End Function
```

### Semantic Version
```vba
Public Const APP_VERSION As String = "1.4.2"
```

---

## 18) Robust CSV Export (UTF-8)

```vba
' True UTF-8 export using ADODB.Stream
Public Sub ExportUtf8(ByVal rng As Range, ByVal path As String)
    Dim r&, c&, arr, line$, s$, txt As String
    arr = rng.Value2
    For r = 1 To UBound(arr, 1)
        line = vbNullString
        For c = 1 To UBound(arr, 2)
            s = CStr(arr(r, c))
            If InStr(s, """") Or InStr(s, ",") Or InStr(s, vbLf) Then
                s = """" & Replace(s, """", """""") & """"
            End If
            If Len(line) = 0 Then line = s Else line = line & "," & s
        Next
        txt = txt & line & vbCrLf
    Next

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2          ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText txt
        .SaveToFile path, 2 ' adSaveCreateOverWrite
        .Close
    End With
End Sub
```

---

## 19) Memory-Friendly Patterns

- Use **arrays + one write-back** instead of looping cells.  
- Turn off updates with `SpeedUp True`.  
- Avoid `.Select`, `.Activate`, `.Copy` — reference objects directly.  
- Split large loops into batches (e.g., 100k rows) and include `DoEvents`.  

---

## 20) Testing & Maintainability

### Module Layout Best Practice
```
modMain
modArray
modHttp
modAdo
modLog
modConfig
```

### Naming Convention
- **Subs** → verbs (`LoadData`, `ExportCsv`)  
- **Functions** → nouns (`GetToken`, `Cfg`)

### Guard Clauses
Validate early, fail fast:
```vba
If ws Is Nothing Then Err.Raise 5, , "Worksheet not set"
```

### Rubberduck VBA
If available, use it for:
- Unit testing  
- Code inspections  
- Code metrics  

### Mini Starter Template
```vba
Public Sub RunETL()
    On Error GoTo EH
    SpeedUp True
    LogMsg "INFO", "RunETL", "v" & APP_VERSION

    Dim inFolder$, outCsv$, dataSh As Worksheet
    inFolder = Cfg("DATA_FOLDER")
    outCsv = ThisWorkbook.Path & "\export.csv"
    Set dataSh = ThisWorkbook.Sheets("Data")

    ' 1) Load (e.g., ADO or readers)
    ' 2) Transform (arrays, regex, joins)
    ' 3) Output CSV / Power BI table

    ExportUtf8 dataSh.Range("A1").CurrentRegion, outCsv
    LogMsg "INFO", "RunETL", "Exported: " & outCsv

CleanExit:
    SpeedUp False
    Exit Sub
EH:
    LogMsg "ERROR", "RunETL", Err.Number & " - " & Err.Description
    Resume CleanExit
End Sub
```

---

## 21) VBA ⇄ Python Interop (Call Python from Excel)

### Resolve Python path dynamically + run and capture output
```vba
Private Function PythonCmd() As String
    Dim p As String
    p = Environ$("LOCALAPPDATA") & "\Programs\Python\Python312\python.exe"
    If Dir$(p) <> "" Then
        PythonCmd = """" & p & """"
    Else
        PythonCmd = "py -3"
    End If
End Function

Private Function Q(ByVal s As String) As String
    Q = """" & s & """"
End Function

Public Function RunPython(ByVal scriptPath As String, Optional ByVal args As String = "") As String
    Dim sh As Object, cmd$, exec As Object, out$, err$
    Set sh = CreateObject("WScript.Shell")

    cmd = PythonCmd() & " " & Q(scriptPath)
    If Len(args) > 0 Then cmd = cmd & " " & args  ' pass args already quoted when needed

    Set exec = sh.Exec(cmd)

    Do Until exec.StdOut.AtEndOfStream
        out = out & exec.StdOut.ReadLine & vbCrLf
    Loop
    Do Until exec.StdErr.AtEndOfStream
        err = err & exec.StdErr.ReadLine & vbCrLf
    Loop

    If exec.ExitCode <> 0 Then Err.Raise 5, , "Python failed:" & vbCrLf & err
    RunPython = out
End Function
```

### JSON data exchange (UTF-8 consistent with ADODB.Stream)
```vba
Public Function CallPythonJson(py As String, ByVal inJson As String) As String
    Dim inPath$, outPath$, args$
    inPath = Environ$("TEMP") & "\in.json"
    outPath = Environ$("TEMP") & "\out.json"

    ' Write UTF-8
    Dim stmW As Object: Set stmW = CreateObject("ADODB.Stream")
    With stmW
        .Type = 2          ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText inJson
        .SaveToFile inPath, 2 ' adSaveCreateOverWrite
        .Close
    End With

    args = Q(inPath) & " " & Q(outPath)
    Dim _out As String
    _out = RunPython(py, args)

    ' Read UTF-8
    Dim stmR As Object: Set stmR = CreateObject("ADODB.Stream")
    With stmR
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile outPath
        CallPythonJson = .ReadText(-1)
        .Close
    End With
End Function
```

**Example Python Script (`process.py`):**
```python
import sys, json, pandas as pd
inp, outp = sys.argv[1], sys.argv[2]
data = json.load(open(inp, encoding="utf-8"))
df = pd.DataFrame(data["rows"])
df["Total"] = df["Qty"] * df["Price"]
json.dump({"ok": True, "rows": df.to_dict(orient="records")}, open(outp, "w", encoding="utf-8"))
```

---

## 22) Python ⇄ Excel Interop (xlwings / COM)

### xlwings: seamless bi-directional Excel access
```python
# pip install xlwings
import xlwings as xw

def write_totals():
    wb = xw.Book.caller()
    sht = wb.sheets["Data"]
    vals = sht["A1"].expand().value
    sht["G1"].value = "Total"
    sht["G2"].options(transpose=True).value = [sum(row[1:3]) for row in vals[1:]]
```

**How to call from Excel:**
Use the xlwings add-in and macro:
```vba
Sub RunPythonTotals()
    RunPython ("import script; script.write_totals()")
End Sub
```

### COM Automation from Python (win32com)
```python
import win32com.client as win32
xl = win32.Dispatch("Excel.Application")
wb = xl.Workbooks.Open(r"C:\path\file.xlsm")
xl.Application.Run("Module1.RunETL")
wb.Close(SaveChanges=True)
xl.Quit()
```

---

## 23) SQL Integration (Secure + Performant)

### Parameterized Stored Procedure (safe & explicit types)
```vba
Public Sub UpsertCustomer(ByVal id As Long, ByVal name As String, ByVal email As String)
    Dim cn As Object, cmd As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=MSOLEDBSQL;Server=YOURSERVER;Database=YOURDB;Trusted_Connection=Yes;"
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.UpsertCustomer"

    cmd.Parameters.Append cmd.CreateParameter("@Id", adInteger, adParamInput, , id)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarWChar, adParamInput, 200, name)
    cmd.Parameters.Append cmd.CreateParameter("@Email", adVarWChar, adParamInput, 320, email)

    cmd.Execute
    cn.Close
End Sub
```

### Common DSN-less connection strings
- **SQL Server:** `Provider=MSOLEDBSQL;Server=SRV;Database=DB;Trusted_Connection=Yes;`
- **PostgreSQL:** `Driver={PostgreSQL Unicode(x64)};Server=host;Port=5432;Database=db;Uid=u;Pwd=p;`
- **MySQL:** `Driver={MySQL ODBC 8.0 ANSI Driver};Server=host;Database=db;User=u;Password=p;Option=3;`

### Bulk Load Pattern (ETL staging)
1. Export CSV with `ExportUtf8`  
2. `BULK INSERT` into staging table  
3. Use `MERGE` for upsert into final table  
4. Clear staging after commit  

**SQL Server MERGE Example:**
```sql
MERGE dbo.Customer AS tgt
USING (SELECT Id, Name, Email FROM dbo.Customer_stg) AS src
ON tgt.Id = src.Id
WHEN MATCHED THEN UPDATE SET Name = src.Name, Email = src.Email
WHEN NOT MATCHED THEN INSERT (Id, Name, Email) VALUES (src.Id, src.Name, src.Email);
```

---

## 24) Power Query + SQL + VBA (Hybrid Pipelines)

**Recommended orchestration pattern:**
1. Refresh Power Query (`RefreshAllAndPost`)
2. Wait until queries complete
3. Validate data quality (using section 13)
4. Export with `ExportUtf8`
5. (Optional) Send results to SQL via parameterized ADO commands

This hybrid architecture leverages:
- **Power Query** → ingestion & transformation  
- **VBA** → orchestration & automation  
- **SQL** → persistence, analytics, and governance  

---

## 25) Pandas → Excel → VBA Reporting

Combine the strengths of Python ETL + VBA visualization.

**Python ETL Example:**
```python
import pandas as pd

df = pd.read_csv("data.csv")
df["Total"] = df["Qty"] * df["Price"]

with pd.ExcelWriter("output.xlsx", engine="xlsxwriter") as xw:
    df.to_excel(xw, sheet_name="Data", index=False)
```

Then open in Excel and use your VBA macros:
- `MakePivot` → build summary pivots  
- `RangeToPpt` → auto-export to PowerPoint  

---

## 26) Environment, Security & Packaging

- Use **environment variables** (`Environ$("VAR")` in VBA, `os.environ` in Python) for secrets.  
- Store configuration in `_CONFIG` sheet + `Cfg("KEY")`.  
- Split code by layer: `/src`, `/examples`, `/docs`.  
- Centralize logs (hidden sheet or lightweight SQLite).  
- Testing stack: **Rubberduck (VBA)** + **pytest (Python)**.  
- Document requirements in `requirements.txt` and this `README.md`.

---

## 27) Changelog

- **2025-11-09**  
  - Fixed `ExportUtf8` signature (`ByVal path As String`)  
  - `RefreshQueryToTable` guard for empty tables  
  - `CallPythonJson` now UTF-8 consistent with `ADODB.Stream`  
  - `RunPython` quoting + StdErr capture  
  - `RangeToPpt` no `Selection` dependency; centered shape  
  - `SplitCsv` uses `VBScript_RegExp_55.Match`  
  - Minor clarifications and comments

---

## 28) Environment Tested

- **Excel**: Microsoft 365 (build 2409) — Windows 11  
- **VBA**: VBE with references: Scripting, RegEx 5.5, MSXML 6.0, ADO 2.x, Outlook/PowerPoint (optional)  
- **Python**: 3.12 (Windows), `pandas`, `xlwings`, `pywin32`  
- **SQL**: MSOLEDBSQL to SQL Server 2019+

---

## Author

**Fábio Silva**  
Senior Data & BI Analyst | Automation Specialist | Renewable Energy Data Architect  
[LinkedIn](https://www.linkedin.com/in/efabio-jss) • [GitHub](https://github.com/efabio-jss)

---

## License
MIT License — free to use, share, and adapt with attribution.
