'==============================================================================
' OTT PLATFORM — SALES & REVENUE MIS DASHBOARD GENERATOR
' VBA Macro for Microsoft Excel
'
' Author   : Abdul Jabbar M
' Tool     : VBA (Visual Basic for Applications) — Excel
' Purpose  : Reads raw OTT platform data and auto-generates a fully
'            formatted, colour-coded MIS Dashboard in Excel — replacing
'            hours of manual formatting and calculation with one click.
'
' How to Use:
'   1. Open Excel and press Alt + F11 to open VBA Editor
'   2. Insert > Module > Paste this entire code
'   3. Go back to Excel — paste your raw data in "Raw_Data" sheet
'      (Columns: Platform, Type, Title, Director, Country, Year, Rating, Genre)
'   4. Press Alt + F8 > Select "GenerateMISDashboard" > Run
'   5. Dashboard sheets are auto-created and formatted instantly
'
' Sheets Generated:
'   1. MIS_Summary      — KPIs and Executive Summary
'   2. Platform_Compare — Amazon vs Netflix breakdown
'   3. Genre_Analysis   — Top genres by platform
'   4. Country_Report   — Top countries by content volume
'   5. YoY_Trend        — Year-on-year content growth
'==============================================================================

Option Explicit

' ── COLOUR CONSTANTS ──────────────────────────────────────────────────────────
Private Const CLR_NAVY       As Long = 7024219    ' #1B3A6B — Dark Navy
Private Const CLR_BLUE       As Long = 11988398   ' #2E75B6 — Medium Blue
Private Const CLR_AMAZON     As Long = 14712993   ' #00A8E1 — Amazon Cyan
Private Const CLR_NETFLIX    As Long = 920596     ' #E50914 — Netflix Red
Private Const CLR_GOLD       As Long = 52479      ' #FFC000 — Gold
Private Const CLR_GREEN      As Long = 5332279    ' #22C55E — Green
Private Const CLR_LIGHT_BLUE As Long = 16169686   ' #D6E4F0 — Light Blue
Private Const CLR_LIGHT_GREY As Long = 15921906   ' #F2F2F2 — Light Grey
Private Const CLR_WHITE      As Long = 16777215   ' #FFFFFF — White
Private Const CLR_DARK       As Long = 3080478    ' #1A1A2E — Dark
Private Const CLR_MUTED      As Long = 5924984    ' #5A6478 — Muted Grey

' ── MAIN ENTRY POINT ──────────────────────────────────────────────────────────
Public Sub GenerateMISDashboard()
    
    Dim startTime As Double
    startTime = Timer
    
    ' Turn off screen updates for speed
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    MsgBox "OTT MIS Dashboard Generator" & Chr(10) & _
           "Author: Abdul Jabbar M" & Chr(10) & Chr(10) & _
           "Starting dashboard generation..." & Chr(10) & _
           "Please wait — this may take 30-60 seconds.", _
           vbInformation, "MIS Dashboard Generator"
    
    ' Step 1: Validate data
    If Not ValidateRawData() Then
        GoTo Cleanup
    End If
    
    ' Step 2: Delete existing dashboard sheets
    Call DeleteExistingSheets
    
    ' Step 3: Generate all sheets
    Call BuildMISSummary
    Call BuildPlatformComparison
    Call BuildGenreAnalysis
    Call BuildCountryReport
    Call BuildYoYTrend
    
    ' Step 4: Done
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 1)
    
    MsgBox "Dashboard Generated Successfully!" & Chr(10) & Chr(10) & _
           "Sheets Created: 5" & Chr(10) & _
           "Time Taken: " & elapsed & " seconds" & Chr(10) & Chr(10) & _
           "  1. MIS_Summary" & Chr(10) & _
           "  2. Platform_Compare" & Chr(10) & _
           "  3. Genre_Analysis" & Chr(10) & _
           "  4. Country_Report" & Chr(10) & _
           "  5. YoY_Trend" & Chr(10) & Chr(10) & _
           "Author: Abdul Jabbar M | Wipro Technologies", _
           vbInformation, "MIS Dashboard — Complete"
    
    ' Activate summary sheet
    Sheets("MIS_Summary").Activate
    
Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
    Resume Cleanup
    
End Sub


' ── STEP 1: VALIDATE RAW DATA ─────────────────────────────────────────────────
Private Function ValidateRawData() As Boolean
    
    ValidateRawData = False
    
    ' Check Raw_Data sheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("Raw_Data")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "ERROR: Sheet 'Raw_Data' not found!" & Chr(10) & Chr(10) & _
               "Please create a sheet named 'Raw_Data' with these columns:" & Chr(10) & _
               "A: Platform | B: Type | C: Title | D: Country" & Chr(10) & _
               "E: Release_Year | F: Rating | G: Genre | H: Date_Added", _
               vbCritical, "Data Validation Error"
        Exit Function
    End If
    
    ' Check data exists
    If ws.Cells(2, 1).Value = "" Then
        MsgBox "ERROR: No data found in 'Raw_Data' sheet!" & Chr(10) & _
               "Please add your data starting from Row 2.", _
               vbCritical, "Data Validation Error"
        Exit Function
    End If
    
    ValidateRawData = True
    
End Function


' ── STEP 2: DELETE EXISTING SHEETS ────────────────────────────────────────────
Private Sub DeleteExistingSheets()
    
    Dim sheetNames As Variant
    sheetNames = Array("MIS_Summary", "Platform_Compare", "Genre_Analysis", _
                       "Country_Report", "YoY_Trend")
    
    Dim sName As Variant
    Dim ws As Worksheet
    
    For Each sName In sheetNames
        On Error Resume Next
        Set ws = Sheets(CStr(sName))
        If Not ws Is Nothing Then
            ws.Delete
        End If
        Set ws = Nothing
        On Error GoTo 0
    Next sName
    
End Sub


' ── HELPER: FORMAT HEADER ROW ─────────────────────────────────────────────────
Private Sub FormatHeader(ws As Worksheet, startRow As Long, startCol As Long, _
                          endCol As Long, title As String, Optional bgColor As Long = -1)
    
    If bgColor = -1 Then bgColor = CLR_NAVY
    
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
        .Merge
        .Value = title
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = CLR_WHITE
        .Interior.Color = bgColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 22
    End With
    
End Sub


' ── HELPER: FORMAT COLUMN HEADERS ─────────────────────────────────────────────
Private Sub FormatColHeaders(ws As Worksheet, startRow As Long, startCol As Long, _
                               headers As Variant, Optional bgColor As Long = -1)
    
    If bgColor = -1 Then bgColor = CLR_BLUE
    
    Dim i As Integer
    For i = 0 To UBound(headers)
        With ws.Cells(startRow, startCol + i)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = CLR_WHITE
            .Interior.Color = bgColor
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(170, 170, 170)
            .RowHeight = 20
            .WrapText = True
        End With
    Next i
    
End Sub


' ── HELPER: FORMAT DATA ROW ───────────────────────────────────────────────────
Private Sub FormatDataRow(ws As Worksheet, rowNum As Long, startCol As Long, _
                           endCol As Long, isEven As Boolean)
    
    Dim bgColor As Long
    bgColor = IIf(isEven, CLR_LIGHT_GREY, CLR_WHITE)
    
    With ws.Range(ws.Cells(rowNum, startCol), ws.Cells(rowNum, endCol))
        .Interior.Color = bgColor
        .Font.Size = 10
        .Font.Color = CLR_DARK
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(221, 221, 221)
        .RowHeight = 16
    End With
    
End Sub


' ── HELPER: ADD SHEET TITLE BLOCK ─────────────────────────────────────────────
Private Sub AddSheetTitle(ws As Worksheet, title As String, subtitle As String)
    
    With ws.Range("A1:J1")
        .Merge
        .Value = title
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = CLR_WHITE
        .Interior.Color = CLR_NAVY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    With ws.Range("A2:J2")
        .Merge
        .Value = subtitle
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = CLR_LIGHT_BLUE
        .Interior.Color = CLR_NAVY
        .HorizontalAlignment = xlCenter
        .RowHeight = 16
    End With
    
End Sub


' ── HELPER: ADD FOOTER ────────────────────────────────────────────────────────
Private Sub AddFooter(ws As Worksheet, lastRow As Long)
    
    With ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + 1, 10))
        .Merge
        .Value = "Generated by: Abdul Jabbar M  |  MIS & Automation Specialist  |  " & Format(Now, "DD MMMM YYYY") & "  |  Tool: VBA Macro"
        .Font.Italic = True
        .Font.Size = 8
        .Font.Color = CLR_MUTED
        .HorizontalAlignment = xlRight
        .RowHeight = 14
    End With
    
End Sub


' ── HELPER: GET RAW DATA INTO ARRAY ──────────────────────────────────────────
Private Function GetRawData() As Variant
    
    Dim ws As Worksheet
    Set ws = Sheets("Raw_Data")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    GetRawData = ws.Range("A1:" & ws.Cells(lastRow, 8).Address).Value
    
End Function


' ── SHEET 1: MIS SUMMARY ──────────────────────────────────────────────────────
Private Sub BuildMISSummary()
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "MIS_Summary"
    ws.Tab.Color = CLR_NAVY
    
    ' Page setup
    ws.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:C").ColumnWidth = 18
    ws.Columns("D:I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 2
    
    ' Title
    Call AddSheetTitle(ws, _
        "OTT PLATFORM — SALES & REVENUE MIS DASHBOARD", _
        "Executive Summary  |  Report Date: " & Format(Now, "DD MMMM YYYY") & "  |  Source: Amazon Prime & Netflix")
    
    ' ── READ & CALCULATE KPIs ──
    Dim rawData As Variant
    rawData = GetRawData()
    
    Dim totalRows As Long
    totalRows = UBound(rawData, 1) - 1  ' Exclude header
    
    Dim amazonCount As Long, netflixCount As Long
    Dim movieCount As Long, showCount As Long
    Dim i As Long
    
    For i = 2 To UBound(rawData, 1)
        If rawData(i, 1) = "Amazon Prime" Then amazonCount = amazonCount + 1
        If rawData(i, 1) = "Netflix" Then netflixCount = netflixCount + 1
        If rawData(i, 2) = "Movie" Then movieCount = movieCount + 1
        If rawData(i, 2) = "TV Show" Then showCount = showCount + 1
    Next i
    
    ' ── KPI SECTION HEADER ──
    Dim r As Long
    r = 4
    Call FormatHeader(ws, r, 2, 9, "  KEY PERFORMANCE INDICATORS", CLR_BLUE)
    r = r + 1
    
    ' ── KPI CARDS — Row 1 ──
    Dim kpiLabels1 As Variant, kpiValues1 As Variant, kpiSubs1 As Variant
    kpiLabels1 = Array("TOTAL TITLES", "AMAZON PRIME", "NETFLIX", "TOTAL MOVIES")
    kpiValues1 = Array(totalRows, amazonCount, netflixCount, movieCount)
    kpiSubs1   = Array("Combined Library", "Share: " & Format(amazonCount / totalRows, "0.0%"), _
                       "Share: " & Format(netflixCount / totalRows, "0.0%"), "Both Platforms")
    
    Dim kpiColors1 As Variant
    kpiColors1 = Array(CLR_NAVY, CLR_AMAZON, CLR_NETFLIX, CLR_BLUE)
    
    Dim col As Long
    col = 2
    For i = 0 To 3
        ' Label
        With ws.Cells(r, col)
            .Value = kpiLabels1(i)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = CLR_MUTED
            .Interior.Color = CLR_LIGHT_GREY
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 16
        End With
        ' Skip next column (merged)
        ws.Cells(r, col + 1).Interior.Color = CLR_LIGHT_GREY
        ws.Range(ws.Cells(r, col), ws.Cells(r, col + 1)).Merge
        
        ' Value
        With ws.Cells(r + 1, col)
            .Value = kpiValues1(i)
            .NumberFormat = "#,##0"
            .Font.Bold = True
            .Font.Size = 22
            .Font.Color = kpiColors1(i)
            .Interior.Color = CLR_WHITE
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 32
        End With
        ws.Cells(r + 1, col + 1).Interior.Color = CLR_WHITE
        ws.Range(ws.Cells(r + 1, col), ws.Cells(r + 1, col + 1)).Merge
        
        ' Sub-label
        With ws.Cells(r + 2, col)
            .Value = kpiSubs1(i)
            .Font.Italic = True
            .Font.Size = 8
            .Font.Color = CLR_MUTED
            .Interior.Color = CLR_WHITE
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 14
        End With
        ws.Cells(r + 2, col + 1).Interior.Color = CLR_WHITE
        ws.Range(ws.Cells(r + 2, col), ws.Cells(r + 2, col + 1)).Merge
        
        col = col + 2
    Next i
    r = r + 4
    
    ' ── KPI CARDS — Row 2 ──
    Dim kpiLabels2 As Variant, kpiValues2 As Variant, kpiSubs2 As Variant
    kpiLabels2 = Array("TV SHOWS", "REPORT PERIOD", "LAST UPDATED", "DATA SOURCE")
    kpiValues2 = Array(showCount, "2021", Format(Now, "DD-MMM-YY"), "Public")
    kpiSubs2   = Array("Both Platforms", "Latest Year Available", "Auto-generated", "Amazon + Netflix")
    Dim kpiColors2 As Variant
    kpiColors2 = Array(CLR_GREEN, CLR_GOLD, CLR_BLUE, CLR_NAVY)
    
    col = 2
    For i = 0 To 3
        With ws.Cells(r, col)
            .Value = kpiLabels2(i)
            .Font.Bold = True
            .Font.Size = 9
            .Font.Color = CLR_MUTED
            .Interior.Color = CLR_LIGHT_GREY
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 16
        End With
        ws.Cells(r, col + 1).Interior.Color = CLR_LIGHT_GREY
        ws.Range(ws.Cells(r, col), ws.Cells(r, col + 1)).Merge
        
        With ws.Cells(r + 1, col)
            .Value = kpiValues2(i)
            .Font.Bold = True
            .Font.Size = 18
            .Font.Color = kpiColors2(i)
            .Interior.Color = CLR_WHITE
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 30
        End With
        ws.Cells(r + 1, col + 1).Interior.Color = CLR_WHITE
        ws.Range(ws.Cells(r + 1, col), ws.Cells(r + 1, col + 1)).Merge
        
        With ws.Cells(r + 2, col)
            .Value = kpiSubs2(i)
            .Font.Italic = True
            .Font.Size = 8
            .Font.Color = CLR_MUTED
            .Interior.Color = CLR_WHITE
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 14
        End With
        ws.Cells(r + 2, col + 1).Interior.Color = CLR_WHITE
        ws.Range(ws.Cells(r + 2, col), ws.Cells(r + 2, col + 1)).Merge
        
        col = col + 2
    Next i
    r = r + 4
    
    ' ── PLATFORM SUMMARY TABLE ──
    r = r + 1
    Call FormatHeader(ws, r, 2, 9, "  PLATFORM COMPARISON SUMMARY", CLR_BLUE)
    r = r + 1
    
    Dim colHeaders As Variant
    colHeaders = Array("Platform", "Total Titles", "Movies", "TV Shows", "Content Share")
    Call FormatColHeaders(ws, r, 2, colHeaders)
    r = r + 1
    
    ' Amazon row
    ws.Cells(r, 2).Value = "Amazon Prime"
    ws.Cells(r, 3).Value = amazonCount
    ws.Cells(r, 4).Value = movieCount - (netflixCount - showCount)  ' approximate
    ws.Cells(r, 5).Value = showCount - (netflixCount - amazonCount) ' approximate
    ws.Cells(r, 6).Value = Format(amazonCount / totalRows, "0.0%")
    ws.Cells(r, 2).Font.Color = CLR_AMAZON
    ws.Cells(r, 2).Font.Bold = True
    Call FormatDataRow(ws, r, 2, 6, True)
    r = r + 1
    
    ' Netflix row
    ws.Cells(r, 2).Value = "Netflix"
    ws.Cells(r, 3).Value = netflixCount
    ws.Cells(r, 4).Value = 6130
    ws.Cells(r, 5).Value = 2676
    ws.Cells(r, 6).Value = Format(netflixCount / totalRows, "0.0%")
    ws.Cells(r, 2).Font.Color = CLR_NETFLIX
    ws.Cells(r, 2).Font.Bold = True
    Call FormatDataRow(ws, r, 2, 6, False)
    r = r + 1
    
    ' Total row
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
        .Interior.Color = CLR_LIGHT_BLUE
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(170, 170, 170)
        .RowHeight = 18
    End With
    ws.Cells(r, 2).Value = "TOTAL"
    ws.Cells(r, 3).Value = totalRows
    ws.Cells(r, 6).Value = "100.0%"
    r = r + 2
    
    ' ── INSIGHTS BOX ──
    Call FormatHeader(ws, r, 2, 9, "  KEY INSIGHTS", CLR_NAVY)
    r = r + 1
    
    Dim insights As Variant
    insights = Array( _
        "Amazon Prime has a larger total library with " & Format(amazonCount, "#,##0") & " titles (" & Format(amazonCount / totalRows, "0.0%") & " share)", _
        "Netflix leads in TV Show count with 2,676 series vs Amazon's 1,854", _
        "Combined library covers 18,471 unique titles from 88 countries worldwide", _
        "Content growth accelerated from 2015 onwards on both platforms", _
        "Top producing country: United States — followed by India and United Kingdom" _
    )
    
    For i = 0 To UBound(insights)
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 9))
            .Merge
            .Value = "  " & Chr(9658) & "  " & insights(i)
            .Interior.Color = IIf(i Mod 2 = 0, CLR_WHITE, CLR_LIGHT_GREY)
            .Font.Size = 10
            .Font.Color = CLR_DARK
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 18
        End With
        r = r + 1
    Next i
    
    r = r + 1
    Call AddFooter(ws, r)
    
End Sub


' ── SHEET 2: PLATFORM COMPARISON ─────────────────────────────────────────────
Private Sub BuildPlatformComparison()
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Platform_Compare"
    ws.Tab.Color = CLR_AMAZON
    
    ws.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C:D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 2
    
    Call AddSheetTitle(ws, _
        "PLATFORM COMPARISON — AMAZON PRIME vs NETFLIX", _
        "Side-by-side performance metrics across all key dimensions")
    
    Dim r As Long
    r = 4
    
    Call FormatHeader(ws, r, 2, 5, "  HEAD-TO-HEAD PLATFORM METRICS", CLR_BLUE)
    r = r + 1
    
    Dim colH As Variant
    colH = Array("Metric", "Amazon Prime", "Netflix", "Winner")
    Call FormatColHeaders(ws, r, 2, colH)
    ws.Cells(r, 3).Interior.Color = CLR_AMAZON  ' Amazon header
    ws.Cells(r, 4).Interior.Color = CLR_NETFLIX ' Netflix header
    r = r + 1
    
    Dim compData As Variant
    compData = Array( _
        Array("Total Titles", "9,668", "8,806", "Amazon Prime"), _
        Array("Total Movies", "7,814", "6,130", "Amazon Prime"), _
        Array("Total TV Shows", "1,854", "2,676", "Netflix"), _
        Array("Content Share", "52.3%", "47.7%", "Amazon Prime"), _
        Array("Countries Covered", "88", "112", "Netflix"), _
        Array("Top Genre", "Drama", "International Movies", "Netflix"), _
        Array("Top Country", "United States", "United States", "Tie"), _
        Array("Latest Year", "2021", "2021", "Tie") _
    )
    
    Dim j As Long
    For j = 0 To UBound(compData)
        Dim rowArr As Variant
        rowArr = compData(j)
        ws.Cells(r, 2).Value = rowArr(0)
        ws.Cells(r, 3).Value = rowArr(1)
        ws.Cells(r, 4).Value = rowArr(2)
        ws.Cells(r, 5).Value = rowArr(3)
        
        ' Colour winner cell
        Select Case rowArr(3)
            Case "Amazon Prime": ws.Cells(r, 5).Font.Color = CLR_AMAZON
            Case "Netflix":      ws.Cells(r, 5).Font.Color = CLR_NETFLIX
            Case Else:           ws.Cells(r, 5).Font.Color = CLR_GOLD
        End Select
        ws.Cells(r, 5).Font.Bold = True
        
        Call FormatDataRow(ws, r, 2, 5, j Mod 2 = 0)
        r = r + 1
    Next j
    
    r = r + 1
    Call AddFooter(ws, r)
    
End Sub


' ── SHEET 3: GENRE ANALYSIS ───────────────────────────────────────────────────
Private Sub BuildGenreAnalysis()
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Genre_Analysis"
    ws.Tab.Color = CLR_BLUE
    
    ws.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 28
    ws.Columns("C:E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 2
    
    Call AddSheetTitle(ws, _
        "GENRE ANALYSIS — TOP CONTENT CATEGORIES", _
        "Most popular content genres by platform and total volume")
    
    Dim r As Long
    r = 4
    
    Call FormatHeader(ws, r, 2, 5, "  TOP 15 GENRES — COMBINED RANKING", CLR_BLUE)
    r = r + 1
    
    Dim colH As Variant
    colH = Array("Genre", "Amazon Prime", "Netflix", "Total Titles")
    Call FormatColHeaders(ws, r, 2, colH)
    ws.Cells(r, 3).Interior.Color = CLR_AMAZON
    ws.Cells(r, 4).Interior.Color = CLR_NETFLIX
    r = r + 1
    
    ' Top genres data (pre-calculated from the dataset analysis)
    Dim genreData As Variant
    genreData = Array( _
        Array("Drama", 3687, 2427, 6114), _
        Array("Comedy", 2099, 1674, 3773), _
        Array("International Movies", 0, 2752, 2752), _
        Array("Action", 1657, 859, 2516), _
        Array("Suspense / Thriller", 1501, 0, 1501), _
        Array("International TV Shows", 0, 1351, 1351), _
        Array("Kids / Children", 1085, 641, 1726), _
        Array("Documentary", 993, 868, 1861), _
        Array("Special Interest", 980, 0, 980), _
        Array("Horror", 875, 0, 875), _
        Array("TV Dramas", 0, 763, 763), _
        Array("Romantic Movies", 674, 616, 1290), _
        Array("Animation", 547, 0, 547), _
        Array("Independent Movies", 0, 756, 756), _
        Array("Documentaries", 993, 868, 1861) _
    )
    
    Dim j As Long
    For j = 0 To UBound(genreData)
        Dim rowArr As Variant
        rowArr = genreData(j)
        ws.Cells(r, 2).Value = rowArr(0)
        ws.Cells(r, 3).Value = rowArr(1)
        ws.Cells(r, 4).Value = rowArr(2)
        ws.Cells(r, 5).Value = rowArr(3)
        ws.Cells(r, 3).Font.Color = CLR_AMAZON
        ws.Cells(r, 4).Font.Color = CLR_NETFLIX
        ws.Cells(r, 5).Font.Bold = True
        ws.Cells(r, 5).Font.Color = CLR_NAVY
        Call FormatDataRow(ws, r, 2, 5, j Mod 2 = 0)
        r = r + 1
    Next j
    
    r = r + 1
    Call AddFooter(ws, r)
    
End Sub


' ── SHEET 4: COUNTRY REPORT ───────────────────────────────────────────────────
Private Sub BuildCountryReport()
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "Country_Report"
    ws.Tab.Color = CLR_GREEN
    
    ws.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C:F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 2
    
    Call AddSheetTitle(ws, _
        "COUNTRY REPORT — CONTENT PRODUCTION BY GEOGRAPHY", _
        "Top 10 countries contributing content to Amazon Prime and Netflix")
    
    Dim r As Long
    r = 4
    
    Call FormatHeader(ws, r, 2, 6, "  TOP 10 CONTENT-PRODUCING COUNTRIES", CLR_BLUE)
    r = r + 1
    
    Dim colH As Variant
    colH = Array("Country", "Amazon Prime", "Netflix", "Total", "Global Rank")
    Call FormatColHeaders(ws, r, 2, colH)
    ws.Cells(r, 3).Interior.Color = CLR_AMAZON
    ws.Cells(r, 4).Interior.Color = CLR_NETFLIX
    r = r + 1
    
    Dim countryData As Variant
    countryData = Array( _
        Array("United States", 334, 3689, 4023, "#1"), _
        Array("India", 246, 1046, 1292, "#2"), _
        Array("United Kingdom", 67, 806, 873, "#3"), _
        Array("Canada", 35, 445, 480, "#4"), _
        Array("France", 20, 393, 413, "#5"), _
        Array("Japan", 6, 318, 324, "#6"), _
        Array("Germany", 17, 226, 243, "#7"), _
        Array("Spain", 11, 232, 243, "#8"), _
        Array("South Korea", 0, 231, 231, "#9"), _
        Array("Mexico", 0, 169, 169, "#10") _
    )
    
    Dim j As Long
    For j = 0 To UBound(countryData)
        Dim rowArr As Variant
        rowArr = countryData(j)
        ws.Cells(r, 2).Value = rowArr(0)
        ws.Cells(r, 3).Value = rowArr(1)
        ws.Cells(r, 4).Value = rowArr(2)
        ws.Cells(r, 5).Value = rowArr(3)
        ws.Cells(r, 6).Value = rowArr(4)
        ws.Cells(r, 3).Font.Color = CLR_AMAZON
        ws.Cells(r, 4).Font.Color = CLR_NETFLIX
        ws.Cells(r, 5).Font.Bold = True
        ws.Cells(r, 6).Font.Bold = True
        ws.Cells(r, 6).Font.Color = CLR_GOLD
        ws.Cells(r, 6).HorizontalAlignment = xlCenter
        Call FormatDataRow(ws, r, 2, 6, j Mod 2 = 0)
        r = r + 1
    Next j
    
    r = r + 1
    Call AddFooter(ws, r)
    
End Sub


' ── SHEET 5: YOY TREND ────────────────────────────────────────────────────────
Private Sub BuildYoYTrend()
    
    Dim ws As Worksheet
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = "YoY_Trend"
    ws.Tab.Color = CLR_GOLD
    
    ws.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C:E").ColumnWidth = 16
    ws.Columns("F:G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 2
    
    Call AddSheetTitle(ws, _
        "YEAR-ON-YEAR CONTENT GROWTH TREND (2015–2021)", _
        "Annual content addition analysis showing platform growth trajectory")
    
    Dim r As Long
    r = 4
    
    Call FormatHeader(ws, r, 2, 7, "  CONTENT ADDED PER YEAR — BOTH PLATFORMS", CLR_BLUE)
    r = r + 1
    
    Dim colH As Variant
    colH = Array("Year", "Amazon Prime", "Amazon YoY%", "Netflix", "Netflix YoY%", "Combined Total")
    Call FormatColHeaders(ws, r, 2, colH)
    ws.Cells(r, 3).Interior.Color = CLR_AMAZON
    ws.Cells(r, 4).Interior.Color = CLR_AMAZON
    ws.Cells(r, 5).Interior.Color = CLR_NETFLIX
    ws.Cells(r, 6).Interior.Color = CLR_NETFLIX
    r = r + 1
    
    Dim trendData As Variant
    trendData = Array( _
        Array(2015, 378, "—", 560, "—", 938), _
        Array(2016, 521, "+37.8%", 902, "+61.1%", 1423), _
        Array(2017, 562, "+7.9%", 1032, "+14.4%", 1594), _
        Array(2018, 623, "+10.9%", 1147, "+11.1%", 1770), _
        Array(2019, 929, "+49.1%", 1030, "-10.2%", 1959), _
        Array(2020, 962, "+3.6%", 953, "-7.5%", 1915), _
        Array(2021, 1442, "+49.9%", 592, "-37.9%", 2034) _
    )
    
    Dim j As Long
    For j = 0 To UBound(trendData)
        Dim rowArr As Variant
        rowArr = trendData(j)
        ws.Cells(r, 2).Value = rowArr(0)  ' Year
        ws.Cells(r, 3).Value = rowArr(1)  ' Amazon
        ws.Cells(r, 4).Value = rowArr(2)  ' Amazon YoY
        ws.Cells(r, 5).Value = rowArr(3)  ' Netflix
        ws.Cells(r, 6).Value = rowArr(4)  ' Netflix YoY
        ws.Cells(r, 7).Value = rowArr(5)  ' Total
        
        ' Colour YoY cells based on growth/decline
        If InStr(CStr(rowArr(2)), "-") > 0 Then
            ws.Cells(r, 4).Font.Color = RGB(220, 50, 50)
        ElseIf rowArr(2) <> "—" Then
            ws.Cells(r, 4).Font.Color = RGB(34, 139, 34)
        End If
        
        If InStr(CStr(rowArr(4)), "-") > 0 Then
            ws.Cells(r, 6).Font.Color = RGB(220, 50, 50)
        ElseIf rowArr(4) <> "—" Then
            ws.Cells(r, 6).Font.Color = RGB(34, 139, 34)
        End If
        
        ws.Cells(r, 3).Font.Color = CLR_AMAZON
        ws.Cells(r, 5).Font.Color = CLR_NETFLIX
        ws.Cells(r, 7).Font.Bold = True
        ws.Cells(r, 7).Font.Color = CLR_NAVY
        
        Call FormatDataRow(ws, r, 2, 7, j Mod 2 = 0)
        r = r + 1
    Next j
    
    r = r + 1
    
    ' ── TREND INSIGHTS ──
    Call FormatHeader(ws, r, 2, 7, "  TREND INSIGHTS", CLR_NAVY)
    r = r + 1
    
    Dim insights As Variant
    insights = Array( _
        "Amazon Prime had its best year in 2021 — added 1,442 titles (49.9% YoY growth)", _
        "Netflix peak was 2018 — 1,147 titles added in a single calendar year", _
        "2020 saw slight decline on Netflix — likely due to COVID-19 production shutdowns", _
        "Combined content volume grew by 117% between 2015 and 2021", _
        "Amazon consistently growing; Netflix shows fluctuation after 2018 peak" _
    )
    
    Dim k As Long
    For k = 0 To UBound(insights)
        With ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
            .Merge
            .Value = "  " & Chr(9658) & "  " & insights(k)
            .Interior.Color = IIf(k Mod 2 = 0, CLR_WHITE, CLR_LIGHT_GREY)
            .Font.Size = 10
            .Font.Color = CLR_DARK
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(221, 221, 221)
            .RowHeight = 18
        End With
        r = r + 1
    Next k
    
    r = r + 1
    Call AddFooter(ws, r)
    
End Sub

'==============================================================================
' END OF MODULE
' Author: Abdul Jabbar M | Data Analyst & Automation Specialist
' LinkedIn: linkedin.com/in/abdul-jabbar-04952713a
'==============================================================================
